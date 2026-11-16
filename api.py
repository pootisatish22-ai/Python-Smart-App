"""
api.py - lets OTHER systems use the plugin: send a PDF, get clean structured data back.

Run:   uvicorn api:app --port 8000        then open  http://localhost:8000/docs
Try:   curl -F "file=@invoice.pdf" http://localhost:8000/process
Optional protection: set PLUGIN_API_KEY in .env and send the header  X-API-Key: <value>
"""
import os
from typing import Dict, Optional

from fastapi import Depends, FastAPI, File, Header, HTTPException, UploadFile
from fastapi.responses import Response

import concur
import core
import publishers as pub
import resume_builder

app = FastAPI(title="Intelligence Plugin API", version="2.0",
              description="Give a PDF, get structured data (JSON / CSV), publish it to connected tools.")
STORE: Dict[str, dict] = {}  # processed documents, kept in memory for this prototype


def auth(x_api_key: Optional[str] = Header(default=None)):
    key = os.getenv("PLUGIN_API_KEY")
    if key and x_api_key != key:
        raise HTTPException(401, "Invalid or missing X-API-Key header")


def get_doc(doc_id: str) -> dict:
    if doc_id not in STORE:
        raise HTTPException(404, "Unknown document id. POST the PDF to /process first.")
    return STORE[doc_id]


@app.get("/health")
def health():
    return {"status": "ok", "provider": os.getenv("LLM_PROVIDER", "gemini"), "ai_ready": core.llm_ready() is None,
            "concur_mode": concur.mode()}


@app.post("/process", dependencies=[Depends(auth)])
def process_pdf(file: UploadFile = File(...)):
    """Detect the document type, extract key/value pairs and item tables, validate. Returns the structured result."""
    try:
        d = core.process(file.file.read(), file.filename or "upload.pdf")
    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:  # noqa
        raise HTTPException(502, f"AI processing failed: {e}")
    STORE[d["id"]] = d
    return core.edited_result(d, d["df"])


@app.get("/documents/{doc_id}", dependencies=[Depends(auth)])
def document_json(doc_id: str):
    d = get_doc(doc_id)
    return core.edited_result(d, d["df"])


@app.get("/documents/{doc_id}/fields.csv", dependencies=[Depends(auth)])
def fields_csv(doc_id: str):
    return Response(get_doc(doc_id)["df"].to_csv(index=False), media_type="text/csv")


@app.get("/documents/{doc_id}/tables/{index}.csv", dependencies=[Depends(auth)])
def table_csv(doc_id: str, index: int):
    tables = get_doc(doc_id)["tables"]
    if not 0 <= index < len(tables):
        raise HTTPException(404, f"This document has {len(tables)} table(s); use an index from 0.")
    return Response(tables[index]["df"].to_csv(index=False), media_type="text/csv")


@app.get("/documents/{doc_id}/resume.docx", dependencies=[Depends(auth)])
def resume_docx(doc_id: str):
    d = get_doc(doc_id)
    if not d.get("profile"):
        raise HTTPException(400, "This document is not a resume.")
    return Response(resume_builder.build_resume(d["profile"]), media_type=
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.post("/documents/{doc_id}/publish/{target}", dependencies=[Depends(auth)])
def publish(doc_id: str, target: str):
    """target: folder | excel | concur | jira | servicenow"""
    d = get_doc(doc_id)
    result = core.edited_result(d, d["df"])
    result["approved_by"], result["approved_at"] = "API", pub.now()
    result_json = __import__("json").dumps(result, indent=2)
    try:
        if target == "folder":
            msg = pub.publish_folder(result, {"result.json": result_json, "fields.csv": d["df"].to_csv(index=False),
                                              "original.pdf": d["pdf"], **core.table_csvs(d)})
        elif target == "excel":
            msg = pub.publish_excel(result)
        elif target == "concur":
            msg = concur.submit(result, d["meta"]["file_name"], d["pdf"])
        elif target == "jira":
            msg = pub.publish_jira(result, "result.json", result_json.encode())
        elif target == "servicenow":
            msg = pub.publish_servicenow(result, "result.json", result_json.encode())
        else:
            raise HTTPException(400, "target must be folder, excel, concur, jira or servicenow")
    except HTTPException:
        raise
    except Exception as e:  # noqa
        pub.audit("publish_failed", {"document_id": doc_id, "target": target, "error": str(e), "via": "api"})
        raise HTTPException(502, str(e))
    pub.audit("published", {"document_id": doc_id, "target": target, "result": msg, "via": "api"})
    return {"target": target, "result": msg}
