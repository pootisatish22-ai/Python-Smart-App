"""
core.py - the Intelligence Plugin engine. No Streamlit code in here, so the web app (app.py)
and the API (api.py) both use exactly the same logic.

Flow:  PDF -> validate -> read text + tables -> classify -> extract key/value pairs and item tables
       -> validate (missing, formats, totals) -> structured result (JSON / CSV / Excel / Word ...)
"""
import functools
import hashlib
import io
import json
import os
import re
import zipfile
from typing import List, Literal, Optional

try:
    import pymupdf as fitz  # PyMuPDF (fast). If it fails to load on Windows, pdfplumber is used instead.
except Exception:  # noqa
    fitz = None
import pandas as pd
import pdfplumber
from dateutil import parser as dateparser
from docx import Document
from dotenv import load_dotenv
from pydantic import BaseModel
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

load_dotenv()

MAX_MB = 20
MAX_PAGES = 60
MAX_CHARS = 60000
TIMEOUT = int(os.getenv("LLM_TIMEOUT", "60"))  # seconds per AI call; errors out instead of hanging

# ---------------------------------------------------------------- Document types (extensible)
# "line_items" is not a text field: it is satisfied when an item table is found.
EXPECTED = {
    "invoice": ["invoice_number", "invoice_date", "due_date", "vendor_name", "customer_name",
                "currency", "subtotal", "tax", "total_amount", "line_items"],
    "receipt": ["merchant_name", "date", "tax", "total_amount", "payment_method", "line_items"],
    "purchase_order": ["po_number", "po_date", "supplier_name", "buyer_name", "delivery_date",
                       "total_amount", "line_items"],
    "resume": ["full_name", "email", "phone", "location", "summary", "skills", "experience",
               "education"],
    "bank_statement": ["account_holder", "account_number", "bank_name", "statement_period",
                       "opening_balance", "closing_balance"],
    "contract": ["parties", "effective_date", "term", "payment_terms", "termination_clause",
                 "governing_law", "signatures"],
    "report": ["title", "author", "date", "executive_summary", "conclusion"],
    "generic": ["title", "date", "author", "subject"],
}
DocType = Literal["invoice", "receipt", "purchase_order", "resume", "bank_statement",
                  "contract", "report", "generic"]


class Classification(BaseModel):
    doc_type: DocType
    reason: str


class FieldItem(BaseModel):
    name: str
    value: Optional[str] = None
    confidence: float = 0.0
    page: Optional[int] = None


class ExtractedTable(BaseModel):
    name: str                 # e.g. line_items, transactions
    columns: List[str]
    rows: List[List[str]]     # every cell as written in the document


class Extraction(BaseModel):
    title: Optional[str] = None
    summary: str = ""
    fields: List[FieldItem]
    tables: List[ExtractedTable] = []
    insights: List[str] = []


class Job(BaseModel):
    title: Optional[str] = None
    company: Optional[str] = None
    location: Optional[str] = None
    start: Optional[str] = None
    end: Optional[str] = None
    bullets: List[str] = []


class Education(BaseModel):
    degree: Optional[str] = None
    institution: Optional[str] = None
    year: Optional[str] = None


class ResumeProfile(BaseModel):
    full_name: Optional[str] = None
    headline: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    location: Optional[str] = None
    links: List[str] = []
    summary: Optional[str] = None
    skills: List[str] = []
    experience: List[Job] = []
    education: List[Education] = []
    certifications: List[str] = []


# ---------------------------------------------------------------- LLM layer (provider neutral)
@functools.lru_cache(maxsize=1)
def get_llm():
    provider = os.getenv("LLM_PROVIDER", "gemini").lower()
    if provider == "claude":
        from langchain_anthropic import ChatAnthropic
        kwargs = {}
        if os.getenv("ANTHROPIC_BASE_URL"):
            kwargs["base_url"] = os.getenv("ANTHROPIC_BASE_URL")
        return ChatAnthropic(model=os.getenv("CLAUDE_MODEL", "claude-sonnet-5-5"),
                             temperature=0, timeout=TIMEOUT, max_retries=2, **kwargs)
    if provider == "groq":
        from langchain_groq import ChatGroq
        return ChatGroq(model=os.getenv("GROQ_MODEL", "openai/gpt-oss-120b"),
                        temperature=0, timeout=TIMEOUT, max_retries=2)
    if provider == "openai":  # any OpenAI-compatible service: OpenRouter, a company gateway, etc.
        from langchain_openai import ChatOpenAI
        return ChatOpenAI(model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
                          base_url=os.getenv("OPENAI_BASE_URL") or None,
                          temperature=0, timeout=TIMEOUT, max_retries=2)
    if provider == "ollama":
        from langchain_ollama import ChatOllama
        return ChatOllama(model=os.getenv("OLLAMA_MODEL", "llama3.2:3b"), temperature=0)
    from langchain_google_genai import ChatGoogleGenerativeAI
    return ChatGoogleGenerativeAI(model=os.getenv("GEMINI_MODEL", "gemini-3.8-flash"),
                                  temperature=0, timeout=TIMEOUT, max_retries=2)


def llm_ready() -> Optional[str]:
    """Return an error message if the chosen provider has no key, else None."""
    p = os.getenv("LLM_PROVIDER", "gemini").lower()
    need = {"gemini": "GOOGLE_API_KEY", "groq": "GROQ_API_KEY", "openai": "OPENAI_API_KEY",
            "claude": "ANTHROPIC_API_KEY"}.get(p)
    if need and not os.getenv(need):
        return f"{need} is missing. Add it to the .env file and restart."
    return None


def ask_structured(schema, prompt):
    chain = get_llm().with_structured_output(schema)
    last = None
    for _ in range(2):  # one retry
        try:
            return chain.invoke(prompt)
        except Exception as e:  # noqa
            last = e
    raise last


def text_of(resp) -> str:
    c = resp.content
    if isinstance(c, list):
        return "".join(p.get("text", "") if isinstance(p, dict) else str(p) for p in c)
    return str(c)


# ---------------------------------------------------------------- Step 1: intake and validation
def read_pages(data: bytes) -> List[str]:
    """Text of every page. Uses PyMuPDF if it loads, otherwise pdfplumber."""
    if fitz is not None:
        doc = fitz.open(stream=data, filetype="pdf")
        if doc.needs_pass:
            raise ValueError("This PDF is password protected. Remove the password and try again.")
        return [p.get_text() for p in doc]
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return [(p.extract_text() or "") for p in pdf.pages]
    except Exception as e:  # noqa
        if "password" in str(e).lower() or "encrypt" in str(e).lower():
            raise ValueError("This PDF is password protected. Remove the password and try again.")
        raise ValueError(f"Could not read this PDF: {e}")


def validate_pdf(data: bytes, name: str):
    if not name.lower().endswith(".pdf"):
        raise ValueError("Only PDF files are supported.")
    if len(data) > MAX_MB * 1024 * 1024:
        raise ValueError(f"File is larger than {MAX_MB} MB.")
    if data[:5] != b"%PDF-":
        raise ValueError("This file is not a valid PDF (signature check failed).")
    pages = read_pages(data)
    if len(pages) > MAX_PAGES:
        raise ValueError(f"PDF has more than {MAX_PAGES} pages.")
    meta = {
        "document_id": hashlib.sha256(data).hexdigest()[:12],
        "file_name": name,
        "size_kb": round(len(data) / 1024, 1),
        "pages": len(pages),
        "sha256": hashlib.sha256(data).hexdigest(),
    }
    return pages, meta


# ---------------------------------------------------------------- Step 2: reading tables and values
def extract_tables(data: bytes):
    """Tables found in the PDF layout (raw, no AI)."""
    out = []
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for pno, page in enumerate(pdf.pages[:MAX_PAGES], 1):
                for t in page.extract_tables():
                    if not t or len(t) < 2:
                        continue
                    header = [str(h or f"col{i + 1}") for i, h in enumerate(t[0])]
                    if len(set(header)) != len(header):
                        header = [f"col{i + 1}" for i in range(len(header))]
                    out.append((pno, pd.DataFrame(t[1:], columns=header)))
    except Exception:
        pass
    return out


EMPTY = {"", "null", "none", "n/a", "na", "not found", "not available", "not specified", "unknown"}


def is_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip().lower() in EMPTY


def parse_amount(text) -> Optional[float]:
    """'Rs 1,155.00' -> 1155.0 ; '$2,865.50' -> 2865.5 ; None if there is no number."""
    if text is None:
        return None
    s = re.sub(r"[^0-9,.\-]", "", str(text))
    if not re.search(r"\d", s):
        return None
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", "") if re.search(r",\d{3}(\D|$)", s) else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def check_value(name, value, conf) -> str:
    """Normalisation and validation: dates, amounts, confidence."""
    if is_empty(value):
        return "MISSING"
    n = str(name).lower()
    try:
        if "date" in n:
            dateparser.parse(str(value), dayfirst=True)
        elif any(k in n for k in ("total", "amount", "subtotal", "tax", "balance", "price")):
            if parse_amount(value) is None:
                raise ValueError
    except Exception:
        return "Check format"
    if conf is not None and not pd.isna(conf) and float(conf) < 0.6:
        return "Low confidence"
    return "OK"


def refresh_status(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Status"] = [check_value(n, v, c) for n, v, c in
                    zip(df["Field"], df["Value"], df["Confidence"])]
    return df


def _unique(cols):
    seen, out = {}, []
    for i, c in enumerate(cols):
        c = str(c).strip() or f"col{i + 1}"
        seen[c] = seen.get(c, 0) + 1
        out.append(c if seen[c] == 1 else f"{c}_{seen[c]}")
    return out


def tables_from_ai(ext_tables) -> list:
    out = []
    for t in ext_tables or []:
        cols = _unique(t.columns)
        if not cols:
            continue
        rows = [(list(r) + [""] * len(cols))[:len(cols)] for r in t.rows]
        if rows:
            out.append({"name": t.name or "table", "df": pd.DataFrame(rows, columns=cols)})
    return out


def reconcile(df: pd.DataFrame, tables: list) -> List[str]:
    """Do the item amounts add up to the stated subtotal (or total)? Returns warnings."""
    vals = {str(r.Field).lower(): r.Value for r in df.itertuples()}
    target = parse_amount(vals.get("subtotal"))
    label = "subtotal"
    if target is None and is_empty(vals.get("tax")):
        target, label = parse_amount(vals.get("total_amount")), "total"
    if target is None:
        return []
    for t in tables:
        for c in t["df"].columns:
            if re.search(r"amount|total", str(c), re.I) and not re.search(r"unit", str(c), re.I):
                nums = [parse_amount(x) for x in t["df"][c]]
                nums = [n for n in nums if n is not None]
                if len(nums) >= 2:
                    s = round(sum(nums), 2)
                    if abs(s - target) > max(0.5, 0.005 * abs(target)):
                        return [f"Item amounts add up to {s:,.2f} but the stated {label} is {target:,.2f}. Please check."]
                    return []
    return []


def build_chunks(pages: List[str]):
    chunks = []
    for pno, text in enumerate(pages, 1):
        for i in range(0, len(text), 1200):
            c = text[i:i + 1200].strip()
            if c:
                chunks.append((pno, c))
    return chunks


# ---------------------------------------------------------------- Step 3: the processing pipeline
def process(data: bytes, name: str) -> dict:
    pages, meta = validate_pdf(data, name)
    if sum(len(p.strip()) for p in pages) < 50 * max(1, len(pages)):
        raise ValueError("No readable text found. This looks like a scanned PDF "
                         "(OCR is in the roadmap). Please use a text-based PDF.")
    raw_tables = extract_tables(data)
    full_text = "\n".join(f"[Page {i}]\n{t}" for i, t in enumerate(pages, 1))[:MAX_CHARS]

    cls = ask_structured(
        Classification,
        "Classify this document. Choose the closest type.\n\n" + full_text[:4000])
    doc_type = cls.doc_type
    expected = EXPECTED.get(doc_type, EXPECTED["generic"])
    scalar_expected = [e for e in expected if e != "line_items"]

    ext = ask_structured(
        Extraction,
        f"You are a document extraction engine. Document type: {doc_type}.\n"
        f"Return: title; summary (3-4 sentences); fields; tables; insights (3-5 key points).\n"
        f"FIELDS: return exactly one entry for EACH expected field: {scalar_expected}. Then add EVERY other "
        f"relevant key/value pair you find (up to 30), for example addresses, tax ids, reference numbers, "
        f"dates, totals, contact details. For each field give: name (snake_case), value as text, "
        f"confidence between 0 and 1, and the page number where it was found.\n"
        f"TABLES: if the document has item-level rows (invoice or receipt line items, purchase order items, "
        f"bank transactions), return each as a table with a name (line_items or transactions), column names "
        f"and rows (at most 100). Every cell is text exactly as written. Do not invent rows.\n"
        f"If the document does not state a value, set value to null. NEVER guess.\n"
        f"Text is marked with [Page N].\n\nDOCUMENT:\n{full_text}")

    tables = tables_from_ai(ext.tables)
    if not tables and raw_tables:  # fall back to what the PDF layout shows
        tables = [{"name": f"table_p{p}_{i}", "df": t} for i, (p, t) in enumerate(raw_tables, 1)]
    item_tbl = next((t for t in tables if "item" in t["name"].lower()), None)
    if item_tbl is None and doc_type in ("invoice", "receipt", "purchase_order") and tables:
        item_tbl = tables[0]

    found = {f.name.strip().lower(): f for f in ext.fields}
    rows = []
    for name_ in expected:  # expected fields first, so missing ones always appear
        if name_ == "line_items":
            rows.append([name_, f"{len(item_tbl['df'])} rows" if item_tbl is not None else None,
                         None, 0.9 if item_tbl is not None else 0.0])
            continue
        f = found.pop(name_.lower(), None)
        rows.append([name_, None if f is None or is_empty(f.value) else f.value,
                     f.page if f else None, f.confidence if f else 0.0])
    for f in found.values():  # every other relevant key/value pair
        rows.append([f.name, None if is_empty(f.value) else f.value, f.page, f.confidence])
    df = pd.DataFrame(rows, columns=["Field", "Value", "Page", "Confidence"])
    df["Value"] = df["Value"].fillna("").astype(str)
    df["Page"] = pd.to_numeric(df["Page"], errors="coerce").astype("Int64")
    df = refresh_status(df)

    profile = None
    if doc_type == "resume":  # document-specific extractor: a full profile for populating a CV
        profile = ask_structured(
            ResumeProfile,
            "Extract the candidate profile from this resume. Keep each job's bullets short. "
            "Use null or an empty list for anything not stated. NEVER invent details.\n\n" + full_text).model_dump()

    chunks = build_chunks(pages)
    vec = mat = None
    try:
        vec = TfidfVectorizer(stop_words="english").fit([c for _, c in chunks])
        mat = vec.transform([c for _, c in chunks])
    except Exception:
        pass

    return {"id": meta["document_id"], "meta": meta, "doc_type": doc_type, "reason": cls.reason,
            "title": ext.title or name, "summary": ext.summary, "insights": ext.insights,
            "pdf": data, "df": df, "tables": tables, "raw_tables": raw_tables, "profile": profile,
            "warnings": reconcile(df, tables), "pages": pages, "chunks": chunks, "vec": vec, "mat": mat}


# ---------------------------------------------------------------- Chat with citations (simple RAG)
def retrieve(d: dict, question: str, k: int = 4):
    if d["vec"] is None:
        return d["chunks"][:k]
    sims = cosine_similarity(d["vec"].transform([question]), d["mat"])[0]
    top = sims.argsort()[::-1][:k]
    return [d["chunks"][i] for i in sorted(top)]


def answer(d: dict, question: str, history: list) -> str:
    ctx = "\n\n".join(f"[Page {p}] {c}" for p, c in retrieve(d, question))
    hist = "\n".join(f"{m['role']}: {m['content']}" for m in history[-6:])
    prompt = ("Answer ONLY from the context. Cite pages like (p.2). "
              "If the answer is not in the context, say: Not found in the document.\n\n"
              f"CONTEXT:\n{ctx}\n\nCHAT SO FAR:\n{hist}\n\nQUESTION: {question}")
    return text_of(get_llm().invoke(prompt))


# ---------------------------------------------------------------- Structured output for other systems
def blank_str(df: pd.DataFrame) -> pd.DataFrame:
    return df.astype(object).where(df.notna(), "").astype(str)


def edited_result(d: dict, df: pd.DataFrame) -> dict:
    """The clean structured result: this is what the JSON export, the API and every publisher use."""
    df = refresh_status(df)
    fields = blank_str(df).to_dict(orient="records")
    return {
        "document": d["meta"], "document_type": d["doc_type"], "title": d["title"],
        "summary": d["summary"], "insights": d["insights"],
        "fields": fields,
        "key_values": {f["Field"]: f["Value"] for f in fields if f["Value"]},
        "missing": df.loc[df["Status"] == "MISSING", "Field"].tolist(),
        "warnings": d.get("warnings", []),
        "tables": [{"name": t["name"], "columns": list(t["df"].columns),
                    "rows": blank_str(t["df"]).to_dict(orient="records")} for t in d["tables"]],
        "profile": d.get("profile"),
    }


def to_excel(d: dict, df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([{"Title": d["title"], "Type": d["doc_type"], "Summary": d["summary"],
                       "File": d["meta"]["file_name"], "Pages": d["meta"]["pages"]}]
                     ).to_excel(w, sheet_name="Overview", index=False)
        df.to_excel(w, sheet_name="Fields", index=False)
        pd.DataFrame({"Missing fields": df.loc[df["Status"] == "MISSING", "Field"]}
                     ).to_excel(w, sheet_name="Missing", index=False)
        for i, t in enumerate(d["tables"][:20], 1):
            t["df"].to_excel(w, sheet_name=re.sub(r"[\[\]:*?/\\]", "_", f"{i}_{t['name']}")[:31], index=False)
    return buf.getvalue()


def to_word(d: dict, df: pd.DataFrame) -> bytes:
    doc = Document()
    doc.add_heading(str(d["title"]), 0)
    doc.add_paragraph(f"Document type: {d['doc_type']}   |   Pages: {d['meta']['pages']}")
    doc.add_heading("Summary", 1)
    doc.add_paragraph(d["summary"])
    doc.add_heading("Key fields", 1)
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    for c, h in zip(t.rows[0].cells, ["Field", "Value", "Status"]):
        c.text = h
    for _, r in df.iterrows():
        cells = t.add_row().cells
        cells[0].text, cells[1].text, cells[2].text = str(r["Field"]), str(r["Value"]), str(r["Status"])
    for tb in d["tables"][:5]:
        doc.add_heading(f"Table: {tb['name']}", 1)
        wt = doc.add_table(rows=1, cols=len(tb["df"].columns))
        wt.style = "Table Grid"
        for c, h in zip(wt.rows[0].cells, tb["df"].columns):
            c.text = str(h)
        for _, r in tb["df"].head(60).iterrows():
            for c, v in zip(wt.add_row().cells, r):
                c.text = str(v)
    doc.add_heading("Missing information", 1)
    missing = df.loc[df["Status"] == "MISSING", "Field"].tolist()
    doc.add_paragraph(", ".join(missing) if missing else "Nothing is missing.")
    doc.add_heading("Key insights", 1)
    for i in d["insights"]:
        doc.add_paragraph(i, style="List Bullet")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def to_markdown(d: dict, df: pd.DataFrame) -> str:
    missing = df.loc[df["Status"] == "MISSING", "Field"].tolist()
    lines = [f"# {d['title']}", f"**Type:** {d['doc_type']}", "", "## Summary", d["summary"], "",
             "## Key fields"]
    lines += [f"- **{r.Field}:** {r.Value or '-'} ({r.Status})" for r in df.itertuples()]
    lines += ["", "## Missing", ", ".join(missing) or "Nothing missing", "", "## Insights"]
    lines += [f"- {i}" for i in d["insights"]]
    return "\n".join(lines)


def table_csvs(d: dict) -> dict:
    return {f"table_{re.sub(r'[^A-Za-z0-9_-]', '_', t['name'])}.csv": t["df"].to_csv(index=False)
            for t in d["tables"]}


def to_zip(d: dict, df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("result.json", json.dumps(edited_result(d, df), indent=2))
        z.writestr("fields.csv", df.to_csv(index=False))
        for n, c in table_csvs(d).items():
            z.writestr(n, c)
        z.writestr("result.xlsx", to_excel(d, df))
        z.writestr("report.docx", to_word(d, df))
        z.writestr("summary.md", to_markdown(d, df))
        z.writestr("extracted_text.txt", "\n\n".join(
            f"--- Page {i} ---\n{t}" for i, t in enumerate(d["pages"], 1)))
    return buf.getvalue()
