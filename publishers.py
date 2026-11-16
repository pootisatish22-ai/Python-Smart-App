"""
publishers.py - sends APPROVED document data to downstream tools.
No Streamlit code in here, so it can be reused later behind an API.
Targets: local/shared folder, Excel register, JIRA, ServiceNow. Plus an audit log.
"""
import datetime
import json
import os

import pandas as pd
import requests

TIMEOUT = 30


def publish_dir() -> str:
    d = os.getenv("PUBLISH_DIR", "published")
    os.makedirs(d, exist_ok=True)
    return d


def now() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def audit(event: str, data: dict) -> None:
    line = {"time": now(), "event": event, **data}
    with open(os.path.join(publish_dir(), "audit_log.jsonl"), "a", encoding="utf-8") as f:
        f.write(json.dumps(line, default=str) + "\n")


def read_audit(last: int = 15) -> list:
    path = os.path.join(publish_dir(), "audit_log.jsonl")
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return [json.loads(x) for x in f.read().splitlines()[-last:]][::-1]


def configured() -> dict:
    return {
        "folder": True,
        "excel": True,
        "concur": True,  # demo mode works offline; live mode needs the CONCUR_* settings
        "jira": bool(os.getenv("JIRA_URL") and os.getenv("JIRA_PROJECT_KEY") and
                     (os.getenv("JIRA_PAT") or (os.getenv("JIRA_EMAIL") and os.getenv("JIRA_API_TOKEN")))),
        "servicenow": bool(os.getenv("SNOW_INSTANCE") and os.getenv("SNOW_USER")
                           and os.getenv("SNOW_PASSWORD")),
    }


# ------------------------------------------------------------------ shared text
def ticket_title(result: dict) -> str:
    return f"[{result['document_type']}] {result['title']}"[:240]


def ticket_text(result: dict) -> str:
    lines = [
        f"Document: {result['document']['file_name']}  ({result['document_type']})",
        f"Document ID: {result['document']['document_id']}",
        f"Approved by: {result.get('approved_by', '-')} at {result.get('approved_at', '-')}",
        "",
        "Summary:", result["summary"], "", "Extracted fields:",
    ]
    for f in result["fields"]:
        lines.append(f"- {f['Field']}: {f['Value'] or '-'}  [{f['Status']}]")
    for t in result.get("tables", [])[:1]:
        lines += ["", f"Table {t['name']} (first 10 rows):"]
        lines += ["- " + " | ".join(str(v) for v in r.values()) for r in t["rows"][:10]]
    lines += ["", "Missing: " + (", ".join(result["missing"]) or "none"), "", "Key insights:"]
    lines += [f"- {i}" for i in result["insights"]]
    lines += ["", "Created automatically by the Intelligence Plugin."]
    return "\n".join(lines)[:30000]


# ------------------------------------------------------------------ 1. folder
def publish_folder(result: dict, files: dict) -> str:
    base = os.path.splitext(result["document"]["file_name"])[0]
    folder = os.path.join(publish_dir(), f"{result['document']['document_id']}_{base}")
    os.makedirs(folder, exist_ok=True)
    for name, content in files.items():
        mode = "wb" if isinstance(content, bytes) else "w"
        with open(os.path.join(folder, name), mode, **({} if mode == "wb" else {"encoding": "utf-8"})) as f:
            f.write(content)
    return f"Saved {len(files)} files to {os.path.abspath(folder)}"


# ------------------------------------------------------------------ 2. Excel register
def publish_excel(result: dict) -> str:
    path = os.path.join(publish_dir(), "document_register.xlsx")
    sheets = pd.read_excel(path, sheet_name=None, dtype=str) if os.path.exists(path) else {}
    reg_row = pd.DataFrame([{
        "Published at": now(), "Document ID": result["document"]["document_id"],
        "File": result["document"]["file_name"], "Type": result["document_type"],
        "Title": result["title"], "Missing count": len(result["missing"]),
        "Missing fields": ", ".join(result["missing"]), "Approved by": result.get("approved_by", ""),
    }])
    sheets["Register"] = pd.concat([sheets["Register"], reg_row], ignore_index=True) \
        if "Register" in sheets else reg_row
    typ = result["document_type"][:31]
    row = pd.DataFrame([{"Document ID": result["document"]["document_id"],
                         "File": result["document"]["file_name"],
                         **{f["Field"]: f["Value"] for f in result["fields"]}}])
    sheets[typ] = pd.concat([sheets[typ], row], ignore_index=True) if typ in sheets else row
    tbl = next((t for t in result.get("tables", []) if "item" in t["name"].lower()),
               (result.get("tables") or [None])[0])
    if tbl and tbl["rows"]:  # item-level rows from every document go into one "Line items" sheet
        items = pd.DataFrame([{"Document ID": result["document"]["document_id"], "File": result["document"]["file_name"],
                               "Table": tbl["name"], **r} for r in tbl["rows"]])
        sheets["Line items"] = pd.concat([sheets["Line items"], items], ignore_index=True) \
            if "Line items" in sheets else items
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            for name, df in sheets.items():
                df.to_excel(w, sheet_name=name[:31], index=False)
    except PermissionError:
        raise RuntimeError("Close document_register.xlsx in Excel and try again.")
    return f"Added a row to {os.path.abspath(path)} (sheets: Register, {typ})"


# ------------------------------------------------------------------ 3. JIRA
def _jira_auth():
    if os.getenv("JIRA_PAT"):  # Jira Server / Data Center personal access token
        return None, {"Authorization": f"Bearer {os.getenv('JIRA_PAT')}"}
    return (os.getenv("JIRA_EMAIL"), os.getenv("JIRA_API_TOKEN")), {}  # Jira Cloud


def jira_payload(result: dict) -> dict:
    return {"fields": {
        "project": {"key": os.getenv("JIRA_PROJECT_KEY", "DEMO")},
        "summary": ticket_title(result),
        "description": ticket_text(result),
        "issuetype": {"name": os.getenv("JIRA_ISSUE_TYPE", "Task")},
        "labels": ["document-intelligence", result["document_type"].replace(" ", "_")],
    }}


def publish_jira(result: dict, attach_name: str, attach_bytes: bytes) -> str:
    base = os.getenv("JIRA_URL").rstrip("/")
    auth, headers = _jira_auth()
    r = requests.post(f"{base}/rest/api/2/issue", json=jira_payload(result), auth=auth,
                      headers=headers, timeout=TIMEOUT)
    if r.status_code >= 300:
        raise RuntimeError(f"JIRA said {r.status_code}: {r.text[:300]}")
    key = r.json()["key"]
    a = requests.post(f"{base}/rest/api/2/issue/{key}/attachments",
                      files={"file": (attach_name, attach_bytes, "application/json")},
                      auth=auth, headers={**headers, "X-Atlassian-Token": "no-check"},
                      timeout=TIMEOUT)
    note = "with attachment" if a.status_code < 300 else f"(attachment failed: {a.status_code})"
    return f"Created JIRA issue {key} {note}: {base}/browse/{key}"


# ------------------------------------------------------------------ 4. ServiceNow
def snow_payload(result: dict) -> dict:
    return {"short_description": ticket_title(result), "description": ticket_text(result)}


def publish_servicenow(result: dict, attach_name: str, attach_bytes: bytes) -> str:
    base = os.getenv("SNOW_INSTANCE").rstrip("/")
    table = os.getenv("SNOW_TABLE", "incident")
    auth = (os.getenv("SNOW_USER"), os.getenv("SNOW_PASSWORD"))
    r = requests.post(f"{base}/api/now/table/{table}", json=snow_payload(result), auth=auth,
                      headers={"Accept": "application/json"}, timeout=TIMEOUT)
    if r.status_code >= 300:
        raise RuntimeError(f"ServiceNow said {r.status_code}: {r.text[:300]}")
    rec = r.json()["result"]
    a = requests.post(f"{base}/api/now/attachment/file",
                      params={"table_name": table, "table_sys_id": rec["sys_id"],
                              "file_name": attach_name},
                      data=attach_bytes, auth=auth,
                      headers={"Content-Type": "application/json", "Accept": "application/json"},
                      timeout=TIMEOUT)
    note = "with attachment" if a.status_code < 300 else f"(attachment failed: {a.status_code})"
    return f"Created ServiceNow {table} {rec.get('number', rec['sys_id'])} {note}"
