"""
concur.py - submit an invoice or receipt to SAP Concur: MATCH an existing expense, or CREATE one, and ATTACH the PDF.

Two modes (CONCUR_MODE in .env):
  demo (default) - works offline against a sample ledger in <PUBLISH_DIR>/concur_demo, so the whole
                   create / match / attach story can be shown without a Concur tenant.
  live           - calls the SAP Concur v4 REST API (Quick Expense with image). Needs the CONCUR_* settings.
                   Written from Concur's public API documentation; confirm the payload in your sandbox first.
                   Live mode creates the expense and attaches the receipt; matching against existing
                   expenses needs the tenant's own endpoint and is demonstrated in demo mode.
"""
import base64
import datetime
import difflib
import json
import os
import re
from typing import List, Optional

import requests
from dateutil import parser as dateparser

SYMBOLS = {"$": "USD", "\u20ac": "EUR", "\u00a3": "GBP", "\u20b9": "INR", "rs": "INR", "inr": "INR", "usd": "USD",
           "eur": "EUR", "gbp": "GBP"}
VENDOR_KEYS = ["vendor_name", "merchant_name", "supplier_name", "vendor", "merchant", "supplier"]
DATE_KEYS = ["invoice_date", "date", "po_date", "transaction_date", "receipt_date"]
TOTAL_KEYS = ["total_amount", "total", "amount_due", "grand_total", "total_due"]
REF_KEYS = ["invoice_number", "po_number", "receipt_number", "reference"]

SEED = [  # sample company-card transactions that are still waiting for a receipt (demo data)
    {"id": "CX-1001", "vendor": "Greenfield Foods", "date": "2026-10-30", "amount": 1155.00, "currency": "INR"},
    {"id": "CX-1002", "vendor": "Bluewave Supplies Pvt Ltd", "date": "2026-11-14", "amount": 2865.50, "currency": "USD"},
    {"id": "CX-1003", "vendor": "Metro Hardware Traders", "date": "2026-11-05", "amount": 178400.00, "currency": "INR"},
    {"id": "CX-1004", "vendor": "Fresh Basket Supermarket", "date": "2026-11-12", "amount": 388.50, "currency": "INR"},
    {"id": "CX-1005", "vendor": "City Cabs", "date": "2026-11-10", "amount": 412.00, "currency": "INR"},
]


def mode() -> str:
    return "live" if os.getenv("CONCUR_MODE", "demo").lower() == "live" else "demo"


def _num(text) -> Optional[float]:
    from core import parse_amount
    return parse_amount(text)


def _date(text) -> Optional[str]:
    try:
        return dateparser.parse(str(text), dayfirst=True).date().isoformat()
    except Exception:
        return None


def _first(kv: dict, keys: List[str]) -> str:
    low = {k.lower(): v for k, v in kv.items()}
    for k in keys:
        if low.get(k):
            return str(low[k])
    return ""


def _currency(kv: dict, amount_text: str) -> Optional[str]:
    cur = _first(kv, ["currency"]).strip().lower()
    if cur in SYMBOLS:
        return SYMBOLS[cur]
    if re.fullmatch(r"[a-z]{3}", cur or ""):
        return cur.upper()
    t = (amount_text or "").lower()
    for sym, code in SYMBOLS.items():
        if sym in t:
            return code
    return None


def expense_from_result(result: dict) -> dict:
    """Turn the structured result into the few things an expense needs."""
    kv = result["key_values"]
    total_text = _first(kv, TOTAL_KEYS)
    exp = {"vendor": _first(kv, VENDOR_KEYS), "date": _date(_first(kv, DATE_KEYS)) or "",
           "amount": _num(total_text), "currency": _currency(kv, total_text),
           "reference": _first(kv, REF_KEYS), "document_type": result["document_type"]}
    exp["description"] = " ".join(x for x in [result["document_type"].replace("_", " ").title(), exp["reference"]] if x)
    exp["missing"] = [k for k, ok in (("vendor", exp["vendor"]), ("date", exp["date"]),
                                      ("total amount", exp["amount"] is not None)) if not ok]
    return exp


# ---------------------------------------------------------------- matching
def _norm_vendor(v: str) -> str:
    v = re.sub(r"[^a-z0-9 ]", " ", (v or "").lower())
    return " ".join(w for w in v.split() if w not in {"pvt", "ltd", "llp", "inc", "co", "limited", "the", "private"})


def _vendor_sim(a: str, b: str) -> float:
    a, b = _norm_vendor(a), _norm_vendor(b)
    if not a or not b:
        return 0.0
    if a in b or b in a:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def match_expense(exp: dict, candidates: List[dict]):
    """Best existing expense for this document, or (None, 0). Same amount, close date, similar vendor."""
    best, best_score = None, 0.0
    if exp.get("amount") is None:
        return None, 0.0
    for c in candidates:
        if abs(float(c["amount"]) - exp["amount"]) > max(0.01, 0.005 * exp["amount"]):
            continue
        if exp.get("currency") and c.get("currency") and exp["currency"] != c["currency"]:
            continue
        days = 0
        if exp.get("date") and c.get("date"):
            days = abs((datetime.date.fromisoformat(c["date"]) - datetime.date.fromisoformat(exp["date"])).days)
            if days > 3:
                continue
        sim = _vendor_sim(exp.get("vendor", ""), c.get("vendor", ""))
        if sim < 0.5:
            continue
        score = round(0.6 * sim + 0.4 * (1 - days / 4), 2)
        if score > best_score:
            best, best_score = c, score
    return best, best_score


# ---------------------------------------------------------------- demo mode (offline sample ledger)
def _demo_dir() -> str:
    d = os.path.join(os.getenv("PUBLISH_DIR", "published"), "concur_demo")
    os.makedirs(os.path.join(d, "receipts"), exist_ok=True)
    return d


def _ledger() -> List[dict]:
    path = os.path.join(_demo_dir(), "expenses.json")
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump([dict(e, source="Corporate card", receipt=None, status="Awaiting receipt") for e in SEED], f, indent=2)
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _save_ledger(rows: List[dict]) -> None:
    with open(os.path.join(_demo_dir(), "expenses.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2)


def find_match(result: dict):
    """Preview: which sample-ledger expense would this document match? (demo mode)"""
    exp = expense_from_result(result)
    open_rows = [c for c in _ledger() if not c.get("receipt")]
    return match_expense(exp, open_rows)


def already_submitted(result: dict):
    """Demo mode: the ledger entry this document was already attached to, if any."""
    exp, name = expense_from_result(result), result["document"]["file_name"]
    return next((r for r in _ledger() if (r.get("receipt") or "").endswith("_" + name)
                 and exp["amount"] is not None and abs(float(r["amount"]) - exp["amount"]) < 0.01), None)


def _submit_demo(exp: dict, receipt_name: str, receipt_bytes: bytes) -> str:
    rows = _ledger()
    dup = next((r for r in rows if (r.get("receipt") or "").endswith("_" + receipt_name)
                and abs(float(r["amount"]) - exp["amount"]) < 0.01), None)
    if dup:
        return f"[Concur demo] Already submitted: {receipt_name} is attached to expense {dup['id']}. Nothing changed."
    match, score = match_expense(exp, [c for c in rows if not c.get("receipt")])
    if match:
        entry = next(r for r in rows if r["id"] == match["id"])
        action = f"Matched existing expense {entry['id']} ({entry['vendor']}, match score {score})"
    else:
        nxt = max([int(r["id"].split("-")[1]) for r in rows] + [1000]) + 1
        entry = {"id": f"CX-{nxt}", "vendor": exp["vendor"], "date": exp["date"], "amount": exp["amount"],
                 "currency": exp["currency"], "source": "Created from receipt",
                 "report": f"RPT-{datetime.date.today():%Y%m%d}"}
        rows.append(entry)
        action = f"No match found, created new expense {entry['id']} in report {entry['report']}"
    fname = f"{entry['id']}_{receipt_name}"
    with open(os.path.join(_demo_dir(), "receipts", fname), "wb") as f:
        f.write(receipt_bytes)
    entry["receipt"], entry["status"], entry["description"] = fname, "Receipt attached", exp["description"]
    _save_ledger(rows)
    return f"[Concur demo] {action}; attached {receipt_name}. Ledger: {os.path.abspath(os.path.join(_demo_dir(), 'expenses.json'))}"


# ---------------------------------------------------------------- live mode (SAP Concur v4)
def live_config() -> dict:
    keys = ["CONCUR_BASE_URL", "CONCUR_CLIENT_ID", "CONCUR_CLIENT_SECRET", "CONCUR_REFRESH_TOKEN", "CONCUR_EXPENSE_TYPE_ID"]
    return {k: os.getenv(k, "") for k in keys}


def quick_expense_payload(exp: dict) -> dict:
    """Body for POST /quickexpense/v4/users/{userId}/context/TRAVELER/quickexpenses/image"""
    return {"comment": exp["description"], "expenseTypeId": os.getenv("CONCUR_EXPENSE_TYPE_ID", "<expense-type-id>"),
            "transactionAmount": {"value": exp["amount"], "currencyCode": exp["currency"] or "USD"},
            "transactionDate": exp["date"], "vendor": {"description": exp["vendor"]}}


def _token(cfg: dict) -> str:
    r = requests.post(cfg["CONCUR_BASE_URL"].rstrip("/") + "/oauth2/v0/token",
                      data={"grant_type": "refresh_token", "refresh_token": cfg["CONCUR_REFRESH_TOKEN"],
                            "client_id": cfg["CONCUR_CLIENT_ID"], "client_secret": cfg["CONCUR_CLIENT_SECRET"]}, timeout=30)
    if r.status_code >= 300:
        raise RuntimeError(f"Concur login failed ({r.status_code}): {r.text[:200]}")
    return r.json()["access_token"]


def _user_id(token: str) -> str:
    if os.getenv("CONCUR_USER_ID"):
        return os.getenv("CONCUR_USER_ID")
    try:  # the user id is the 'sub' claim of the access token
        part = token.split(".")[1]
        return json.loads(base64.urlsafe_b64decode(part + "=" * (-len(part) % 4)))["sub"]
    except Exception:
        raise RuntimeError("Set CONCUR_USER_ID in .env (could not read it from the token).")


def _submit_live(exp: dict, receipt_name: str, receipt_bytes: bytes) -> str:
    cfg = live_config()
    gaps = [k for k, v in cfg.items() if not v]
    if gaps:
        raise RuntimeError("Concur live mode needs these settings in .env: " + ", ".join(gaps))
    tok = _token(cfg)
    url = f"{cfg['CONCUR_BASE_URL'].rstrip('/')}/quickexpense/v4/users/{_user_id(tok)}/context/TRAVELER/quickexpenses/image"
    r = requests.post(url, headers={"Authorization": f"Bearer {tok}"}, timeout=60,
                      files={"body": (None, json.dumps(quick_expense_payload(exp)), "application/json"),
                             "fileContent": (receipt_name, receipt_bytes, "application/pdf")})
    if r.status_code >= 300:
        raise RuntimeError(f"Concur said {r.status_code}: {r.text[:300]}")
    return f"[Concur live] Created quick expense with receipt attached ({r.status_code}). Response: {r.text[:120]}"


# ---------------------------------------------------------------- one entry point
def submit(result: dict, receipt_name: str, receipt_bytes: bytes) -> str:
    exp = expense_from_result(result)
    if exp["missing"]:
        raise ValueError("Cannot submit to Concur, missing: " + ", ".join(exp["missing"]))
    return (_submit_live if mode() == "live" else _submit_demo)(exp, receipt_name, receipt_bytes)
