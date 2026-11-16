# Intelligence Plugin - source code v2

**Give PDF -> bot understands it -> outputs clean structured data -> sends it into tools.**

## What this version covers (the use-case brief)

| Brief | Where it is |
|---|---|
| Detect the document type (resume, invoice, receipt, ...) | `core.py` classification (8 types, easy to extend in `EXPECTED`) |
| Extract ALL relevant key/value pairs, including item-level tables | `core.py` `process()`: expected fields first, then every other relevant field (up to 30), plus item tables (`line_items`, `transactions`) |
| Export to Word / Excel | Export tab, and `core.to_word / to_excel` |
| Submit to Concur: create or match an expense and attach the invoice/receipt | `concur.py` (demo mode works offline, live mode calls the SAP Concur v4 API) |
| Build / update resumes by populating a resume template | `resume_builder.py` and the **Resume builder** tab (edit the profile, choose the built-in or your own .docx template) |
| Other systems consume structured data (API, JSON, CSV) | `api.py` (FastAPI) and the JSON / CSV downloads |

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit screens only |
| `core.py` | The engine: validation, extraction, checks, exports. No Streamlit inside |
| `api.py` | REST API for other systems |
| `concur.py` | Concur: expense matching, demo ledger, live API call |
| `resume_builder.py` | Fills a Word resume template from the extracted profile |
| `publishers.py` | Folder, Excel register (now also a Line items sheet), JIRA, ServiceNow, audit log |
| `resume_template.docx` | The built-in resume template (edit it and upload it as your own) |
| `tests/offline_check.py` | Runs everything with a stand-in AI (no key needed) |
| `test_llm.py`, `.env.example`, `requirements.txt` | Connection test, settings, libraries |

## Upgrade from the previous version
1. Stop the app (`Ctrl + C`). Copy ALL these files into your project folder (next to the `venv` folder, not inside it), replacing the old ones.
2. In the terminal with `(venv)` showing: `pip install -r requirements.txt`
3. Add the new lines from `.env.example` to your `.env` (at least `CONCUR_MODE=demo`). Keep your existing key lines. Use `GROQ_MODEL=openai/gpt-oss-120b` or the current model name.
4. `python test_llm.py` should print SUCCESS, then `streamlit run app.py`.

## Using the web app
- **Invoice or receipt:** Review tab shows every key/value pair; Tables tab shows the line items (CSV download per table). Approve, then **Submit to Concur**. Open "Preview what will be sent" to see whether it matches an existing expense or creates a new one.
- **Resume:** a **Resume builder** tab appears. Check or edit the profile, pick the template, click **Build resume (Word)**, download it. "Build resume and save to folder" in the publish step stores it with the other outputs.
- A warning appears when item amounts do not add up to the stated subtotal (or total).

## Concur
- **Demo mode (default, `CONCUR_MODE=demo`):** uses a sample ledger of company-card expenses that are waiting for receipts (`<PUBLISH_DIR>/concur_demo/expenses.json`). A document matches when the amount is the same, the date is within 3 days and the vendor name is similar. A match gets the PDF attached; otherwise a new expense is created in a new report with the PDF attached. Submitting the same file twice does not duplicate. The sample ledger contains entries that match the dummy test invoices, purchase order and receipt.
- **Live mode (`CONCUR_MODE=live`):** needs `CONCUR_BASE_URL`, `CONCUR_CLIENT_ID`, `CONCUR_CLIENT_SECRET`, `CONCUR_REFRESH_TOKEN`, `CONCUR_EXPENSE_TYPE_ID` (and `CONCUR_USER_ID` if it cannot be read from the token). It calls the Concur v4 Quick Expense API with the image upload (`POST /quickexpense/v4/users/{userId}/context/TRAVELER/quickexpenses/image`). **This was written from Concur's public documentation and has not been run against a real tenant.** Check the payload shown in the preview against your sandbox. Live mode creates the expense and attaches the receipt; matching against existing expenses needs your tenant's endpoint and is shown in demo mode only.

## Resume templates
Make any Word document and type placeholders: `{{full_name}} {{headline}} {{email}} {{phone}} {{location}} {{contact}} {{links}} {{summary}} {{skills}}`. Put `{{experience}}`, `{{education}}` and `{{certifications}}` each alone on their own line. Empty values remove their line and the heading just above it.

## API for other systems
```
uvicorn api:app --port 8000        # interactive docs at http://localhost:8000/docs
curl -F "file=@invoice.pdf" http://localhost:8000/process
curl http://localhost:8000/documents/<id>/fields.csv
curl http://localhost:8000/documents/<id>/tables/0.csv
curl -X POST http://localhost:8000/documents/<id>/publish/concur       # folder | excel | concur | jira | servicenow
curl -o resume.docx http://localhost:8000/documents/<id>/resume.docx    # for resumes
```
Set `PLUGIN_API_KEY` in `.env` to require an `X-API-Key` header. Documents are kept in memory while the API runs.

Shape of the JSON (values illustrative):
```
{"document": {"document_id": "...", "file_name": "...", "pages": 1},
 "document_type": "invoice", "title": "...", "summary": "...", "insights": ["..."],
 "fields": [{"Field": "invoice_number", "Value": "INV-1001", "Page": "1", "Confidence": "0.99", "Status": "OK"}],
 "key_values": {"invoice_number": "INV-1001", "vendor_name": "..."},
 "missing": ["due_date"], "warnings": [],
 "tables": [{"name": "line_items", "columns": ["Description", "Qty", "Unit price", "Amount"], "rows": [{"Description": "..."}]}],
 "profile": null}
```

## Checking it works
- `python tests/offline_check.py test_pdfs` runs 12 checks with a stand-in AI (no key). It proves the plumbing, not AI accuracy.
- Measure real accuracy with your AI on the 20 test PDFs and the Results and Evaluation sheet.

## Known limits
Text PDFs only (no OCR yet). Dates such as 11/05/2026 are read day-first. Item tables rely on the AI, with the PDF's own table detection as a fallback. JIRA and ServiceNow were tested against mock services only. Free AI tiers may keep submitted data, so use dummy documents.
