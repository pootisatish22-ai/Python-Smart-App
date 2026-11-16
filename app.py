"""
Intelligence Plugin - "Upload a PDF. Understand it. Structure it. Export it."  (Streamlit screens)
All the logic lives in core.py, resume_builder.py, concur.py and publishers.py.
Run:  streamlit run app.py        (other systems can use api.py instead)
"""
import hashlib
import json
import os

import pandas as pd
import streamlit as st

import concur
import publishers as pub
import resume_builder
from core import (answer, edited_result, llm_ready, process, refresh_status, table_csvs, to_excel,
                  to_markdown, to_word, to_zip)

JOB_COLS = ["title", "company", "location", "start", "end", "highlights"]
EDU_COLS = ["degree", "institution", "year"]


def cell(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


# ---------------------------------------------------------------- resume builder tab
def resume_panel(d: dict, did: str):
    p = d["profile"] or {}
    st.write("Check the profile the AI extracted, change or add anything, then build the resume from a template.")
    a, b = st.columns(2)
    full_name = a.text_input("Full name", p.get("full_name") or "", key=f"rn_{did}")
    headline = b.text_input("Headline", p.get("headline") or "", key=f"rh_{did}")
    c1, c2, c3 = st.columns(3)
    email = c1.text_input("Email", p.get("email") or "", key=f"re_{did}")
    phone = c2.text_input("Phone", p.get("phone") or "", key=f"rp_{did}")
    location = c3.text_input("Location", p.get("location") or "", key=f"rl_{did}")
    links = st.text_area("Links (one per line)", "\n".join(p.get("links") or []), height=70, key=f"rk_{did}")
    summary = st.text_area("Summary", p.get("summary") or "", height=100, key=f"rs_{did}")
    skills = st.text_area("Skills (comma separated)", ", ".join(p.get("skills") or []), height=70, key=f"rsk_{did}")
    st.markdown("**Experience** (add or delete rows; separate highlights with a semicolon)")
    jobs = pd.DataFrame([{**{k: (j.get(k) or "") for k in JOB_COLS[:5]}, "highlights": "; ".join(j.get("bullets") or [])}
                         for j in p.get("experience") or []], columns=JOB_COLS)
    jobs_ed = st.data_editor(jobs, num_rows="dynamic", hide_index=True, key=f"rj_{did}")
    st.markdown("**Education**")
    edu = pd.DataFrame([{k: (e.get(k) or "") for k in EDU_COLS} for e in p.get("education") or []], columns=EDU_COLS)
    edu_ed = st.data_editor(edu, num_rows="dynamic", hide_index=True, key=f"rd_{did}")
    certs = st.text_area("Certifications (one per line)", "\n".join(p.get("certifications") or []), height=70, key=f"rc_{did}")

    profile = {
        "full_name": full_name, "headline": headline, "email": email, "phone": phone, "location": location,
        "links": [x.strip() for x in links.splitlines() if x.strip()], "summary": summary,
        "skills": [x.strip() for x in skills.split(",") if x.strip()],
        "experience": [{"title": cell(r.title), "company": cell(r.company), "location": cell(r.location),
                        "start": cell(r.start), "end": cell(r.end),
                        "bullets": [x.strip() for x in cell(r.highlights).split(";") if x.strip()]}
                       for r in jobs_ed.itertuples() if cell(r.title) or cell(r.company)],
        "education": [{"degree": cell(r.degree), "institution": cell(r.institution), "year": cell(r.year)}
                      for r in edu_ed.itertuples() if cell(r.degree) or cell(r.institution)],
        "certifications": [x.strip() for x in certs.splitlines() if x.strip()],
    }

    st.markdown("**Template**")
    choice = st.radio("Template", ["Built-in template", "My own template (.docx)"], horizontal=True,
                      label_visibility="collapsed", key=f"rtm_{did}")
    template = None
    if choice.startswith("My own"):
        up = st.file_uploader("Upload a .docx that uses {{placeholders}}", type=["docx"], key=f"rtu_{did}")
        template = up.getvalue() if up else None
        if up is None:
            st.info("Upload a template, or switch to the built-in one.")
    with st.expander("Placeholders you can use in your own template"):
        st.code("{{full_name}} {{headline}} {{email}} {{phone}} {{location}} {{contact}} {{links}} {{summary}} {{skills}}\n"
                "Own line each: {{experience}}  {{education}}  {{certifications}}")
    st.download_button("Download the built-in template", resume_builder.default_template(), "resume_template.docx",
                       key=f"rdt_{did}")
    if st.button("Build resume (Word)", key=f"rb_{did}", disabled=choice.startswith("My own") and template is None):
        try:
            st.session_state["resumes"][did] = resume_builder.build_resume(profile, template)
        except Exception as e:  # noqa
            st.error(f"Could not build the resume: {e}")
    built = st.session_state["resumes"].get(did)
    if built:
        st.success("Resume built from the template.")
        st.download_button("Download resume (Word)", built, f"{(full_name or 'resume').replace(' ', '_')}_resume.docx", key=f"rdl_{did}")
    return profile, built


# ---------------------------------------------------------------- publish panel
def publish_panel(d: dict, did: str, ap: dict, history: list, profile, built):
    cfg = pub.configured()
    result = {**ap["result"], "approved_by": ap["by"], "approved_at": ap["at"]}
    df = ap["df"]
    base = os.path.splitext(d["meta"]["file_name"])[0]
    result_json = json.dumps(result, indent=2)
    files = {"result.json": result_json, "fields.csv": df.to_csv(index=False), **table_csvs(d),
             "report.docx": to_word(d, df), "summary.md": to_markdown(d, df),
             "extracted_text.txt": "\n\n".join(d["pages"]), "original.pdf": d["pdf"]}

    def run(label, fn):
        try:
            msg = fn()
            history.append(f"{label}: {msg}")
            pub.audit("published", {"document_id": did, "target": label, "result": msg})
            st.success(msg)
        except Exception as e:  # noqa
            pub.audit("publish_failed", {"document_id": did, "target": label, "error": str(e)})
            st.error(f"{label} failed: {e}")

    exp = concur.expense_from_result(result)
    c = st.columns(5)
    if c[0].button("Save to folder", key=f"pf_{did}"):
        run("Folder", lambda: pub.publish_folder(result, files))
    if c[1].button("Add to Excel register", key=f"pe_{did}"):
        run("Excel", lambda: pub.publish_excel(result))
    if c[2].button("Submit to Concur", key=f"pc_{did}", disabled=bool(exp["missing"])):
        run("Concur", lambda: concur.submit(result, d["meta"]["file_name"], d["pdf"]))
    if c[3].button("Create JIRA issue", key=f"pj_{did}", disabled=not cfg["jira"]):
        run("JIRA", lambda: pub.publish_jira(result, f"{base}.json", result_json.encode()))
    if c[4].button("Create ServiceNow ticket", key=f"ps_{did}", disabled=not cfg["servicenow"]):
        run("ServiceNow", lambda: pub.publish_servicenow(result, f"{base}.json", result_json.encode()))
    if d.get("profile") is not None:
        if st.button("Build resume and save to folder", key=f"prs_{did}"):
            run("Resume", lambda: pub.publish_folder(
                result, {"resume.docx": built or resume_builder.build_resume(profile or d["profile"])}))
    if exp["missing"]:
        st.caption("Concur needs a vendor, a date and a total amount. Missing: " + ", ".join(exp["missing"]))
    else:
        st.caption(f"Concur mode: {concur.mode()}" + (" (sample ledger, works offline)" if concur.mode() == "demo" else ""))
    if not cfg["jira"]:
        st.caption("JIRA is not configured. Add the JIRA_* lines to .env and restart.")
    if not cfg["servicenow"]:
        st.caption("ServiceNow is not configured. Add the SNOW_* lines to .env and restart.")

    with st.expander("Preview what will be sent"):
        if not exp["missing"] and concur.mode() == "demo":
            m, score = concur.find_match(result)
            done = concur.already_submitted(result)
            st.markdown("**Concur: " + (f"already submitted, attached to expense {done['id']}**" if done else
                        f"would match existing expense {m['id']} ({m['vendor']}, score {score})**" if m
                        else "no matching expense found, a new one would be created**"))
        st.markdown("**Concur expense**")
        st.json({k: v for k, v in exp.items() if k != "missing"} if not exp["missing"] else exp)
        if not exp["missing"] and concur.mode() == "live":
            st.json(concur.quick_expense_payload(exp))
        st.markdown("**JIRA request body**")
        st.json(pub.jira_payload(result))
        st.markdown("**ServiceNow request body**")
        st.json(pub.snow_payload(result))
    if history:
        st.markdown("**Published so far**")
        for h in history:
            st.write(f"- {h}")


# ---------------------------------------------------------------- the screens
def main():
    st.set_page_config(page_title="Intelligence Plugin", layout="wide")
    for k in ("docs", "chats", "approvals", "published", "resumes"):
        st.session_state.setdefault(k, {})

    with st.sidebar:
        st.header("Intelligence Plugin")
        st.caption(f"LLM provider: **{os.getenv('LLM_PROVIDER', 'gemini')}**")
        st.caption(f"Concur: **{concur.mode()}**")
        st.subheader("Processing history")
        for dd in st.session_state["docs"].values():
            st.write(f"- {dd['meta']['file_name']} ({dd['doc_type']})")

    st.title("Upload a PDF. Understand it. Structure it. Export it.")
    err = llm_ready()
    if err:
        st.error(err)
        st.stop()

    up = st.file_uploader("Drag and drop a text-based PDF", type=["pdf"])
    if not up:
        st.info("Upload a PDF to begin.")
        return

    data = up.getvalue()
    did = hashlib.sha256(data).hexdigest()[:12]
    if did not in st.session_state["docs"]:
        with st.spinner("Validating, reading, classifying and structuring..."):
            try:
                st.session_state["docs"][did] = process(data, up.name)
            except ValueError as e:
                st.error(str(e))
                return
            except Exception as e:  # noqa
                st.error(f"Processing failed: {e}")
                st.caption("Check your API key and model name in .env, then try again.")
                return
        st.rerun()
    d = st.session_state["docs"][did]

    names = ["Overview", "Review and correct", "Tables"] + (["Resume builder"] if d["profile"] is not None else []) + \
            ["Ask the document", "Export", "Approve and publish"]
    T = dict(zip(names, st.tabs(names)))

    with T["Overview"]:
        st.subheader(d["title"])
        c1, c2, c3 = st.columns(3)
        c1.metric("Detected type", d["doc_type"])
        c2.metric("Pages", d["meta"]["pages"])
        c3.metric("File size (KB)", d["meta"]["size_kb"])
        st.caption(f"Why: {d['reason']}")
        st.write(d["summary"])
        st.markdown("**Key insights**")
        for i in d["insights"]:
            st.write(f"- {i}")
        st.caption(f"Document ID {d['meta']['document_id']}  |  SHA-256 {d['meta']['sha256'][:24]}...")

    with T["Review and correct"]:
        st.write("These are all the key/value pairs found. Edit any value; missing items update automatically.")
        edited = st.data_editor(d["df"], key=f"ed_{did}", hide_index=True,
                                disabled=["Field", "Page", "Confidence", "Status"])
        edited = refresh_status(edited)
        missing = edited.loc[edited["Status"] == "MISSING", "Field"].tolist()
        if missing:
            st.error("Missing information: " + ", ".join(missing))
        else:
            st.success("All expected fields were found.")
        flagged = edited[edited["Status"].isin(["Check format", "Low confidence"])]
        if len(flagged):
            st.warning("Please verify: " + ", ".join(flagged["Field"].tolist()))
        for w in d["warnings"]:
            st.warning(w)

    with T["Tables"]:
        if not d["tables"]:
            st.info("No item tables were found in this document.")
        for i, t in enumerate(d["tables"]):
            st.markdown(f"**{t['name']}**  ({len(t['df'])} rows)")
            st.dataframe(t["df"])
            st.download_button(f"Download {t['name']} (CSV)", t["df"].to_csv(index=False),
                               f"{t['name']}.csv", key=f"tcsv_{did}_{i}")
        if d["raw_tables"]:
            with st.expander("Tables detected in the PDF layout (raw)"):
                for p, t in d["raw_tables"]:
                    st.caption(f"Page {p}")
                    st.dataframe(t)

    profile, built = (None, None)
    if "Resume builder" in T:
        with T["Resume builder"]:
            profile, built = resume_panel(d, did)

    with T["Ask the document"]:
        hist = st.session_state["chats"].setdefault(did, [])
        for m in hist:
            with st.chat_message(m["role"]):
                st.write(m["content"])
        q = st.chat_input("Ask something about this document")
        if q:
            hist.append({"role": "user", "content": q})
            try:
                a = answer(d, q, hist[:-1])
            except Exception as e:  # noqa
                a = f"Sorry, the AI call failed: {e}"
            hist.append({"role": "assistant", "content": a})
            st.rerun()

    with T["Export"]:
        base = os.path.splitext(d["meta"]["file_name"])[0]
        st.write("Downloads include your edits from the Review tab.")
        c = st.columns(4)
        c[0].download_button("JSON", json.dumps(edited_result(d, edited), indent=2), f"{base}.json", "application/json")
        c[1].download_button("CSV (fields)", edited.to_csv(index=False), f"{base}.csv", "text/csv")
        c[2].download_button("Excel", to_excel(d, edited), f"{base}.xlsx")
        c[3].download_button("Word", to_word(d, edited), f"{base}.docx")
        c = st.columns(3)
        c[0].download_button("Markdown", to_markdown(d, edited), f"{base}.md")
        c[1].download_button("TXT (raw text)", "\n\n".join(d["pages"]), f"{base}.txt")
        c[2].download_button("ZIP (all outputs)", to_zip(d, edited), f"{base}_outputs.zip")
        st.caption("Other systems can consume this data through the API: run  uvicorn api:app --port 8000  "
                   "and POST the PDF to /process (JSON), then GET /documents/<id>/fields.csv or /tables/0.csv.")

    with T["Approve and publish"]:
        approvals = st.session_state["approvals"]
        history = st.session_state["published"].setdefault(did, [])
        missing_now = edited.loc[edited["Status"] == "MISSING", "Field"].tolist()
        st.markdown("**Step 1. Approve the reviewed data**")
        allowed = True
        if missing_now:
            st.warning("Still missing: " + ", ".join(missing_now))
            allowed = st.checkbox("I accept publishing with these fields missing", key=f"acc_{did}")
        who = st.text_input("Your name (recorded in the audit trail)", key=f"who_{did}")
        if st.button("Approve data", key=f"appb_{did}", disabled=not (allowed and who.strip())):
            snap = edited_result(d, edited)
            approvals[did] = {"by": who.strip(), "at": pub.now(), "result": snap, "df": edited.copy()}
            pub.audit("approved", {"document_id": did, "by": who.strip(), "missing": snap["missing"]})
        ap = approvals.get(did)
        st.markdown("**Step 2. Publish to downstream tools**")
        if not ap:
            st.info("Approve the data first. Publishing is locked until then.")
        else:
            st.success(f"Approved by {ap['by']} at {ap['at']}")
            if ap["result"] != edited_result(d, edited):
                st.warning("You changed values after approving. Approve again to publish "
                           "the latest values (the approved version is what gets sent).")
            publish_panel(d, did, ap, history, profile, built)
        with st.expander("Audit trail"):
            for e in pub.read_audit():
                st.write(e)


if __name__ == "__main__":
    main()
