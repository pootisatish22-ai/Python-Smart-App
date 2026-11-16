"""
resume_builder.py - populate a Word resume template from an extracted profile.

Template placeholders (type them in any .docx):
  one-line values : {{full_name}} {{headline}} {{email}} {{phone}} {{location}} {{contact}} {{links}} {{summary}} {{skills}}
  whole-line blocks (put each alone on its own line): {{experience}} {{education}} {{certifications}}
{{contact}} is email | phone | location, joined only with the parts that exist.
An empty block or empty value removes its line, and the heading just above it if there is one.
"""
import io
import re
from typing import Optional

from docx import Document
from docx.shared import Inches, Pt

BLOCKS = ("experience", "education", "certifications")
BLOCK_RE = re.compile(r"\{\{\s*(experience|education|certifications)\s*\}\}")
INLINE_RE = re.compile(r"\{\{\s*(\w+)\s*\}\}")


def default_template() -> bytes:
    d = Document()
    for s in d.sections:
        s.left_margin = s.right_margin = Inches(0.9)
        s.top_margin = s.bottom_margin = Inches(0.8)
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(10.5)
    d.add_paragraph("{{full_name}}", style="Title")
    d.add_paragraph("{{headline}}")
    d.add_paragraph("{{contact}}")
    d.add_paragraph("{{links}}")
    for head, ph in [("Professional Summary", "{{summary}}"), ("Skills", "{{skills}}"),
                     ("Experience", "{{experience}}"), ("Education", "{{education}}"),
                     ("Certifications", "{{certifications}}")]:
        d.add_heading(head, 2)
        d.add_paragraph(ph)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def _clean(v) -> str:
    return "" if v is None else str(v).strip()


def _mapping(p: dict) -> dict:
    contact = " | ".join(x for x in [_clean(p.get("email")), _clean(p.get("phone")), _clean(p.get("location"))] if x)
    return {
        "full_name": _clean(p.get("full_name")), "headline": _clean(p.get("headline")),
        "email": _clean(p.get("email")), "phone": _clean(p.get("phone")), "location": _clean(p.get("location")),
        "contact": contact, "links": " | ".join(_clean(x) for x in p.get("links") or [] if _clean(x)),
        "summary": _clean(p.get("summary")), "skills": ", ".join(_clean(x) for x in p.get("skills") or [] if _clean(x)),
    }


def _all_paragraphs(doc):
    out = list(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                out.extend(cell.paragraphs)
    return out


def _remove(par):
    el = par._element
    el.getparent().remove(el)


def _remove_with_heading(par):
    prev = par._element.getprevious()
    _remove(par)
    if prev is not None and prev.tag.endswith("}p"):
        from docx.text.paragraph import Paragraph
        pp = Paragraph(prev, par._parent)
        if (pp.style.name or "").lower().startswith("heading"):
            _remove(pp)


def _bullet(anchor, text):
    try:
        anchor.insert_paragraph_before(text, style="List Bullet")
    except (KeyError, ValueError):
        anchor.insert_paragraph_before("\u2022 " + text)


def _block_lines(kind: str, p: dict, anchor) -> int:
    """Insert the block's paragraphs before the anchor paragraph. Returns how many were added."""
    n = 0
    if kind == "experience":
        for j in p.get("experience") or []:
            title, co = _clean(j.get("title")), _clean(j.get("company"))
            if not (title or co):
                continue
            para = anchor.insert_paragraph_before()
            para.add_run(", ".join(x for x in [title, co] if x)).bold = True
            when = " - ".join(x for x in [_clean(j.get("start")), _clean(j.get("end"))] if x)
            extra = "  |  ".join(x for x in [_clean(j.get("location")), when] if x)
            if extra:
                para.add_run("   " + extra)
            n += 1
            for b in j.get("bullets") or []:
                if _clean(b):
                    _bullet(anchor, _clean(b)); n += 1
    elif kind == "education":
        for e in p.get("education") or []:
            line = ", ".join(x for x in [_clean(e.get("degree")), _clean(e.get("institution"))] if x)
            if _clean(e.get("year")):
                line += f" ({_clean(e.get('year'))})"
            if line:
                anchor.insert_paragraph_before(line); n += 1
    else:
        for c in p.get("certifications") or []:
            if _clean(c):
                _bullet(anchor, _clean(c)); n += 1
    return n


def build_resume(profile: dict, template: Optional[bytes] = None) -> bytes:
    """Return a .docx made by filling the template (or the built-in default) with the profile."""
    doc = Document(io.BytesIO(template or default_template()))
    mapping = _mapping(profile)
    for par in _all_paragraphs(doc):
        m = BLOCK_RE.fullmatch(par.text.strip())
        if m:
            if _block_lines(m.group(1), profile, par) == 0:
                _remove_with_heading(par)
            else:
                _remove(par)
    for par in _all_paragraphs(doc):
        text = par.text
        if "{{" not in text:
            continue
        new = INLINE_RE.sub(lambda m: mapping.get(m.group(1), m.group(0)), text)
        if new.strip() == "" or new.strip(" |") == "":
            _remove_with_heading(par)
        elif new != text:
            runs = par.runs
            if runs:
                runs[0].text = new
                for r in runs[1:]:
                    r.text = ""
            else:
                par.add_run(new)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
