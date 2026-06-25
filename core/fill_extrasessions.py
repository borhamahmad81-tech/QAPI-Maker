"""Extrasessions filler.

Excel columns: Name, ID, No. of sessions, Date, Cause, Date of MOH approval

Patients are identified by ID (not name). All rows sharing an ID are merged into
ONE Doc row:
  - Name: the first non-empty name found for that ID (minor spelling variants
    across rows are ignored; first wins).
  - Sessions: total count of session-rows for that ID -> "N session(s)".
  - Dates: every session date, in order of appearance, comma-separated.
  - Reasons: all reasons, de-duplicated (case-insensitive), comma-separated.
  - MOH approval: all approval dates, de-duplicated, comma-separated.
"""
from docx import Document
from .helpers import read_sheet, fmt_date, clean_id, clean_text, is_formula_or_blank
from .docx_tools import set_cell_text, insert_row_after


def _find_table_and_header(doc):
    for t in doc.tables:
        for ri, row in enumerate(t.rows):
            joined = " ".join(c.text.lower() for c in row.cells)
            if "patient id" in joined and "extra session" in joined:
                return t, ri
    raise ValueError("Extrasessions table/header not found in template.")


def _sessions_label(n):
    return f"{n} session" if n == 1 else f"{n} sessions"


def _dedupe_keep_order(items):
    seen = set()
    out = []
    for it in items:
        s = (it or "").strip()
        if not s:
            continue
        key = s.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out


def _group_by_id(rows):
    order = []
    groups = {}
    for r in rows[1:]:
        if not r:
            continue
        name = clean_text(r[0]) if len(r) > 0 else ""
        rid = clean_id(r[1]) if len(r) > 1 else ""

        if not rid and is_formula_or_blank(name):
            continue
        if not rid and not name:
            continue

        key = rid if rid else f"name::{name.lower()}"

        date = fmt_date(r[3]) if len(r) > 3 else ""
        cause = clean_text(r[4]) if len(r) > 4 else ""
        moh = fmt_date(r[5]) if len(r) > 5 else ""

        if key not in groups:
            groups[key] = {"name": name, "id": rid, "dates": [],
                           "reasons": [], "moh": [], "count": 0}
            order.append(key)

        g = groups[key]
        if not g["name"] and name:
            g["name"] = name
        g["count"] += 1
        if date:
            g["dates"].append(date)
        if cause:
            g["reasons"].append(cause)
        if moh:
            g["moh"].append(moh)

    return [groups[k] for k in order]


def fill(doc_path, xlsx_path, out_path):
    rows = read_sheet(xlsx_path)
    patients = _group_by_id(rows)

    doc = Document(doc_path)
    table, header_idx = _find_table_and_header(doc)
    first_data_idx = header_idx + 1

    existing = len(table.rows) - first_data_idx
    while existing < len(patients):
        insert_row_after(table, len(table.rows) - 1)
        existing += 1

    for i, p in enumerate(patients):
        cells = table.rows[first_data_idx + i].cells
        # Real template format: Name on line 1, ID on line 2.
        name_id = p["name"]
        if p["id"]:
            name_id = f"{p['name']}\n{p['id']}" if p["name"] else p["id"]
        # "N sessions" on line 1, dates (comma-separated) on line 2.
        sess = _sessions_label(p["count"])
        dates = ", ".join(p["dates"])
        col1 = f"{sess}\n{dates}" if dates else sess
        reasons = ", ".join(_dedupe_keep_order(p["reasons"]))
        approvals = ", ".join(_dedupe_keep_order(p["moh"]))
        set_cell_text(cells[0], name_id)
        set_cell_text(cells[1], col1)
        set_cell_text(cells[2], reasons)
        set_cell_text(cells[3], approvals)

    for j in range(first_data_idx + len(patients), len(table.rows)):
        for cell in table.rows[j].cells:
            set_cell_text(cell, "")

    doc.save(out_path)
    return len(patients)
