"""Hospitalization filler.

Excel columns: Name, ID, Hospital, Admission Date, Discharge Date, Cause,
               blood Transfusion
Doc: "Patients Census" table. We fill ONLY the "Hospitalized Patients" merged
cell (the long remarks cell on that row) with a numbered list. All count cells
are left for manual entry.

Per-patient block (from filled example):
  N. Name: ID
        Date admitted: <Admission>
        Cause: <Cause>
        Hospital: <Hospital>
        Outcome: Discharged <Discharge>   (if discharge present)
        Blood transfusion: <value>        (only if Excel has a value)
        Received in stable condition       (if discharged)

Rules confirmed by user:
- Every row = separate entry (second admissions included).
- If discharged -> outcome improved / received in stable condition.
- If no blood transfusion value in Excel -> omit that line entirely.
"""
from docx import Document
from .helpers import read_sheet, fmt_date, clean_id, clean_text, is_formula_or_blank
from .docx_tools import set_cell_text


def _resolve(value, rows, col, ref_map):
    """Resolve a value that may be an Excel cross-reference like =$A$7."""
    if value is None:
        return None
    s = str(value).strip()
    if s.startswith("=$"):
        # e.g. =$A$7  -> column A, row 7 (1-indexed in Excel)
        try:
            letter = s[2]
            rownum = int(s.split("$")[2])
            cidx = ord(letter.upper()) - ord("A")
            return rows[rownum - 1][cidx]
        except Exception:
            return None
    return value


def _build_block(idx, p):
    indent = " " * 14
    lines = [f"{idx}. {p['name']}: {p['id']}"]
    if p["admission"]:
        lines.append(f"{indent}Date admitted: {p['admission']}")
    if p["cause"]:
        lines.append(f"{indent}Cause: {p['cause']}")
    if p["hospital"]:
        lines.append(f"{indent}Hospital: {p['hospital']}")
    if p["discharge"]:
        lines.append(f"{indent}Outcome: Discharged on {p['discharge']}")
    if p["transfusion"]:
        lines.append(f"{indent}Blood transfusion: {p['transfusion']}")
    if p["discharge"]:
        lines.append(f"{indent}Received in stable condition")
    return "\n".join(lines)


def _find_hosp_cell(doc):
    """Return (table, row_idx) for the 'Hospitalized Patients' row."""
    for t in doc.tables:
        for ri, row in enumerate(t.rows):
            if row.cells and "hospitalized patients" in row.cells[0].text.strip().lower():
                return t, ri
    raise ValueError("Hospitalized Patients row not found in template.")


def fill(doc_path, xlsx_path, out_path):
    rows = read_sheet(xlsx_path)
    patients = []
    for r in rows[1:]:
        if not r:
            continue
        name_raw = r[0] if len(r) > 0 else None
        id_raw = r[1] if len(r) > 1 else None
        # resolve cross-references for name/id
        name = clean_text(_resolve(name_raw, rows, 0, None))
        rid = clean_id(_resolve(id_raw, rows, 1, None))
        if not name and not rid:
            continue
        patients.append({
            "name": name,
            "id": rid,
            "hospital": clean_text(r[2]) if len(r) > 2 else "",
            "admission": fmt_date(r[3]) if len(r) > 3 else "",
            "discharge": fmt_date(r[4]) if len(r) > 4 else "",
            "cause": clean_text(r[5]) if len(r) > 5 else "",
            "transfusion": clean_text(r[6]) if len(r) > 6 else "",
        })

    doc = Document(doc_path)
    table, row_idx = _find_hosp_cell(doc)

    blocks = [_build_block(i + 1, p) for i, p in enumerate(patients)]
    text = "\n\n".join(blocks)

    # The remarks portion is the merged cell spanning columns 2..end on that row.
    # Find the first cell on the row that is NOT the label and NOT the count col.
    cells = table.rows[row_idx].cells
    # label = cells[0]; count = cells[1]; remarks merged = cells[2]
    set_cell_text(cells[1], str(len(patients)))
    set_cell_text(cells[2], text)

    doc.save(out_path)
    return len(patients)
