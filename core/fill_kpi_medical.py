"""MOH Key Performance Indicators (Medical) filler - FRESH REPLACE.

Target Doc table: the one whose first row contains
"MOH Key Performance Indicators (Medical)". Its columns are:
  KPI | Target | % Before Exclusion | % After Exclusion | No. of excluded Patient

From the QAPI Excel KPI block (header row where col A='Target', col B='KPI'):
  col B = KPI name
  col D = % Before Exclusion [without <90/180]   -> Doc "% Before Exclusion"
  col E = % After Exclusion                       -> Doc "% After Exclusion"
  col I = total no. of excluded                   -> Doc "No. of excluded"

Only these three data columns are overwritten, matched by KPI name. Target is
left as-is. KPIs in the Doc not found in the Excel are left untouched. The
Operational KPI table is NOT touched by this filler.
"""
import re
from docx import Document
from .helpers import read_sheet
from .docx_tools import set_cell_text


def _pct(x):
    if x is None or x == "":
        return None
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    # Excel stores 0.90 -> 90%. If already >1 assume it's already a percent.
    if v <= 1.0:
        v *= 100
    return f"{round(v)}%"


def _norm(s):
    """Normalize a KPI name for fuzzy matching: lowercase, strip non-alnum."""
    if s is None:
        return ""
    s = str(s).lower()
    s = s.replace("\\", "/")
    return re.sub(r"[^a-z0-9]", "", s)


# Map Excel KPI name -> list of distinctive tokens to find the Doc row.
_KPI_KEYS = [
    ("vascularaccess", ["vascularaccess"]),
    ("map", ["meanarterialpressure", "map105", "map"]),
    ("ktv", ["ktv"]),
    ("bfr", ["bfr"]),
    ("idwg", ["interdialyticweightgain", "idwg"]),
    ("calcium", ["calcium", "correctedcalcium"]),
    ("phosphorus", ["phosphorus"]),
    ("ipth", ["ipth", "pth"]),
    ("hemoglobin", ["hemoglobin", "haemoglobin"]),
    ("albumin", ["albumin"]),
]


def _excel_key(name):
    n = _norm(name)
    for key, toks in _KPI_KEYS:
        if any(t in n for t in toks):
            return key
    return None


def _read_kpi_block(rows):
    """Find the KPI block (row with A='Target', B='KPI') and read data rows
    until names run out. Return dict key -> (before, after, excluded)."""
    start = None
    for i, r in enumerate(rows):
        a = str(r[0]).strip().lower() if len(r) > 0 and r[0] is not None else ""
        b = str(r[1]).strip().lower() if len(r) > 1 and r[1] is not None else ""
        if a == "target" and b == "kpi":
            # ensure this is the % block (has a 'total' header nearby), not the
            # outliers block (which has 'outliers' in C/D)
            tail = " ".join(str(c).lower() for c in r if c is not None)
            if "outlier" in tail:
                continue
            start = i + 1
            break
    if start is None:
        return {}
    out = {}
    for r in rows[start:]:
        name = r[1] if len(r) > 1 else None
        if name is None or str(name).strip() == "":
            break
        key = _excel_key(name)
        if not key:
            continue
        before = r[3] if len(r) > 3 else None   # col D
        after = r[4] if len(r) > 4 else None     # col E
        total = r[8] if len(r) > 8 else None     # col I
        out[key] = (_pct(before), _pct(after), total)
    return out


def _find_medical_table(doc):
    for t in doc.tables:
        if t.rows and "medical" in t.rows[0].cells[0].text.lower() \
                and "key performance" in t.rows[0].cells[0].text.lower():
            return t
    return None


def fill_table(doc, xlsx_path):
    """Fill the Medical KPI table in an already-open doc. Returns rows updated."""
    rows = read_sheet(xlsx_path)
    data = _read_kpi_block(rows)
    table = _find_medical_table(doc)
    if table is None or not data:
        return 0

    # locate the column indices from the header row (row 1)
    hdr = [c.text.strip().lower() for c in table.rows[1].cells]
    def col(*names):
        for idx, h in enumerate(hdr):
            if any(n in h for n in names):
                return idx
        return None
    c_before = col("before exclusion")
    c_after = col("after exclusion")
    c_excl = col("excluded")

    updated = 0
    for row in table.rows[2:]:
        kpi_name = row.cells[0].text.strip()
        key = _excel_key(kpi_name)
        if key and key in data:
            before, after, excl = data[key]
            if c_before is not None and before is not None:
                set_cell_text(row.cells[c_before], before)
            if c_after is not None and after is not None:
                set_cell_text(row.cells[c_after], after)
            if c_excl is not None and excl is not None:
                set_cell_text(row.cells[c_excl], str(excl))
            updated += 1
    return updated


def fill(doc_path, xlsx_path, out_path):
    doc = Document(doc_path)
    n = fill_table(doc, xlsx_path)
    doc.save(out_path)
    return n
