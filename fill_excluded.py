"""Details of Excluded Patients filler - RECONCILE by (ID + KPI category).

Doc table: first row contains "Details of Excluded Patients". Columns:
  Excluded Patient ID | Excluded Medical KPI | Exclusion Reason | Action Plan

A patient may be excluded for several KPIs -> one row per (patient, KPI).

Excel source (QAPI sheet): each KPI block has a Name_ID column and an exclusion
code column. We locate them by header name:
  Hemoglobin  : Name_ID just left of 'HB_ex_code'
  Phosphorus  : Name_ID just left of the 'BMD_ex_code' in the Phosph block
  Calcium     : Name_ID just left of the 'BMD_ex_code' in the Calcium block
  PTH         : Name_ID just left of the 'BMD_ex_code' in the PTH block
  Vascular    : Name_ID just left of 'VAC_exclusion'

Because 'BMD_ex_code' repeats for Phosph/Calcium/PTH, we identify those blocks
by the value-header to the right (Phosph_mg/dl, Calcium, PTH_pg/ml).

Reconcile rules:
  - matched (ID + KPI category in both)  -> keep the Doc row untouched
  - new (in Excel, not in Doc)           -> add row: Name (ID) | KPI category |
                                            exclusion code | (blank action plan)
  - stale (in Doc, not in Excel)         -> delete row
"""
import re
from docx import Document
from .helpers import read_sheet, clean_text
from .docx_tools import set_cell_text, insert_row_after, delete_row

_ID_RE = re.compile(r"\b(\d{10})\b")


def _extract_id(text):
    if not text:
        return ""
    m = _ID_RE.search(str(text))
    return m.group(1) if m else ""


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower()) if s else ""


# KPI category label as it should appear in the Doc's "Excluded Medical KPI" col.
CATEGORY_LABEL = {
    "hemoglobin": "Hemoglobin",
    "phosphorus": "Phosphorus",
    "calcium": "Calcium",
    "pth": "iPTH",
    "vascular": "Vascular access",
}

# How to find each block's (Name_ID col, ex_code col) via header anchors.
# anchor = the exclusion-code header; value_hint = a header to the right that
# disambiguates repeated 'BMD_ex_code' columns.
BLOCK_ANCHORS = [
    ("hemoglobin", "hb_ex_code", None),
    ("phosphorus", "bmd_ex_code", "phosph"),
    ("calcium", "bmd_ex_code", "calcium"),
    ("pth", "bmd_ex_code", "pth"),
    ("vascular", "vac_exclusion", None),
]


def _find_header_row(rows):
    for i, r in enumerate(rows):
        cells = [str(c).strip().lower() if c is not None else "" for c in r]
        if cells.count("name_id") >= 2 and any("ex_code" in c or "exclusion" in c
                                               for c in cells):
            return i, cells
    raise ValueError("Could not find the excluded-patients header row.")


def _locate_block(header_cells, anchor, value_hint):
    """Return (name_id_col, code_col) for a block, or None."""
    occurrences = [i for i, h in enumerate(header_cells) if anchor in h]
    for code_col in occurrences:
        # find Name_ID to the left
        name_col = None
        for j in range(code_col - 1, -1, -1):
            if header_cells[j] == "name_id":
                name_col = j
                break
        if name_col is None:
            continue
        if value_hint:
            # check a header to the right of code_col contains the hint
            right = " ".join(header_cells[code_col:code_col + 6])
            if value_hint not in right:
                continue
        return name_col, code_col
    return None


def _dedupe_code(code):
    """Some source cells contain the same code repeated, e.g.
    'Dvexclusion_005,Dvexclusion_005'. Collapse exact repeats, keep order."""
    if not code:
        return code
    parts = [p.strip() for p in code.split(",") if p.strip()]
    seen = []
    for p in parts:
        if p not in seen:
            seen.append(p)
    return ",".join(seen)


def _collect_excluded(rows):
    """Return list of dicts: {id, name_id, category, code} from the Excel."""
    header_row, header_cells = _find_header_row(rows)
    out = []
    seen = set()
    for cat_key, anchor, hint in BLOCK_ANCHORS:
        loc = _locate_block(header_cells, anchor, hint)
        if not loc:
            continue
        name_col, code_col = loc
        for r in rows[header_row + 1:]:
            if name_col >= len(r) or code_col >= len(r):
                continue
            name_id = clean_text(r[name_col])
            code = _dedupe_code(clean_text(r[code_col]))
            if not name_id or name_id.lower() in ("name_id", "(blank)"):
                continue
            # The medical blocks use 'Dvexclusion_NNN' codes; the vascular block
            # uses free text like 'Unfit' / 'Refusal'. Accept any non-blank code.
            if not code or code.lower() in ("(blank)", "none"):
                continue
            pid = _extract_id(name_id)
            if not pid:
                continue
            key = (pid, cat_key)
            if key in seen:
                continue
            seen.add(key)
            out.append({"id": pid, "name_id": name_id,
                        "category": cat_key, "code": code})
    return out


def _doc_category_key(text):
    """Map a Doc 'Excluded Medical KPI' cell to a category key."""
    n = _norm(text)
    if "vascular" in n or "avf" in n or "avg" in n:
        return "vascular"
    if "hemoglobin" in n or "haemoglobin" in n or n == "hb":
        return "hemoglobin"
    if "phosph" in n:
        return "phosphorus"
    if "calcium" in n:
        return "calcium"
    if "pth" in n:
        return "pth"
    return None


def _find_table(doc):
    for t in doc.tables:
        if t.rows and "details of excluded" in t.rows[0].cells[0].text.lower():
            return t
    return None


def _row_texts(table, row_idx):
    """Read a row's raw <w:tc> cells as plain text, regardless of how many
    physical cells it has. Real DaVita rows come in two shapes:
      4 cells -> ID | KPI category | Exclusion Reason | Action Plan
      3 cells -> ID | Exclusion Reason | Action Plan   (the KPI-category cell
                 is simply absent on that row; it belongs to whichever group
                 it sits inside, same as the rows above it)
    Returns (id_text, category_text_or_None, reason_text, plan_text).
    """
    from docx.oxml.ns import qn
    tcs = table.rows[row_idx]._tr.findall(qn("w:tc"))

    def txt(tc):
        return "".join(n.text or "" for n in tc.iter(qn("w:t"))).strip()

    vals = [txt(tc) for tc in tcs]
    if len(vals) >= 4:
        return vals[0], vals[1], vals[2], vals[3]
    if len(vals) == 3:
        return vals[0], None, vals[1], vals[2]
    if len(vals) == 2:
        return vals[0], None, vals[1], ""
    if len(vals) == 1:
        return vals[0], None, "", ""
    return "", None, "", ""


def fill_table(doc, xlsx_path):
    rows = read_sheet(xlsx_path)
    excluded = _collect_excluded(rows)
    table = _find_table(doc)
    if table is None:
        return 0

    excel_map = {}
    excel_order = []
    for e in excluded:
        k = (e["id"], e["category"])
        if k not in excel_map:
            excel_map[k] = e
            excel_order.append(k)

    # Walk existing Doc data rows (row 0 = title, row 1 = headers), carrying
    # the category label DOWN through rows whose KPI-category cell is absent.
    # This matches how the document was actually typed: a block of rows under
    # one category heading, not every row repeating the label.
    #
    # Special case: rows BEFORE the first labelled row in the whole table have
    # no label to carry down from. They belong to the group that follows them
    # (the document's very first category section), so for those leading rows
    # we look FORWARD to the next labelled row instead.
    raw_rows = []  # (row_idx, id_text, cat_text_or_None, reason, plan)
    for ri in range(2, len(table.rows)):
        id_text, cat_text, reason_text, plan_text = _row_texts(table, ri)
        raw_rows.append((ri, id_text, cat_text, reason_text, plan_text))

    first_label_idx = None
    for i, (ri, idt, cat, rs, pl) in enumerate(raw_rows):
        if cat is not None and cat.strip():
            first_label_idx = i
            break
    leading_label = None
    if first_label_idx is not None and first_label_idx > 0:
        leading_label = raw_rows[first_label_idx][2].strip()

    last_cat_key = None
    last_cat_label = ""
    row_info = []  # (row_idx, id, cat_key, cat_label)
    for i, (ri, id_text, cat_text, reason_text, plan_text) in enumerate(raw_rows):
        pid = _extract_id(id_text)
        if cat_text is not None and cat_text.strip():
            last_cat_label = cat_text.strip()
            last_cat_key = _doc_category_key(cat_text)
        elif last_cat_key is None and leading_label:
            # before the first labelled row: borrow the upcoming label
            last_cat_label = leading_label
            last_cat_key = _doc_category_key(leading_label)
        if not pid:
            row_info.append((ri, None, None, last_cat_label))
            continue
        row_info.append((ri, pid, last_cat_key, last_cat_label))

    # Decide keep vs delete, and record the LAST row of each category group so
    # new patients can be inserted at the end of their own group rather than
    # at the bottom of the whole table.
    keep_keys = set()
    rows_to_delete = []
    group_last_row = {}

    for ri, pid, cat_key, cat_label in row_info:
        if cat_key:
            group_last_row[cat_key] = ri
        if pid is None:
            continue
        key = (pid, cat_key) if cat_key else None
        if key and key in excel_map:
            keep_keys.add(key)
        else:
            rows_to_delete.append(ri)

    # Delete stale rows from the BOTTOM up so earlier indices stay valid, and
    # shift any recorded group_last_row index that was at/after a deleted row.
    for ri in sorted(rows_to_delete, reverse=True):
        delete_row(table, ri)
        for k in list(group_last_row):
            if group_last_row[k] >= ri:
                group_last_row[k] -= 1

    # Add NEW excluded (in Excel, not already kept) inside their own group,
    # right after the last existing row of that category. A category with no
    # existing rows in the doc is appended at the very end of the table.
    added = 0
    by_cat = {}
    for key in excel_order:
        if key in keep_keys:
            continue
        by_cat.setdefault(key[1], []).append(key)

    for cat_key, keys in by_cat.items():
        insert_after = group_last_row.get(cat_key, len(table.rows) - 1)
        # Clone a row that actually HAS the KPI-category cell (4 physical
        # cells) so the new row keeps that column. Prefer the group's own
        # last row if it qualifies; otherwise scan for any 4-cell row.
        _, cat_at_insert, _, _ = _row_texts(table, insert_after)
        clone_src = insert_after
        existing_label = cat_at_insert
        if cat_at_insert is None:
            for probe in range(2, len(table.rows)):
                _, cprobe, _, _ = _row_texts(table, probe)
                if cprobe is not None and _doc_category_key(cprobe) == cat_key:
                    clone_src = probe
                    existing_label = cprobe
                    break
        label = existing_label.strip() if existing_label else CATEGORY_LABEL.get(cat_key, "")
        for key in keys:
            e = excel_map[key]
            new_row = insert_row_after(table, insert_after, src_row_idx=clone_src)
            insert_after += 1
            set_cell_text(new_row.cells[0], e["name_id"])
            set_cell_text(new_row.cells[1], label)
            set_cell_text(new_row.cells[2], e["code"])
            set_cell_text(new_row.cells[3], "")
            added += 1
            for k in group_last_row:
                if k != cat_key and group_last_row[k] >= insert_after - 1:
                    group_last_row[k] += 1

    # Tidy up: remove fully-empty rows (all cells blank).
    for ri in range(len(table.rows) - 1, 1, -1):
        id_text, cat_text, reason_text, plan_text = _row_texts(table, ri)
        if not id_text.strip() and not (cat_text or "").strip() \
                and not reason_text.strip() and not plan_text.strip():
            delete_row(table, ri)

    return len(keep_keys) + added


def fill(doc_path, xlsx_path, out_path):
    doc = Document(doc_path)
    n = fill_table(doc, xlsx_path)
    doc.save(out_path)
    return n
