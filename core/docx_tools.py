"""Low-level docx helpers that PRESERVE formatting.

Strategy:
- To write text into a cell without disturbing fonts, we keep the first run's
  formatting and replace only its text, deleting extra runs.
- To add table rows, we deep-copy an existing data row's XML (carrying borders,
  shading, cell widths, font of the template) and then overwrite the text.
"""
import copy
from docx.oxml.ns import qn


def _first_paragraph(cell):
    return cell.paragraphs[0]


def set_cell_text(cell, text, keep_font_from_run=True):
    """Replace a cell's entire text with `text`, preserving the font of the
    first existing run. Supports multi-line text via line breaks."""
    text = "" if text is None else str(text)
    para = _first_paragraph(cell)

    # Remove all paragraphs after the first (keep formatting anchor in p[0])
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)

    runs = para.runs
    if runs:
        template_run = runs[0]
        # delete trailing runs, keep the first as our formatting template
        for r in runs[1:]:
            r._element.getparent().remove(r._element)
    else:
        template_run = para.add_run("")

    # Write text into template_run, converting \n into <w:br/>
    lines = text.split("\n")
    template_run.text = lines[0]
    for line in lines[1:]:
        br = template_run._element.makeelement(qn("w:br"), {})
        template_run._element.append(br)
        t = template_run._element.makeelement(qn("w:t"), {})
        t.set(qn("xml:space"), "preserve")
        t.text = line
        template_run._element.append(t)


def clone_row(table, src_row_idx):
    """Deep-copy the row at src_row_idx, append it to the table, return new row."""
    src_tr = table.rows[src_row_idx]._tr
    new_tr = copy.deepcopy(src_tr)
    src_tr.addnext(new_tr)
    # python-docx rebuilds .rows lazily from XML; fetch the newly added one
    return table.rows[src_row_idx + 1]


def insert_row_after(table, src_row_idx):
    """Clone formatting of src row, insert directly after it, blank the text."""
    new_row = clone_row(table, src_row_idx)
    for cell in new_row.cells:
        set_cell_text(cell, "")
    return new_row
