"""Shared helpers: Excel reading, date formatting, ID cleaning."""
import datetime
import openpyxl


def read_sheet(path):
    """Return list of rows (tuples) from the first worksheet, values only."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    return [row for row in ws.iter_rows(values_only=True)]


def fmt_date(value):
    """Render a cell that represents a date as D/M/YYYY.

    Excel often mis-parses day/month text into datetime objects. We always
    output day/month/year to match the clinic's filled examples.
    Text cells (already strings) are returned cleaned.
    """
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        day, month, year = value.day, value.month, value.year
        # Excel often swaps day/month when auto-converting typed dates.
        # When the day component is <= 12 the value is ambiguous, so Excel may
        # have stored month-as-day. Swap to recover the intended day/month.
        # When day > 12 the order is unambiguous and we keep it as-is.
        if day <= 12 and month <= 12:
            day, month = month, day
        return f"{day}/{month}/{year}"
    return str(value).strip()


def clean_id(value):
    """Strip a leading MOH prefix and surrounding whitespace from an ID."""
    if value is None:
        return ""
    s = str(value).strip()
    # Drop trailing '.0' that openpyxl sometimes adds to integer-like floats
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    up = s.upper()
    if up.startswith("MOH"):
        s = s[3:].strip()
    return s


def clean_text(value):
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def is_formula_or_blank(value):
    """True if cell is empty or an Excel formula reference (e.g. =$A$7)."""
    if value is None:
        return True
    s = str(value).strip()
    return s == "" or s.startswith("=")
