"""Master QAPI document filler.

Processes ONE master Word document containing many sections, filling only the
sections for which an Excel file is supplied. Everything else is left untouched.

Section -> behavior:
  Hospitalization (Patients Census)        : fresh replace      <- hosp xlsx
  Extra HD sessions                        : fresh replace      <- extra xlsx
  Outliers (Details of patients...)        : reconcile by ID    <- qapi xlsx
  Details of Excluded Patients             : reconcile ID+KPI   <- qapi xlsx
  MOH KPI (Medical)                        : fresh replace      <- qapi xlsx
  MOH KPI (Operational) and all others     : untouched

Each section is found inside the document by its heading text (handled within
the individual fillers' table-finder functions), so section order/position in
the document does not matter.
"""
from docx import Document

from . import fill_hospitalization
from . import fill_extrasessions
from . import fill_outliers
from . import fill_excluded
from . import fill_kpi_medical


def fill_master(doc_path, out_path, qapi_xlsx=None, hosp_xlsx=None,
                extra_xlsx=None):
    """Fill the supplied sections in the master document.

    Any of the xlsx arguments may be None -> that section is skipped and left
    untouched. Returns a dict of section -> result (count or error string).
    """
    doc = Document(doc_path)
    results = {}

    if hosp_xlsx:
        try:
            results["Hospitalization"] = fill_hospitalization.fill_doc(doc, hosp_xlsx)
        except Exception as e:
            results["Hospitalization"] = f"error: {e}"

    if extra_xlsx:
        try:
            results["Extra Sessions"] = fill_extrasessions.fill_doc(doc, extra_xlsx)
        except Exception as e:
            results["Extra Sessions"] = f"error: {e}"

    if qapi_xlsx:
        # KPI Medical (fresh replace)
        try:
            results["KPI (Medical)"] = fill_kpi_medical.fill_table(doc, qapi_xlsx)
        except Exception as e:
            results["KPI (Medical)"] = f"error: {e}"
        # Excluded (reconcile by ID + KPI)
        try:
            results["Excluded Patients"] = fill_excluded.fill_table(doc, qapi_xlsx)
        except Exception as e:
            results["Excluded Patients"] = f"error: {e}"
        # Outliers (reconcile by ID) - do LAST as it does heavy row merging
        try:
            results["Outliers"] = fill_outliers.fill_doc(doc, qapi_xlsx)
        except Exception as e:
            results["Outliers"] = f"error: {e}"

    doc.save(out_path)
    return results
