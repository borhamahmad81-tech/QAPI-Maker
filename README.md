# MR Form Filler

Fills three medical-report Word templates from Excel data **without changing
fonts, formatting, or layout**. Patient IDs are written **without the MOH prefix**.

## How to use

1. Open **MR_FormFiller.exe**.
2. For each category, drag the Word template into the top box and the Excel
   file into the bottom box (or click a box to browse).
3. Click **Fill Document**.
4. The result is saved next to the template as `<name>_filled.docx`.

The program finds each table by its content, so tables may be moved within the
document. Keep the Excel **column headers** stable.

## Extra Sessions

Excel columns: `Name`, `ID`, `No. of sessions`, `Date`, `Cause`,
`Date of MOH approval`.

- Patients are identified by **ID**, not name. All rows with the same ID become
  **one row** in the document.
- If an ID appears under slightly different names (Arabic spelling variants),
  the **first name found** is used for all of that patient's sessions.
- Output per patient:
  - Patient ID cell: name on line 1, ID on line 2.
  - Sessions cell: "N sessions" on line 1, all dates (comma-separated) on line 2.
  - Reason cell: all reasons, identical ones removed.
  - Notes cell: all MOH approval dates, identical ones removed (one approval
    covering several sessions is written once).

## Hospitalization

Excel columns: `Name`, `ID`, `Hospital`, `Admission Date`, `Discharge Date`,
`Cause`, `blood Transfusion`.

- Only the "Hospitalized Patients" cell is filled, as a numbered list.
- The count cell is auto-filled with the number of entries.
- Each row is a separate admission (second admissions listed separately).
- Per entry: `Name: ID`, then Date admitted, Cause, Hospital.
- If a discharge date exists: `Outcome: Discharged on <date>` and
  `Received in stable condition`.
- The `Blood transfusion` line appears only when the Excel cell has a value.
- Special outcomes (death, transfer details, serology notes, etc.) are clinical
  notes you add manually.

## Outliers

Keep the block layout with `Name_ID`, `ID`, and a value column per KPI.
Recognised value headers: `Hb_g/dl`, `Phosph_mg/dl`, `Calcium`, `Access-can`,
`Kt/V`, `Albumin_g/l`, `Sum of BFR`, `Average of MAP`, and (when present)
`PTH_pg/ml`.

- The `Name_ID` cell reads `Surname, First (ID)` and is copied verbatim.
- Only the "Patient /Reason" column is filled, as `Name_ID: value.`
  The Action plan column is left blank for manual entry.
- The Count column is filled and merged per KPI with the number of patients.
- Thresholds: Hb <10 / >12, Phosphorus <2.5 / >5.5, Calcium <8.42 / >10,
  Kt/V <1.4, Albumin ≤3.5 g/dl, BFR <350, MAP >105, PTH <150 / >600.
- Albumin is stored as g/L in the sheet and reported as g/dl (divided by 10).
- If the template has no BFR row but the Excel has BFR outliers, a BFR row is
  added after Kt/V.
- PTH rows fill automatically once a `PTH_pg/ml` block is added to the sheet;
  if absent, those rows stay blank.

## A note on dates

Excel sometimes swaps day/month when auto-converting a typed date. For real
date cells where the day is 12 or less (ambiguous), the program swaps them back
to day/month order. Dates typed as text (e.g. `20/5/2026`) are used as-is.

## Building the EXE on GitHub

1. Create a new GitHub repository and upload the contents of this folder
   (including the `.github` folder and `core` folder).
2. Go to the **Actions** tab; the workflow runs automatically on push, or run
   it manually with **Run workflow**.
3. When it finishes, open the run and download the **MR_FormFiller-windows**
   artifact. Inside is `MR_FormFiller.exe`.

No admin rights or local Python install are needed.
