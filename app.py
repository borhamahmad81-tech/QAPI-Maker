"""MR Form Filler — drag & drop GUI.

Three categories: Extrasessions, Hospitalization, Outliers.
For each, drop the Word template (.docx) and the Excel data (.xlsx), then click
Fill. The filled document is saved next to the template as *_filled.docx.

Built with tkinterdnd2 for drag & drop. Falls back to click-to-browse if DnD
is unavailable.
"""
import os
import sys
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:
    _HAS_DND = False

# Allow running both as script and frozen exe. PyInstaller --onefile extracts
# bundled data to sys._MEIPASS, so prefer that for imports.
if getattr(sys, "frozen", False):
    BASE = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from core import (fill_extrasessions, fill_hospitalization, fill_outliers,
                  fill_master, compare_docs)
from core.action_plan_templates import load_templates, save_templates
from core.fill_excluded import EXCLUSION_CODE_LEGEND

CATEGORIES = [
    ("Extra Sessions", fill_extrasessions.fill),
    ("Hospitalization", fill_hospitalization.fill),
    ("Outliers", fill_outliers.fill),
]

ACCENT = "#2563eb"
BG = "#f8fafc"
CARD = "#ffffff"
BORDER = "#cbd5e1"
OK = "#16a34a"
MUTED = "#64748b"


def _strip_dnd_path(data):
    """tkinterdnd2 returns paths possibly wrapped in {}; handle one path."""
    data = data.strip()
    if data.startswith("{") and data.endswith("}"):
        data = data[1:-1]
    return data


class DropZone(tk.Frame):
    def __init__(self, master, label, extensions, **kw):
        super().__init__(master, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1, **kw)
        self.extensions = extensions
        self.path = None
        self.label_text = label

        self.title = tk.Label(self, text=label, bg=CARD, fg=MUTED,
                              font=("Segoe UI", 9, "bold"))
        self.title.pack(anchor="w", padx=10, pady=(8, 0))

        self.drop = tk.Label(self, text="Drop file here\nor click to browse",
                             bg="#eef2ff", fg=MUTED, font=("Segoe UI", 9),
                             height=3, cursor="hand2", relief="flat")
        self.drop.pack(fill="x", padx=10, pady=8)
        self.drop.bind("<Button-1>", lambda e: self.browse())

        if _HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self.on_drop)

    def on_drop(self, event):
        path = _strip_dnd_path(event.data)
        self.set_path(path)

    def browse(self):
        ftypes = [("Supported", " ".join("*" + e for e in self.extensions))]
        path = filedialog.askopenfilename(filetypes=ftypes)
        if path:
            self.set_path(path)

    def set_path(self, path):
        if not any(path.lower().endswith(e) for e in self.extensions):
            messagebox.showerror("Wrong file type",
                                 f"{self.label_text} expects: {', '.join(self.extensions)}")
            return
        self.path = path
        self.drop.config(text=os.path.basename(path), bg="#dcfce7", fg=OK)


class CategoryCard(tk.Frame):
    def __init__(self, master, name, fill_func):
        super().__init__(master, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1)
        self.name = name
        self.fill_func = fill_func

        header = tk.Label(self, text=name, bg=ACCENT, fg="white",
                          font=("Segoe UI", 11, "bold"), pady=6)
        header.pack(fill="x")

        self.doc_zone = DropZone(self, "Word template (.docx)", [".docx"])
        self.doc_zone.pack(fill="x", padx=8, pady=(8, 4))

        self.xls_zone = DropZone(self, "Excel data (.xlsx)", [".xlsx", ".xlsm"])
        self.xls_zone.pack(fill="x", padx=8, pady=4)

        self.btn = tk.Button(self, text="Fill Document", bg=ACCENT, fg="white",
                             font=("Segoe UI", 10, "bold"), relief="flat",
                             cursor="hand2", command=self.run)
        self.btn.pack(fill="x", padx=8, pady=8)

        self.status = tk.Label(self, text="", bg=CARD, fg=MUTED,
                               font=("Segoe UI", 8), wraplength=240, justify="left")
        self.status.pack(fill="x", padx=8, pady=(0, 8))

    def run(self):
        doc = self.doc_zone.path
        xls = self.xls_zone.path
        if not doc or not xls:
            messagebox.showwarning("Missing file",
                                   "Please drop both the Word template and the Excel file.")
            return
        out = os.path.splitext(doc)[0] + "_filled.docx"
        try:
            count = self.fill_func(doc, xls, out)
            self.status.config(
                text=f"Done. {count} entries written.\nSaved: {os.path.basename(out)}",
                fg=OK)
            messagebox.showinfo("Success",
                                f"{self.name}: {count} entries filled.\n\nSaved to:\n{out}")
        except Exception as e:
            traceback.print_exc()
            self.status.config(text=f"Error: {e}", fg="#dc2626")
            messagebox.showerror("Error", f"{self.name} failed:\n\n{e}")


class MasterPanel(tk.Frame):
    """Master-document mode: one Doc + QAPI/Hosp/Extra Excel files."""
    def __init__(self, master):
        super().__init__(master, bg=BG)
        intro = ("Drop the master QAPI Word document, then the Excel files you "
                 "have. Only the sections you provide are filled; everything "
                 "else stays untouched.")
        tk.Label(self, text=intro, bg=BG, fg=MUTED, font=("Segoe UI", 9),
                 wraplength=760, justify="left").pack(pady=(4, 10), padx=14, anchor="w")

        grid = tk.Frame(self, bg=BG)
        grid.pack(fill="x", padx=14)

        self.doc_zone = DropZone(grid, "Master Document (.docx)", [".docx"])
        self.qapi_zone = DropZone(grid, "QAPI Excel — Outliers + Excluded + KPI (.xlsx)",
                                  [".xlsx", ".xlsm"])
        self.hosp_zone = DropZone(grid, "Hospitalization Excel (.xlsx)", [".xlsx", ".xlsm"])
        self.extra_zone = DropZone(grid, "Extra Sessions Excel (.xlsx)", [".xlsx", ".xlsm"])
        for i, z in enumerate([self.doc_zone, self.qapi_zone, self.hosp_zone, self.extra_zone]):
            z.grid(row=i // 2, column=i % 2, sticky="ew", padx=6, pady=6)
            grid.grid_columnconfigure(i % 2, weight=1)

        self.btn = tk.Button(self, text="Fill Master Document", bg=ACCENT, fg="white",
                             font=("Segoe UI", 11, "bold"), relief="flat",
                             cursor="hand2", command=self.run)
        self.btn.pack(fill="x", padx=20, pady=(12, 4))

        self.status = tk.Label(self, text="", bg=BG, fg=MUTED,
                               font=("Segoe UI", 9), wraplength=760, justify="left")
        self.status.pack(fill="x", padx=20, pady=(0, 8))

    def run(self):
        doc = self.doc_zone.path
        if not doc:
            messagebox.showwarning("Missing document",
                                   "Please drop the master Word document.")
            return
        if not (self.qapi_zone.path or self.hosp_zone.path or self.extra_zone.path):
            messagebox.showwarning("No data",
                                   "Drop at least one Excel file to fill a section.")
            return
        out = os.path.splitext(doc)[0] + "_filled.docx"
        try:
            res = fill_master.fill_master(
                doc, out,
                qapi_xlsx=self.qapi_zone.path,
                hosp_xlsx=self.hosp_zone.path,
                extra_xlsx=self.extra_zone.path)
            lines = [f"{k}: {v}" for k, v in res.items()]
            self.status.config(text="Done.\n" + "\n".join(lines), fg=OK)
            messagebox.showinfo("Success",
                                "Master document filled.\n\n" + "\n".join(lines)
                                + f"\n\nSaved to:\n{out}")
        except Exception as e:
            traceback.print_exc()
            self.status.config(text=f"Error: {e}", fg="#dc2626")
            messagebox.showerror("Error", f"Failed:\n\n{e}")


class TemplatePanel(tk.Frame):
    """Separate-templates mode: the original three cards."""
    def __init__(self, master):
        super().__init__(master, bg=BG)
        sub = ("Drag & drop each Word template and its Excel sheet, then click Fill."
               if _HAS_DND else
               "Click each box to browse, then click Fill. (Drag & drop unavailable.)")
        tk.Label(self, text=sub, bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(pady=(4, 10))
        container = tk.Frame(self, bg=BG)
        container.pack(fill="both", expand=True, padx=14, pady=4)
        for i, (name, func) in enumerate(CATEGORIES):
            card = CategoryCard(container, name, func)
            card.grid(row=0, column=i, sticky="nsew", padx=6)
            container.grid_columnconfigure(i, weight=1)
        container.grid_rowconfigure(0, weight=1)


class ActionPlanTemplatesPanel(tk.Frame):
    """Lets the person set a standard Action Plan for each exclusion code.
    Saved templates are applied automatically to NEWLY excluded patients only;
    already-present, matched patients are never touched."""
    def __init__(self, master):
        super().__init__(master, bg=BG)
        intro = ("Write a standard Action Plan for each exclusion code. When the "
                 "program adds a newly-excluded patient with that code, this text "
                 "is filled in automatically. Existing patients already in the "
                 "document are never changed.")
        tk.Label(self, text=intro, bg=BG, fg=MUTED, font=("Segoe UI", 9),
                 wraplength=780, justify="left").pack(pady=(4, 10), padx=14, anchor="w")

        # Scrollable area
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True, padx=14)
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        vbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        inner = tk.Frame(canvas, bg=BG)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self.boxes = {}
        saved = load_templates()
        for code, meaning in EXCLUSION_CODE_LEGEND.items():
            row = tk.Frame(inner, bg=CARD, highlightbackground=BORDER,
                           highlightthickness=1)
            row.pack(fill="x", pady=4, padx=2)
            head = tk.Label(row, text=f"{code} — {meaning}", bg=CARD, fg="#0f172a",
                            font=("Segoe UI", 9, "bold"), anchor="w",
                            wraplength=700, justify="left")
            head.pack(fill="x", padx=8, pady=(6, 2))
            box = tk.Text(row, height=2, font=("Segoe UI", 9), wrap="word",
                         relief="flat", bg="#f8fafc", highlightbackground=BORDER,
                         highlightthickness=1)
            box.pack(fill="x", padx=8, pady=(0, 8))
            if code in saved:
                box.insert("1.0", saved[code])
            self.boxes[code] = box

        self.status = tk.Label(self, text="", bg=BG, fg=MUTED, font=("Segoe UI", 9))
        self.status.pack(pady=(4, 0))

        save_btn = tk.Button(self, text="Save Templates", bg=ACCENT, fg="white",
                             font=("Segoe UI", 11, "bold"), relief="flat",
                             cursor="hand2", command=self.save)
        save_btn.pack(fill="x", padx=20, pady=10)

    def save(self):
        templates = {}
        for code, box in self.boxes.items():
            text = box.get("1.0", "end").strip()
            if text:
                templates[code] = text
        ok = save_templates(templates)
        if ok:
            self.status.config(text=f"Saved {len(templates)} template(s).", fg=OK)
        else:
            self.status.config(text="Failed to save — check folder permissions.",
                               fg="#dc2626")


class ComparePanel(tk.Frame):
    """Compare a program-filled document against a manually-prepared one and
    list differences only, across all five sections."""
    def __init__(self, master):
        super().__init__(master, bg=BG)
        intro = ("Drop the program-filled document as Doc A and the manually "
                 "prepared document as Doc B. The report lists only real "
                 "differences \u2014 IDs/categories missing on either side, and "
                 "values that differ by more than a small tolerance. Wording "
                 "of reasons and action plans is never compared.")
        tk.Label(self, text=intro, bg=BG, fg=MUTED, font=("Segoe UI", 9),
                 wraplength=780, justify="left").pack(pady=(4, 10), padx=14, anchor="w")

        grid = tk.Frame(self, bg=BG)
        grid.pack(fill="x", padx=14)
        self.a_zone = DropZone(grid, "Doc A — Program-filled (.docx)", [".docx"])
        self.b_zone = DropZone(grid, "Doc B — Manually prepared (.docx)", [".docx"])
        self.a_zone.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        self.b_zone.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

        self.btn = tk.Button(self, text="Compare Documents", bg=ACCENT, fg="white",
                             font=("Segoe UI", 11, "bold"), relief="flat",
                             cursor="hand2", command=self.run)
        self.btn.pack(fill="x", padx=20, pady=(12, 8))

        self.text = tk.Text(self, font=("Consolas", 9), wrap="word",
                            relief="flat", bg="#f8fafc",
                            highlightbackground=BORDER, highlightthickness=1)
        self.text.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        self.text.insert("1.0", "Comparison results will appear here.")
        self.text.config(state="disabled")

    def run(self):
        a, b = self.a_zone.path, self.b_zone.path
        if not a or not b:
            messagebox.showwarning("Missing file",
                                   "Please drop both documents to compare.")
            return
        try:
            report = compare_docs.compare(a, b)
            text = compare_docs.format_report(report)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Comparison failed:\n\n{e}")
            return
        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", text)
        self.text.config(state="disabled")


def main():
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    root.title("MR Form Filler")
    root.configure(bg=BG)
    root.geometry("860x600")

    tk.Label(root, text="Medical Report Form Filler", bg=BG, fg="#0f172a",
             font=("Segoe UI", 16, "bold")).pack(pady=(14, 2))

    # Mode switch
    mode = tk.StringVar(value="master")
    switch = tk.Frame(root, bg=BG)
    switch.pack(pady=(2, 6))
    panels = {}

    def show(which):
        mode.set(which)
        for name, p in panels.items():
            p.pack_forget()
        panels[which].pack(fill="both", expand=True)
        for name, b in buttons.items():
            b.config(bg=ACCENT if name == which else "#e2e8f0",
                     fg="white" if name == which else "#334155")

    buttons = {}
    for key, label in [("master", "Master Document"),
                       ("templates", "Separate Templates"),
                       ("action_plans", "Action Plan Templates"),
                       ("compare", "Compare Documents")]:
        b = tk.Button(switch, text=label, font=("Segoe UI", 10, "bold"),
                      relief="flat", cursor="hand2", padx=16, pady=4,
                      command=lambda k=key: show(k))
        b.pack(side="left", padx=4)
        buttons[key] = b

    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True)
    panels["master"] = MasterPanel(body)
    panels["templates"] = TemplatePanel(body)
    panels["action_plans"] = ActionPlanTemplatesPanel(body)
    panels["compare"] = ComparePanel(body)

    tk.Label(root, text="Patient IDs are written without the MOH prefix.",
             bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(side="bottom", pady=8)

    show("master")
    root.mainloop()


if __name__ == "__main__":
    main()
