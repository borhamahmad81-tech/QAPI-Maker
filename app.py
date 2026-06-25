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

from core import fill_extrasessions, fill_hospitalization, fill_outliers

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


def main():
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    root.title("MR Form Filler")
    root.configure(bg=BG)
    root.geometry("820x520")

    tk.Label(root, text="Medical Report Form Filler", bg=BG, fg="#0f172a",
             font=("Segoe UI", 16, "bold")).pack(pady=(14, 2))
    sub = "Drag & drop each Word template and its Excel sheet, then click Fill."
    if not _HAS_DND:
        sub = "Click each box to browse for files, then click Fill. (Drag & drop unavailable.)"
    tk.Label(root, text=sub, bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(pady=(0, 12))

    container = tk.Frame(root, bg=BG)
    container.pack(fill="both", expand=True, padx=14, pady=4)

    for i, (name, func) in enumerate(CATEGORIES):
        card = CategoryCard(container, name, func)
        card.grid(row=0, column=i, sticky="nsew", padx=6)
        container.grid_columnconfigure(i, weight=1)
    container.grid_rowconfigure(0, weight=1)

    tk.Label(root, text="Patient IDs are written without the MOH prefix. "
                        "Action plans in Outliers are left blank for manual entry.",
             bg=BG, fg=MUTED, font=("Segoe UI", 8)).pack(side="bottom", pady=8)

    root.mainloop()


if __name__ == "__main__":
    main()
