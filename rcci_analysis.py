"""RCCI combustion and engine performance analysis GUI.

Run with: python rcci_analysis.py
Dependencies: openpyxl, xlrd, python-docx (Tkinter is included with Python).
"""

from __future__ import annotations

import math
import re
import csv
import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import Tk, StringVar, filedialog, messagebox, ttk

import openpyxl
import xlrd
from docx import Document


# The user's 15-point design matrix: (load %, SOI °bTDC).
DOE = [(load, soi) for load in (0, 25, 50, 75, 100) for soi in (12, 15, 18)]


@dataclass
class RunResult:
    test: int
    load: int
    soi: int
    folder: str
    rpm: float | None = None
    torque: float | None = None
    brake_power: float | None = None
    bsfc: float | None = None
    bte: float | None = None
    mech_eff: float | None = None
    ca10: float | None = None
    ca50: float | None = None
    ca90: float | None = None
    ignition_delay: float | None = None
    peak_pressure: float | None = None
    peak_pressure_ca: float | None = None
    mprr: float | None = None
    hrr_peak: float | None = None
    hrr_peak_ca: float | None = None
    emissions: str = "Not supplied"
    warnings: str = ""


def number(value):
    try:
        v = float(value)
        return v if math.isfinite(v) else None
    except (TypeError, ValueError):
        return None


def locate(folder: Path, patterns):
    for pattern in patterns:
        hits = sorted(folder.glob(pattern))
        if hits:
            return hits[0]
    return None


def ca_at_fraction(ca, mfb, fraction):
    """Linear interpolation of crank angle at a normalized MFB fraction."""
    pairs = [(x, y) for x, y in zip(ca, mfb) if x is not None and y is not None]
    if len(pairs) < 2:
        return None
    pairs.sort()
    # Normalize the MFB trace to its observed min/max so offsets are harmless.
    lo, hi = min(y for _, y in pairs), max(y for _, y in pairs)
    if hi - lo <= 1e-12:
        return None
    vals = [(x, (y - lo) / (hi - lo)) for x, y in pairs]
    for (x0, y0), (x1, y1) in zip(vals, vals[1:]):
        if y0 <= fraction <= y1 and y1 > y0:
            return x0 + (fraction - y0) * (x1 - x0) / (y1 - y0)
    return None


def read_combustion(book_path: Path, cycle_path: Path):
    wb = openpyxl.load_workbook(book_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row = None
    headers = {}
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=50, values_only=True), 1):
        normalized = {str(v).strip().lower(): i for i, v in enumerate(row) if v is not None}
        if any("angle" in key for key in normalized) and any("pressure" in key for key in normalized):
            header_row, headers = row_idx, normalized
            break
    if header_row is None:
        raise ValueError(f"Could not find crank-angle and pressure columns in {book_path.name}")
    def col(*terms):
        for term in terms:
            for key, idx in headers.items():
                if term in key:
                    return idx
        return None
    ca_i, vol_i, mfb_i, nhr_i = col("angle"), col("volume"), col("mfb"), col("nhr")
    if ca_i is None:
        raise ValueError("Crank-angle column missing")
    ca, volume, mfb, nhr = [], [], [], []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        x = number(row[ca_i] if ca_i < len(row) else None)
        if x is None:
            continue
        ca.append(x)
        volume.append(number(row[vol_i]) if vol_i is not None else None)
        mfb.append(number(row[mfb_i]) if mfb_i is not None else None)
        nhr.append(number(row[nhr_i]) if nhr_i is not None else None)
    wb.close()

    raw = xlrd.open_workbook(str(cycle_path), on_demand=True)
    sheet = raw.sheet_by_index(0)
    # Cycle sheet has 10 columns, one pressure trace per cycle and no angle column.
    cycle_rows = []
    for ri in range(1, sheet.nrows):
        vals = [number(sheet.cell_value(ri, ci)) for ci in range(sheet.ncols)]
        vals = [v for v in vals if v is not None]
        cycle_rows.append(sum(vals) / len(vals) if vals else None)
    raw.release_resources()
    if len(cycle_rows) != len(ca):
        raise ValueError(f"Pressure samples ({len(cycle_rows)}) do not match crank-angle samples ({len(ca)})")

    pressure = cycle_rows
    good = [(x, p) for x, p in zip(ca, pressure) if p is not None]
    peak_x, peak_p = max(good, key=lambda pair: pair[1]) if good else (None, None)
    rates = []
    for (x0, p0), (x1, p1) in zip(good, good[1:]):
        if x1 != x0:
            rates.append((p1 - p0) / (x1 - x0))
    peak_rate = max(rates) if rates else None
    ca10, ca50, ca90 = (ca_at_fraction(ca, mfb, f) for f in (.1, .5, .9))
    hrr = [(x, y) for x, y in zip(ca, nhr) if x is not None and y is not None]
    hrr_peak = max(hrr, key=lambda pair: pair[1]) if hrr else (None, None)
    return dict(ca10=ca10, ca50=ca50, ca90=ca90, peak_pressure=peak_p,
                peak_pressure_ca=peak_x, mprr=peak_rate,
                hrr_peak=hrr_peak[1], hrr_peak_ca=hrr_peak[0])


def read_performance(path: Path):
    book = xlrd.open_workbook(str(path), on_demand=True)
    sh = book.sheet_by_index(0)
    if sh.nrows < 2:
        raise ValueError("Performance report is empty")
    headers = [str(x).strip().lower() for x in sh.row_values(0)]
    vals = sh.row_values(1)
    data = {h: vals[i] for i, h in enumerate(headers) if i < len(vals)}
    def find(*words):
        return next((number(v) for k, v in data.items() if all(w in k for w in words) and number(v) is not None), None)
    rpm, torque = find("rpm"), find("torque")
    # Recompute brake power from torque and speed when possible; otherwise use the report value.
    bhp = find("bhp")
    brake_power = 2 * math.pi * rpm * torque / 60000 if rpm is not None and torque is not None else bhp
    bsfc = find("sfc")
    bte = find("b.t.effi")
    mech = find("mech.effi")
    book.release_resources()
    return dict(rpm=rpm, torque=torque, brake_power=brake_power, bsfc=bsfc, bte=bte, mech_eff=mech)


def read_emissions(path: Path):
    """Read a header-row analyzer export; retain header units in the report."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as stream:
            rows = list(csv.reader(stream))
    elif suffix == ".xlsx":
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rows = list(book.active.iter_rows(values_only=True)); book.close()
    elif suffix == ".xls":
        book = xlrd.open_workbook(str(path), on_demand=True)
        sh = book.sheet_by_index(0)
        rows = [sh.row_values(i) for i in range(sh.nrows)]
        book.release_resources()
    else:
        return "File supplied; unsupported extension"
    for i, header_row in enumerate(rows[:30]):
        headers = [str(v).strip() if v is not None else "" for v in header_row]
        joined = " ".join(headers).lower()
        if not any(x in joined for x in ("co", "co2", "nox", "hc", "o2", "smoke")):
            continue
        for data in rows[i + 1:]:
            if not data or not any(str(v).strip() for v in data if v is not None):
                continue
            values = []
            for h, v in zip(headers, data):
                if h and number(v) is not None and re.search(r"\b(co2?|nox|hc|o2|smoke|lambda|opacity|soot)\b", h, re.I):
                    values.append(f"{h}: {number(v):g}")
            if values:
                return "; ".join(values)
    return "File supplied, but no pollutant header/data row was recognized"


def analyze_folder(folder: Path, index: int):
    test = index if 1 <= index <= len(DOE) else 1
    load, soi = DOE[test - 1]
    result = RunResult(test, load, soi, str(folder))
    book = locate(folder, ["Book*.xlsx"])
    cycle = locate(folder, ["Sat*.xls", "*Cycle*.xls"])
    perf = locate(folder, ["Performance Report.xls", "Performance*.xlsx"])
    if not (book and cycle and perf):
        missing = [name for name, path in (("Book*.xlsx", book), ("Sat*.xls", cycle), ("Performance Report.xls", perf)) if path is None]
        raise ValueError("Missing " + ", ".join(missing))
    for key, value in read_combustion(book, cycle).items():
        setattr(result, key, value)
    for key, value in read_performance(perf).items():
        setattr(result, key, value)
    if result.ca10 is not None:
        result.ignition_delay = result.ca10 + result.soi
    emission_file = locate(folder, ["*emission*.csv", "*exhaust*.csv", "*emission*.xlsx", "*exhaust*.xlsx",
                                    "*emission*.xls", "*exhaust*.xls"])
    if emission_file:
        result.emissions = read_emissions(emission_file)
    else:
        result.warnings = "No exhaust analyzer data file found"
    return result


def fmt(value, digits=3):
    return "—" if value is None else f"{value:.{digits}f}"


def write_report(results, target: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    headers = ["Test", "Load (%)", "SOI (°bTDC)", "RPM", "Torque (Nm)", "Brake power (kW)",
               "BSFC (kg/kWh)", "BTE (%)", "Mechanical efficiency (%)", "CA10 (°CA)", "CA50 (°CA)",
               "CA90 (°CA)", "Ignition delay (°CA)", "Peak pressure (bar)", "Peak pressure CA (°CA)",
               "MPRR (bar/°CA)", "Peak NHR (J/°CA)", "Peak NHR CA (°CA)", "Emissions", "Notes"]
    ws.append(headers)
    for r in results:
        ws.append([r.test, r.load, r.soi, r.rpm, r.torque, r.brake_power, r.bsfc, r.bte, r.mech_eff,
                   r.ca10, r.ca50, r.ca90, r.ignition_delay, r.peak_pressure, r.peak_pressure_ca,
                   r.mprr, r.hrr_peak, r.hrr_peak_ca, r.emissions, r.warnings])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill("solid", fgColor="17365D")
    for column in ws.columns:
        letter = column[0].column_letter
        ws.column_dimensions[letter].width = min(34, max(13, max(len(str(c.value or "")) for c in column) + 2))
    wb.save(target)

    doc = Document()
    doc.add_heading("RCCI Combustion and Performance Analysis", 0)
    doc.add_paragraph("Engine: Mahindra Jeeto single-cylinder, four-stroke diesel converted to RCCI; 670 cc; CR 18:1.")
    doc.add_paragraph("Fuel setup: E20 low-reactivity fuel at 30% energy share and diesel high-reactivity fuel.")
    doc.add_paragraph("CA10/50/90 are interpolated from the supplied MFB trace. Ignition delay is CA10 − SOI, with SOI entered as °bTDC and represented as a negative crank angle. MPRR is the maximum positive slope of cycle-averaged pressure. Peak NHR uses the supplied NHR column.")
    doc.add_paragraph("Performance values are read from the engine performance report, except brake power, recomputed from torque and RPM when available. Emissions are not calculated unless analyzer readings are supplied in a compatible file.")
    table = doc.add_table(rows=1, cols=9)
    table.style = "Light Shading Accent 1"
    labels = ["Test", "Load %", "SOI", "CA10", "CA50", "CA90", "ID", "Pmax bar", "MPRR bar/°"]
    for cell, label in zip(table.rows[0].cells, labels): cell.text = label
    for r in results:
        row = table.add_row().cells
        vals = [r.test, r.load, r.soi, fmt(r.ca10), fmt(r.ca50), fmt(r.ca90), fmt(r.ignition_delay), fmt(r.peak_pressure), fmt(r.mprr)]
        for cell, value in zip(row, vals): cell.text = str(value)
    doc.add_heading("Engine performance", level=1)
    perf_table = doc.add_table(rows=1, cols=7)
    perf_table.style = "Light Shading Accent 1"
    for cell, label in zip(perf_table.rows[0].cells,
                           ["Test", "RPM", "Torque Nm", "Brake kW", "BSFC kg/kWh", "BTE %", "Mechanical %"]):
        cell.text = label
    for r in results:
        row = perf_table.add_row().cells
        for cell, value in zip(row, [r.test, fmt(r.rpm, 0), fmt(r.torque), fmt(r.brake_power), fmt(r.bsfc), fmt(r.bte), fmt(r.mech_eff)]):
            cell.text = str(value)
    doc.add_heading("Exhaust analyzer readings", level=1)
    for r in results:
        if r.emissions != "Not supplied":
            doc.add_paragraph(f"Test {r.test}: {r.emissions}", style="List Bullet")
    if all(r.emissions == "Not supplied" for r in results):
        doc.add_paragraph("No exhaust analyzer files were supplied; emissions could not be calculated.")
    doc.add_heading("Notes", level=1)
    for r in results:
        if r.warnings:
            doc.add_paragraph(f"Test {r.test}: {r.warnings}", style="List Bullet")
    doc.save(target.with_suffix(".docx"))


class App:
    columns = ["Test", "Load %", "SOI", "RPM", "Torque Nm", "Brake kW", "BSFC", "BTE %", "Mech. %",
               "CA10", "CA50", "CA90", "ID", "Pmax bar", "MPRR bar/°", "Peak NHR J/°"]
    attrs = ["test", "load", "soi", "rpm", "torque", "brake_power", "bsfc", "bte", "mech_eff",
             "ca10", "ca50", "ca90", "ignition_delay", "peak_pressure", "mprr", "hrr_peak"]

    def __init__(self, root):
        self.root = root
        root.title("RCCI Combustion & Emission Analysis")
        root.geometry("1220x700")
        self.folder = StringVar()
        top = ttk.Frame(root, padding=12); top.pack(fill="x")
        ttk.Label(top, text="Folder containing the 15 numbered run folders:").pack(side="left")
        ttk.Entry(top, textvariable=self.folder, width=75).pack(side="left", padx=8)
        ttk.Button(top, text="Browse…", command=self.choose).pack(side="left")
        ttk.Button(top, text="Analyze runs", command=self.analyze).pack(side="left", padx=8)
        self.status = StringVar(value="Choose the RCCI folder that contains subfolders 1–15.")
        ttk.Label(root, textvariable=self.status, padding=(12, 0)).pack(fill="x")
        frame = ttk.Frame(root, padding=12); frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(frame, columns=self.columns, show="headings", height=18)
        for c in self.columns:
            self.tree.heading(c, text=c); self.tree.column(c, width=79, anchor="center")
        self.tree.column("Test", width=48); self.tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview); scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<<TreeviewSelect>>", self.show_detail)
        self.detail = StringVar(value="Select a result row to see the run source and notes.")
        ttk.Label(root, textvariable=self.detail, padding=12, justify="left", wraplength=1150).pack(fill="x")
        bottom = ttk.Frame(root, padding=12); bottom.pack(fill="x")
        ttk.Button(bottom, text="Export Excel + Word report…", command=self.export).pack(side="left")
        ttk.Label(bottom, text="Optional exhaust input: place a CSV/XLS/XLSX analyzer export named *emission* or *exhaust* in its run folder.").pack(side="left", padx=12)
        self.results = []

    def choose(self):
        path = filedialog.askdirectory(title="Select the RCCI run-data folder")
        if path: self.folder.set(path)

    def analyze(self):
        base = Path(self.folder.get())
        if not base.is_dir():
            messagebox.showerror("Folder required", "Choose a valid RCCI run-data folder first."); return
        self.status.set("Reading run files…")
        threading.Thread(target=self._analyze, args=(base,), daemon=True).start()

    def _analyze(self, base):
        results, errors = [], []
        for i in range(1, 16):
            folder = base / str(i)
            if not folder.is_dir():
                errors.append(f"Run folder {i} missing"); continue
            try: results.append(analyze_folder(folder, i))
            except Exception as exc: errors.append(f"Run {i}: {exc}")
        self.root.after(0, lambda: self._show_results(results, errors))

    def _show_results(self, results, errors):
        self.results = results
        self.tree.delete(*self.tree.get_children())
        for r in results:
            self.tree.insert("", "end", values=[r.test, r.load, r.soi, fmt(r.rpm, 0), fmt(r.torque), fmt(r.brake_power),
                fmt(r.bsfc), fmt(r.bte), fmt(r.mech_eff), fmt(r.ca10), fmt(r.ca50), fmt(r.ca90), fmt(r.ignition_delay),
                fmt(r.peak_pressure), fmt(r.mprr), fmt(r.hrr_peak)])
        self.status.set(f"Analyzed {len(results)} of 15 runs; {len(errors)} issue(s)." + (" See details / folder contents." if errors else ""))
        if errors: messagebox.showwarning("Some runs need attention", "\n".join(errors[:15]))

    def show_detail(self, _event):
        selected = self.tree.selection()
        if not selected: return
        index = int(self.tree.item(selected[0], "values")[0]) - 1
        r = next((x for x in self.results if x.test == index + 1), None)
        if r:
            self.detail.set(f"Test {r.test} • {r.folder}\nPeak pressure at {fmt(r.peak_pressure_ca)}°CA; peak NHR at {fmt(r.hrr_peak_ca)}°CA. {r.emissions}. {r.warnings}")

    def export(self):
        if not self.results:
            messagebox.showinfo("No results", "Analyze at least one run first."); return
        path = filedialog.asksaveasfilename(title="Save analysis report", defaultextension=".xlsx",
                    filetypes=[("Excel report", "*.xlsx")], initialfile="RCCI_Analysis_Report.xlsx")
        if not path: return
        try:
            target = Path(path); write_report(self.results, target)
            messagebox.showinfo("Report saved", f"Excel: {target}\nWord: {target.with_suffix('.docx')}")
        except Exception as exc: messagebox.showerror("Export failed", str(exc))


if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()
