"""Audit every count and amount in the main workbook (run after recalculation).

Checks, per month sheet:
  * every visit / radiology row is inside its TOTAL formula range (nothing left out of the totals)
  * every Bill Amount is a number (not text) or blank; nothing typed outside the data rows
  * Sl No runs 1..n without gaps
  * the recalculated totals equal an independent Python recount/resum of the rows
  * combined rows = visits + radiology
Summary: each month row links to the right cells and equals the sheet; grand totals = sum of months.

Usage: python scripts/audit_totals.py <master.xlsx>   (exit code 1 if anything is wrong)
"""
import re
import sys
from datetime import datetime

import openpyxl

TOL = 0.005


def rng(f):
    m = re.fullmatch(r"=(COUNTA|SUM)\(([A-Z])(\d+):([A-Z])(\d+)\)", str(f))
    return (m.group(1), m.group(2), int(m.group(3)), int(m.group(5))) if m else None


def main(path):
    wf = openpyxl.load_workbook(path)
    wv = openpyxl.load_workbook(path, data_only=True)
    errs, report = [], []
    months = wf.sheetnames[1:]
    per_month = {}
    for name in months:
        f, v = wf[name], wv[name]
        rad_title = next(r for r in range(1, f.max_row + 1) if f.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS")
        lab = lambda col, text: next(r for r in range(1, f.max_row + 1)  # noqa: E731
                                     if str(f[f"{col}{r}"].value or "").startswith(text))
        t_vis, t_bill = lab("A", "TOTAL PATIENTS (VISITS)"), lab("A", "TOTAL BILL AMOUNT")
        c_cnt, c_bill = lab("A", "COMBINED TOTAL PATIENTS"), lab("A", "COMBINED TOTAL BILL")
        r_cnt, r_bill = lab("D", "Total Referrals Count"), lab("D", "Total Bill Amount")
        data = lambda a, b: [r for r in range(a, b) if isinstance(f.cell(r, 2).value, datetime)]  # noqa: E731
        vis = data(4, t_vis)
        rad = data(rad_title + 2, r_cnt)
        # 1. formula ranges cover every row
        for cell, kind, col, rows, slot_end in (
                (f.cell(t_vis, 5), "COUNTA", "B", vis, t_vis), (f.cell(t_bill, 6), "SUM", "F", vis, t_vis),
                (f.cell(r_cnt, 5), "COUNTA", "C", rad, r_cnt), (f.cell(r_bill, 6), "SUM", "F", rad, r_cnt)):
            p = rng(cell.value)
            if not p or p[0] != kind or p[1] != col:
                errs.append(f"{name} {cell.coordinate}: unexpected formula {cell.value}")
            elif rows and (p[2] > rows[0] or p[3] < rows[-1]):
                errs.append(f"{name} {cell.coordinate}: {cell.value} misses rows {rows[0]}-{rows[-1]}")
            elif p[3] >= slot_end:
                errs.append(f"{name} {cell.coordinate}: range runs into the total rows")
        if f.cell(c_cnt, 5).value != f"=E{t_vis}+E{r_cnt}":
            errs.append(f"{name} combined count formula {f.cell(c_cnt, 5).value}")
        if f.cell(c_bill, 6).value != f"=F{t_bill}+F{r_bill}":
            errs.append(f"{name} combined bill formula {f.cell(c_bill, 6).value}")
        # 2. amounts are numbers; Sl numbering
        for part, rows in (("visit", vis), ("radiology", rad)):
            for r in rows:
                a = f.cell(r, 6).value
                if a is not None and not isinstance(a, (int, float)):
                    errs.append(f"{name} F{r} ({part}) is text: {a!r}")
            sl = [f.cell(r, 1).value for r in rows]
            if sl != list(range(1, len(rows) + 1)):
                errs.append(f"{name} {part} Sl No not 1..{len(rows)}")
        for r in list(range(4, t_vis)) + list(range(rad_title + 2, r_cnt)):
            if r not in vis + rad and f.cell(r, 6).value not in (None, ""):
                errs.append(f"{name} F{r} has an amount outside the data rows")
        # 3. independent recount
        nv, sv = len(vis), round(sum(f.cell(r, 6).value or 0 for r in vis), 2)
        nr, sr = len(rad), round(sum(f.cell(r, 6).value or 0 for r in rad), 2)
        got = (v.cell(t_vis, 5).value, v.cell(t_bill, 6).value, v.cell(r_cnt, 5).value, v.cell(r_bill, 6).value,
               v.cell(c_cnt, 5).value, v.cell(c_bill, 6).value)
        exp = (nv, sv, nr, sr, nv + nr, round(sv + sr, 2))
        if any(g is None or abs(g - e) > TOL for g, e in zip(got, exp)):
            errs.append(f"{name}: sheet totals {got} != recount {exp}")
        pend = (sum(f.cell(r, 6).value is None for r in vis), sum(f.cell(r, 6).value is None for r in rad))
        per_month[name] = (exp, (t_vis, t_bill, r_cnt, r_bill))
        report.append(f"{name}: visits {nv} = AED {sv:,.2f} | radiology {nr} = AED {sr:,.2f} | "
                      f"total {nv + nr} = AED {sv + sr:,.2f} | pending {pend[0]} visit, {pend[1]} radiology")
    # 4. Summary
    s, sv_ = wf["Summary"], wv["Summary"]
    grand = [0, 0, 0, 0, 0, 0]
    for i, name in enumerate(months):
        r = 4 + i
        (nv, sv, nr, sr, nt, st), (t_vis, t_bill, r_cnt, r_bill) = per_month[name]
        want = {3: f"='{name}'!E{t_vis}", 4: f"='{name}'!F{t_bill}", 5: f"='{name}'!E{r_cnt}",
                6: f"='{name}'!F{r_bill}", 7: f"=C{r}+E{r}", 8: f"=D{r}+F{r}"}
        for c, w in want.items():
            if s.cell(r, c).value != w:
                errs.append(f"Summary row {r} ({name}) col {c}: {s.cell(r, c).value} != {w}")
        vals = (nv, sv, nr, sr, nt, st)
        for c, e in zip(range(3, 9), vals):
            if abs((sv_.cell(r, c).value or 0) - e) > TOL:
                errs.append(f"Summary row {r} ({name}) col {c}: {sv_.cell(r, c).value} != {e}")
        grand = [g + x for g, x in zip(grand, vals)]
    tot = 4 + len(months) + 1
    for c, e in zip(range(3, 9), grand):
        if s.cell(tot, c).value != f"=SUM({'CDEFGH'[c - 3]}4:{'CDEFGH'[c - 3]}{3 + len(months)})":
            errs.append(f"Summary grand total col {c} formula {s.cell(tot, c).value}")
        if abs((sv_.cell(tot, c).value or 0) - e) > TOL:
            errs.append(f"Summary grand total col {c}: {sv_.cell(tot, c).value} != {round(e, 2)}")
    print("\n".join(report))
    print(f"GRAND: visits {grand[0]} = AED {grand[1]:,.2f} | radiology {grand[2]} = AED {grand[3]:,.2f} | "
          f"total {grand[4]} = AED {grand[5]:,.2f}")
    print("AUDIT:", "OK - no problems" if not errs else f"{len(errs)} PROBLEM(S)")
    for e in errs:
        print("  -", e)
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
