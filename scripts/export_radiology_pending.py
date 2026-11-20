"""Export a separate workbook with one radiology sheet per month that still has pending bill amounts.

Each sheet is a copy of that month's 'RADIOLOGY DATA - OUT PATIENTS' section from the master
(same design), with every blank Bill Amount cell highlighted as a fill-in cell.

Usage: python scripts/export_radiology_pending.py <master.xlsx> <output.xlsx>
"""
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, PatternFill

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402

PENDING_FILL = PatternFill("solid", fgColor=ds.YELLOW)


def copy_cell(src, dst, value=None):
    dst.value = src.value if value is None else value
    dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
    dst.alignment, dst.number_format = copy(src.alignment), src.number_format


def radiology_block(ws):
    title = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS")
    rows = [r for r in range(title + 2, ws.max_row + 1)
            if isinstance(ws.cell(r, 1).value, int) and isinstance(ws.cell(r, 2).value, datetime)]
    return title, rows


def main(master, out):
    src = openpyxl.load_workbook(master)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    summary = []
    for ws in src.worksheets[1:]:
        title, rows = radiology_block(ws)
        pending = [r for r in rows if ws.cell(r, 6).value is None]
        if not pending:
            continue
        month = ws["A2"].value.replace("PATIENT DATA OF ", "")
        o = wb.create_sheet(ws.title)
        # name line + title bar (same styles as the master)
        copy_cell(ws["A1"], o["A1"])
        copy_cell(ws["A2"], o["A2"], f"RADIOLOGY DATA - OUT PATIENTS - {month}")
        o["A2"].fill = copy(ws.cell(title, 1).fill)
        for col in "BCDEF":
            copy_cell(ws[f"{col}1"], o[f"{col}1"])
            copy_cell(ws[f"{col}2"], o[f"{col}2"])
            o[f"{col}2"].fill = copy(ws.cell(title, 1).fill)
        o.merge_cells("A1:F1")
        o.merge_cells("A2:F2")
        for c in range(1, 7):  # column header row
            copy_cell(ws.cell(title + 1, c), o.cell(3, c))
        for i, r in enumerate(rows):
            t = 4 + i
            for c in range(1, 7):
                copy_cell(ws.cell(r, c), o.cell(t, c))
            if ws.cell(r, 4).comment:
                o.cell(t, 4).comment = copy(ws.cell(r, 4).comment)
            if o.cell(t, 6).value is None:
                o.cell(t, 6).fill = PENDING_FILL
            o.row_dimensions[t].height = ds.ROW_H["data"]
        last = 3 + len(rows)
        lab = ws.cell(rows[-1] + 2, 4)  # 'Total Referrals Count' label style in the master
        num = ws.cell(rows[-1] + 3, 6)
        for k, (label, formula, col) in enumerate([
            ("Total Referrals Count", f"=COUNTA(C4:C{last})", 5),
            ("Total Bill Amount (entered)", f"=SUM(F4:F{last})", 6),
            ("Pending Bill Amounts", f'=COUNTBLANK(F4:F{last})', 5),
        ]):
            r = last + 2 + k
            copy_cell(lab, o.cell(r, 4), label)
            copy_cell(num, o.cell(r, col), formula)
            if col == 5:
                o.cell(r, 5).number_format = "0"
                o.cell(r, 5).alignment = Alignment(horizontal="center", vertical="center")
        note = o.cell(last + 6, 1, "Yellow cells: bill amount pending - fill in and send back to update the main file.")
        note.font = ds.font(10, False, "FF555555")
        for col, w in ds.COL_WIDTHS.items():
            o.column_dimensions[col].width = w
        o.row_dimensions[1].height = ds.ROW_H["name"]
        o.row_dimensions[2].height = ds.ROW_H["title"]
        o.row_dimensions[3].height = ds.ROW_H["header"]
        o.freeze_panes = "A4"
        ds.apply_page_setup(o)
        summary.append((ws.title, len(rows), len(pending)))
    wb.save(out)
    for s in summary:
        print("%s: %d referrals, %d pending" % s)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
