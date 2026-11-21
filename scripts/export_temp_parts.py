"""Export the TEMPORARY working copies of the main data, used only to fill in pending bill amounts.

  Part A = Nov 2025 - Jun 2026      Part B = Jul 2026 - last month in the master

Each month sheet is an exact copy of the master month sheet (visits + radiology, same design and
formulas); every blank Bill Amount cell is highlighted. A month whose amounts are all filled is
left out (finished months are removed from the temporary files). The first sheet, 'Pending',
counts the blank amounts per month with live formulas.

The master workbook is only read, never written.

Usage: python scripts/export_temp_parts.py <master.xlsx> <part_a.xlsx> <part_b.xlsx>
"""
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402

PART_A = ["Nov-2025", "Dec-2025", "Jan-2026", "Feb-2026", "Mar-2026", "Apr-2026", "May-2026", "Jun-2026"]
PENDING_FILL = PatternFill("solid", fgColor=ds.YELLOW)


def sections(ws):
    """Return (visit_rows, radiology_rows) of data rows in a month sheet."""
    visit, rad, in_rad = [], [], False
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS":
            in_rad = True
        if isinstance(ws.cell(r, 1).value, int) and isinstance(ws.cell(r, 2).value, datetime):
            (rad if in_rad else visit).append(r)
    return visit, rad


def copy_sheet(src, dst):
    for row in src.iter_rows():
        for c in row:
            t = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                t.font, t.fill, t.border = copy(c.font), copy(c.fill), copy(c.border)
                t.alignment, t.number_format = copy(c.alignment), c.number_format
            if c.comment:
                t.comment = copy(c.comment)
    for rng in src.merged_cells.ranges:
        dst.merge_cells(str(rng))
    for k, d in src.column_dimensions.items():
        dst.column_dimensions[k].width = d.width
    for k, d in src.row_dimensions.items():
        if d.height is not None:
            dst.row_dimensions[k].height = d.height
    dst.freeze_panes = src.freeze_panes
    ds.apply_page_setup(dst)


def build(master, months, out, label):
    wb = openpyxl.Workbook()
    idx = wb.active
    idx.title = "Pending"
    kept = []
    for name in months:
        src = master[name]
        visit, rad = sections(src)
        if not any(src.cell(r, 6).value is None for r in visit + rad) and name != master.sheetnames[-1]:
            continue  # month finished -> not in the temporary file (the current month always stays)
        dst = wb.create_sheet(name)
        copy_sheet(src, dst)
        for r in visit + rad:
            if dst.cell(r, 6).value is None:
                dst.cell(r, 6).fill = PENDING_FILL
        kept.append((name, visit, rad))

    # --- Pending overview (live formulas) ---
    idx["A1"] = "NAME : MUHAMMED FAIZ"
    ds.style_name_line(idx["A1"])
    idx["A2"] = f"TEMPORARY - PENDING BILL AMOUNTS - {label}"
    ds.style_title(idx["A2"])
    heads = ["Sl No", "Month", "Patient Visits", "Visit Bills Pending",
             "Radiology Referrals", "Radiology Bills Pending", "Status"]
    for i, h in enumerate(heads, 1):
        idx.cell(3, i, h)
        ds.style_header(idx.cell(3, i))
    for i, (name, visit, rad) in enumerate(kept, 1):
        r = 3 + i
        vrng = f"'{name}'!F{visit[0]}:F{visit[-1]}" if visit else None
        rrng = f"'{name}'!F{rad[0]}:F{rad[-1]}" if rad else None
        vals = [i, name, len(visit), f"=COUNTBLANK({vrng})" if vrng else 0,
                len(rad), f"=COUNTBLANK({rrng})" if rrng else 0,
                f'=IF(D{r}+F{r}=0,"Complete - ready to move to main file","Pending")']
        for c, v in enumerate(vals, 1):
            idx.cell(r, c, v)
            ds.style_data(idx.cell(r, c), "D" if c == 2 else "A")
    end = 3 + len(kept)
    tot = end + 2
    idx.cell(tot, 1, "TOTAL")
    for c in range(1, 8):
        ds.style_total(idx.cell(tot, c), h="right" if c == 1 else "center")
    idx.merge_cells(f"A{tot}:B{tot}")
    for col in "CDEF":
        idx[f"{col}{tot}"] = f"=SUM({col}4:{col}{end})"
    note = idx.cell(tot + 2, 1, "Temporary file: yellow cells = bill amount pending. Fill them and send the "
                                "file back; once a month has no pending amounts it is moved to the main file "
                                "and its sheet is removed from here.")
    note.font = ds.font(10, False, "FF555555")
    note.alignment = ds.align("left")
    idx.row_dimensions[tot + 2].height = 30
    idx.merge_cells("A1:G1")
    idx.merge_cells("A2:G2")
    idx.merge_cells(f"A{tot + 2}:G{tot + 2}")
    for col, w in zip("ABCDEFG", (8, 18, 16, 18, 18, 20, 34)):
        idx.column_dimensions[col].width = w
    idx.row_dimensions[1].height = ds.ROW_H["name"]
    idx.row_dimensions[2].height = ds.ROW_H["title"]
    idx.row_dimensions[3].height = 36
    ds.apply_page_setup(idx, print_titles=False)
    wb.save(out)
    return kept


def main(master_path, out_a, out_b):
    master = openpyxl.load_workbook(master_path)
    months = master.sheetnames[1:]
    part_b = [m for m in months if m not in PART_A]
    for out, ms, label in ((out_a, PART_A, "PART A (NOV 2025 - JUN 2026)"),
                           (out_b, part_b, f"PART B ({part_b[0].upper()} - {part_b[-1].upper()})")):
        kept = build(master, ms, out, label)
        print(out)
        for name, visit, rad in kept:
            vp = sum(master[name].cell(r, 6).value is None for r in visit)
            rp = sum(master[name].cell(r, 6).value is None for r in rad)
            print(f"  {name}: visits {len(visit)} ({vp} pending), radiology {len(rad)} ({rp} pending)")


if __name__ == "__main__":
    main(*sys.argv[1:4])
