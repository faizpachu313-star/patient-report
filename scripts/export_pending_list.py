"""TEMPORARY share file for the billing team: only the rows whose bill amount is still pending.

One sheet per month (Part B: Jul 2026 onwards) listing pending patient visits and pending radiology
referrals, with a yellow Bill Amount column to fill and a Remarks column for the team. An overview
sheet counts what is pending. The main workbook is only read.

Usage: python scripts/export_pending_list.py <master.xlsx> <out.xlsx> [first_month_sheet]
"""
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402

YELLOW = PatternFill("solid", fgColor=ds.YELLOW)
HEAD = ["Sl No", "Section", "Date", "File Number", "Patient Name", "Service / Department",
        "Bill Amount (AED)", "Remarks (billing team)"]
WIDTHS = [7, 12, 14, 20, 40, 40, 16, 34]


def pending_rows(ws):
    out, sec = [], "Patient Visit"
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS":
            sec = "Radiology"
        if isinstance(ws.cell(r, 1).value, int) and isinstance(ws.cell(r, 2).value, datetime) \
                and ws.cell(r, 6).value is None:
            out.append((sec, ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value))
    return out


def header(o, title, ncol):
    o["A1"] = "NAME : MUHAMMED FAIZ"
    ds.style_name_line(o["A1"])
    o["A2"] = title
    ds.style_title(o["A2"])
    for c in range(2, ncol + 1):
        ds.style_name_line(o.cell(1, c))
        ds.style_title(o.cell(2, c))
    o.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    o.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    o.row_dimensions[1].height = ds.ROW_H["name"]
    o.row_dimensions[2].height = ds.ROW_H["title"]


def main(master, out, first="Jul-2026"):
    src = openpyxl.load_workbook(master)
    months = src.sheetnames[src.sheetnames.index(first):]
    wb = openpyxl.Workbook()
    ov = wb.active
    ov.title = "Overview"
    summary = []
    for name in months:
        rows = pending_rows(src[name])
        if not rows:
            continue
        o = wb.create_sheet(name)
        month = src[name]["A2"].value.replace("PATIENT DATA OF ", "")
        header(o, f"PENDING BILL AMOUNTS - {month}", len(HEAD))
        for i, h in enumerate(HEAD, 1):
            o.cell(3, i, h)
            ds.style_header(o.cell(3, i))
        for i, (sec, d, fno, nm, svc) in enumerate(rows, 1):
            r = 3 + i
            for c, v in enumerate((i, sec, d, fno, nm, svc, None, None), 1):
                cell = o.cell(r, c, v)
                ds.style_data(cell, "D" if c in (5, 8) else "A")
            o.cell(r, 3).number_format = ds.DATE_FMT
            o.cell(r, 7).fill = YELLOW
            o.cell(r, 7).number_format = ds.AMOUNT_FMT
            o.row_dimensions[r].height = ds.ROW_H["data"]
        for c, w in enumerate(WIDTHS, 1):
            o.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
        o.row_dimensions[3].height = ds.ROW_H["header"]
        o.freeze_panes = "A4"
        ds.apply_page_setup(o)
        nv = sum(1 for x in rows if x[0] == "Patient Visit")
        summary.append((name, nv, len(rows) - nv))
    header(ov, "TEMPORARY - PENDING BILL AMOUNTS FOR BILLING TEAM", 5)
    for i, h in enumerate(["Sl No", "Month", "Patient Visits Pending", "Radiology Pending", "Total Pending"], 1):
        ov.cell(3, i, h)
        ds.style_header(ov.cell(3, i))
    for i, (name, nv, nr) in enumerate(summary, 1):
        r = 3 + i
        for c, v in enumerate((i, name, nv, nr, f"=C{r}+D{r}"), 1):
            ds.style_data(ov.cell(r, c, v), "A")
    t = 4 + len(summary) + 1
    ov.cell(t, 1, "TOTAL")
    for c in range(1, 6):
        ds.style_total(ov.cell(t, c), h="right" if c == 1 else "center")
    ov.merge_cells(start_row=t, start_column=1, end_row=t, end_column=2)
    for col in "CDE":
        ov[f"{col}{t}"] = f"=SUM({col}4:{col}{3 + len(summary)})"
    note = ov.cell(t + 2, 1, "Fill the yellow Bill Amount cells from the system invoices (sum all invoices of the "
                             "same day) and add a remark if needed. Temporary file - returned amounts are checked "
                             "row by row before they go into the main file.")
    note.font = ds.font(10, False, "FF555555")
    note.alignment = ds.align("left")
    ov.merge_cells(start_row=t + 2, start_column=1, end_row=t + 2, end_column=5)
    ov.row_dimensions[t + 2].height = 45
    for col, w in zip("ABCDE", (8, 16, 22, 20, 16)):
        ov.column_dimensions[col].width = w
    ov.row_dimensions[3].height = 36
    ds.apply_page_setup(ov, print_titles=False)
    wb.save(out)
    for s in summary:
        print("%s: %d visits, %d radiology pending" % s)


if __name__ == "__main__":
    main(*sys.argv[1:])
