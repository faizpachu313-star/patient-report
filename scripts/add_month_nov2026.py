"""Create the Nov-2026 month sheet (in-progress month) from the clinic system's invoice screens.

Same base layout as the other months (name line, title, header, data rows, spacer, total), so the
later steps (radiology section, Bill Amount column, Summary columns) apply to it unchanged.
Also adds the November 2026 row to the Summary sheet.

Run after scripts/merge_uploads_nov2026.py and before scripts/add_radiology_sections.py:
    python scripts/add_month_nov2026.py <master.xlsx>
"""
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402

SHEET, TITLE, LABEL = "Nov-2026", "PATIENT DATA OF NOVEMBER 2026", "November 2026"

# (date, file, name, [(department/doctor, amount), ...]) - every Nov 2026 invoice in the screenshots
VISITS = [
    ("01.11.2026", "TVIP01020920", "Zaid Mustafa Refaieh", [("Dental (Dr Ahmed Yousif Ali)", 158.00)]),
    ("02.11.2026", "TVIP01001897", "Mohammed Ameer Puzhangaraillath", [("Family Medicine (Dr Nighet)", 202.60)]),
    ("02.11.2026", "TVIP00830448", "Kavitha Satheesh",
     [("General (Dr Roaa)", 95.00), ("Internal Medicine (Dr Ali)", 95.00)]),
    ("04.11.2026", "TVIP00946055", "Hessa Musallem Hamad", [("Dermatology (Dr Shiraz)", 251.50)]),
    ("06.11.2026", "TVIP00830448", "Kavitha Satheesh", [("Physiotherapy (Hadir Nour)", 169.00)]),
    ("14.11.2026", "TVIP00901649", "Muhammed Ameen", [("Internal Medicine (Dr Ali)", 95.00)]),
    ("18.11.2026", "TVIP00979556", "Zayan Hamad Maheen",
     [("General (Dr Roaa)", 95.00), ("Pediatrics (Dr Enas)", 299.69)]),
    ("19.11.2026", "TVIP00979556", "Zayan Hamad Maheen", [("Pediatrics (Basel Abdulhamid)", 95.19)]),
    ("20.11.2026", "TVIP00979556", "Zayan Hamad Maheen", [("Pediatrics (Dr Ahmad Abdul Baset)", 163.11)]),
    ("20.11.2026", "TVIP00830448", "Kavitha Satheesh", [("Physiotherapy (Hadir Nour)", 169.00)]),
]


def build_sheet(wb):
    ws = wb.create_sheet(SHEET)
    ws["A1"] = "NAME : MUHAMMED FAIZ"
    ds.style_name_line(ws["A1"])
    ws["A2"] = TITLE
    ds.style_title(ws["A2"])
    ws["A2"].comment = Comment("Month in progress: built from the clinic system's invoice screens "
                               "(patients checked so far, up to 21 Nov 2026).", "Claude")
    for i, h in enumerate(["Sl No", "Date", "File Number", "Patient Name",
                           "Consultation / Service / Department"], 1):
        ws.cell(3, i, h)
        ds.style_header(ws.cell(3, i))
    for col in "BCDE":
        ds.style_name_line(ws[f"{col}1"])
        ds.style_title(ws[f"{col}2"])
    for i, (d, fno, name, parts) in enumerate(VISITS, 1):
        r = 3 + i
        svc = ", ".join(p for p, _ in parts)
        for col, v in zip("ABCDE", (i, datetime.strptime(d, "%d.%m.%Y"), fno, name, svc)):
            ws[f"{col}{r}"] = v
            ds.style_data(ws[f"{col}{r}"], col)
        ws.row_dimensions[r].height = 30 if len(svc) > 45 else ds.ROW_H["data"]
    last = 3 + len(VISITS)
    sp, tot = last + 1, last + 2
    for col in "ABCDE":
        ds.style_spacer(ws[f"{col}{sp}"])
        ds.style_total(ws[f"{col}{tot}"], h="right" if col == "A" else "center")
    ws.row_dimensions[sp].height = ds.ROW_H["spacer"]
    ws.row_dimensions[tot].height = ds.ROW_H["total"]
    ws[f"A{tot}"] = "TOTAL PATIENTS (VISITS) - NOVEMBER 2026"
    ws[f"E{tot}"] = f"=COUNTA(B4:B{last})"
    for ref in ("A1:E1", "A2:E2", f"A{sp}:E{sp}", f"A{tot}:D{tot}"):
        ws.merge_cells(ref)
    for col in "ABCDE":
        ws.column_dimensions[col].width = ds.COL_WIDTHS[col]
    ws.row_dimensions[1].height = ds.ROW_H["name"]
    ws.row_dimensions[2].height = ds.ROW_H["title"]
    ws.row_dimensions[3].height = ds.ROW_H["header"]
    ws.freeze_panes = "A4"
    ds.apply_page_setup(ws)
    return tot


def add_summary_row(wb, tot):
    s = wb["Summary"]
    last = 3 + len(wb.sheetnames[1:]) - 1        # current last month row (before November)
    merges = [str(m) for m in s.merged_cells.ranges]
    for m in merges:
        s.unmerge_cells(m)
    s.insert_rows(last + 1)
    new = last + 1
    for c in range(1, s.max_column + 1):
        src, dst = s.cell(last, c), s.cell(new, c)
        dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
        dst.alignment, dst.number_format = copy(src.alignment), src.number_format
    # Excel/openpyxl do not move row heights on insert: shift them by hand
    heights = {k: d.height for k, d in s.row_dimensions.items() if d.height is not None}
    for k in list(s.row_dimensions.keys()):
        s.row_dimensions[k].height = None
    for k, h in heights.items():
        s.row_dimensions[k + (k >= new)].height = h
    s.row_dimensions[new].height = heights.get(last)
    for m in merges:  # ranges below the new row move down one
        a, b = m.split(":")
        ra, rb = int("".join(ch for ch in a if ch.isdigit())), int("".join(ch for ch in b if ch.isdigit()))
        ca, cb = a.rstrip("0123456789"), b.rstrip("0123456789")
        s.merge_cells(f"{ca}{ra + (ra >= new)}:{cb}{rb + (rb >= new)}")
    s.cell(new, 1, s.cell(last, 1).value + 1)
    s.cell(new, 2, LABEL)
    if s.cell(last, 2).hyperlink:  # month names link to their sheets
        s.cell(new, 2).hyperlink = f"#'{SHEET}'!A1"
        s.cell(new, 2).hyperlink.location = f"'{SHEET}'!A1"
    s.cell(new, 3, f"='{SHEET}'!E{tot}")
    grand = new + 2
    s.cell(grand, 3, f"=SUM(C4:C{new})")
    s["A2"] = s["A2"].value.replace("OCT 2026", "NOV 2026")
    note = grand + 2
    s.cell(note, 1).value = (s.cell(note, 1).value.rstrip(".") +
                             "; Nov 2026 (in progress) from the clinic system's invoice screens.")


def main(path):
    wb = openpyxl.load_workbook(path)
    assert SHEET not in wb.sheetnames
    tot = build_sheet(wb)
    add_summary_row(wb, tot)
    # bill amounts (sum of each day's invoices) with breakdown note
    ws = wb[SHEET]
    for i, (_, _, _, parts) in enumerate(VISITS):
        c = ws.cell(4 + i, 6, round(sum(a for _, a in parts), 2))
        c.comment = Comment("Invoices: " + " + ".join(f"{p} {a:.2f}" for p, a in parts), "Claude")
    wb.save(path)
    print(f"{SHEET}: {len(VISITS)} visits; Summary row added")


if __name__ == "__main__":
    main(sys.argv[1])
