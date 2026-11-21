"""Add visits that exist in the clinic system's invoice lists but were missing from the staff sheets.

The invoice screens are the source of truth (user instruction, 21 Nov 2026). Each patient's full
invoice list (Jul-Oct 2026) was reconciled against the main file; these visit days had invoices but
no row. One row per patient per day, amount = sum of that day's invoices, breakdown in a cell note.

Run last, after scripts/add_bill_amounts_invoices_nov2026.py:
    python scripts/add_invoice_visits_nov2026.py <master.xlsx>
"""
import sys
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402
from merge_uploads_nov2026 import insert_row  # noqa: E402

# (sheet, date, file, name, [(department/doctor, amount), ...])
NEW_VISITS = [
    ("Jul-2026", "02.07.2026", "TVIP00946055", "Hessa Musallem Hamad", [("Dermatology (Dr Shiraz)", 291.60)]),
    ("Jul-2026", "07.07.2026", "TVIP00946055", "Hessa Musallem Hamad",
     [("Allergy & Immunology (Dr Sally)", 1371.50), ("Dermatology (Dr Shiraz)", 302.00)]),
    ("Jul-2026", "08.07.2026", "TVIP00946055", "Hessa Musallem Hamad", [("Dermatology (Dr Shiraz)", 184.60)]),
    ("Jul-2026", "14.07.2026", "TVIP01015196", "Adel Fares Afifi Ibrahim Afifi", [("Radiology (Dr Allaeldein)", 835.00)]),
    ("Aug-2026", "03.08.2026", "TVIP00946055", "Hessa Musallem Hamad", [("Dermatology (Dr Shiraz)", 128.00)]),
    ("Aug-2026", "10.08.2026", "TVIP01017206", "Rana Muhammad Zubair", [("Physiotherapy (Jimshad)", 169.00)]),
    ("Aug-2026", "13.08.2026", "TVIP01015196", "Adel Fares Afifi Ibrahim Afifi",
     [("Pulmonology (Dr Issa)", 638.17), ("General (Afsheen)", 160.00)]),
    ("Aug-2026", "13.08.2026", "TVIP00830448", "Kavitha Satheesh",
     [("Gynecology (Dr Walaa)", 1535.00), ("Internal Medicine (Dr Ali)", 95.00),
      ("Allergy & Immunology (Dr Sally)", 95.00)]),
    ("Aug-2026", "13.08.2026", "TVIP01017206", "Rana Muhammad Zubair", [("Physiotherapy (Jimshad)", 169.00)]),
    ("Aug-2026", "22.08.2026", "TVIP00830448", "Kavitha Satheesh", [("Gynecology (Dr Walaa)", 1194.00)]),
    ("Oct-2026", "02.10.2026", "TVIP00849060", "Yasmine Mahmod Sayyo", [("Physiotherapy (Hadir Nour)", 135.00)]),
    ("Oct-2026", "03.10.2026", "TVIP00946055", "Hessa Musallem Hamad", [("General (Dr Mozammil)", 560.14)]),
    ("Oct-2026", "09.10.2026", "TVIP00849060", "Yasmine Mahmod Sayyo", [("Physiotherapy (Hadir Nour)", 135.00)]),
    ("Oct-2026", "09.10.2026", "TVIP01020917", "Shahid Mustafa Refaieh", [("Pediatrics (Dr Ghaith)", 95.00)]),
    ("Oct-2026", "11.10.2026", "TVIP01015196", "Adel Fares Afifi Ibrahim Afifi",
     [("Internal Medicine (Dr Ali)", 95.00), ("Dermatology (Dr Shiraz)", 142.00), ("General (Dr Mozammil)", 95.00)]),
    ("Oct-2026", "25.10.2026", "TVIP00849060", "Yasmine Mahmod Sayyo", [("Physiotherapy (Hadir Nour)", 135.00)]),
    ("Oct-2026", "25.10.2026", "TVIP01020920", "Zaid Mustafa Refaieh", [("Dental (Dr Ahmed Yousif Ali)", 675.00)]),
]

# Rows in the staff sheets that have NO invoice on that date in the patient's full invoice list.
# Kept (not deleted) and flagged with a note until the user decides.
NO_INVOICE = [
    ("Aug-2026", "01.08.2026", "TVIP01017202"), ("Aug-2026", "04.08.2026", "TVIP01015196"),
    ("Aug-2026", "04.08.2026", "TVIP01001897"), ("Aug-2026", "30.08.2026", "TVIP01020920"),
    ("Sep-2026", "02.09.2026", "TVIP00950460"), ("Sep-2026", "05.09.2026", "TVIP00604064"),
    ("Sep-2026", "14.09.2026", "TVIP01017200"), ("Sep-2026", "17.09.2026", "TVIP01017202"),
]


def visit_rows(ws):
    rows = []
    for r in range(4, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").startswith("TOTAL PATIENTS"):
            return rows, r
        if isinstance(ws.cell(r, 2).value, datetime):
            rows.append(r)
    raise ValueError(ws.title)


def main(path):
    wb = openpyxl.load_workbook(path)
    for sheet, d, fno, name, parts in NEW_VISITS:
        ws = wb[sheet]
        dd = datetime.strptime(d, "%d.%m.%Y")
        rows, total = visit_rows(ws)
        assert not any(ws.cell(r, 2).value == dd and str(ws.cell(r, 3).value) == fno for r in rows), (sheet, d, fno)
        later = [r for r in rows if ws.cell(r, 2).value > dd]
        idx = later[0] if later else rows[-1] + 1
        insert_row(wb, ws, idx, rows[0])
        amount = round(sum(a for _, a in parts), 2)
        svc = ", ".join(p for p, _ in parts)
        for col, v in zip("BCDEF", (dd, fno, name, svc, amount)):
            ws[f"{col}{idx}"] = v
            ds.style_data(ws[f"{col}{idx}"], col)
        ws.row_dimensions[idx].height = 30 if len(svc) > 45 else ds.ROW_H["data"]
        ws[f"D{idx}"].comment = Comment("Added from the clinic system's invoice list (missing in staff sheet).", "Claude")
        ws[f"F{idx}"].comment = Comment("Invoices: " + " + ".join(f"{p} {a:.2f}" for p, a in parts), "Claude")
        # renumber and widen the month's visit totals to the new last row
        rows, total = visit_rows(ws)
        for i, r in enumerate(rows, 1):
            ws.cell(r, 1).value = i
        last = rows[-1]
        ws.cell(total, 5).value = f"=COUNTA(B4:B{last})"
        bill = total + 1
        assert str(ws.cell(bill, 1).value).startswith("TOTAL BILL AMOUNT"), (sheet, bill)
        ws.cell(bill, 6).value = f"=SUM(F4:F{last})"
    for sheet, d, fno in NO_INVOICE:
        ws = wb[sheet]
        dd = datetime.strptime(d, "%d.%m.%Y")
        rows, _ = visit_rows(ws)
        hit = [r for r in rows if ws.cell(r, 2).value == dd and str(ws.cell(r, 3).value) == fno]
        assert len(hit) == 1, (sheet, d, fno, hit)
        ws.cell(hit[0], 6).comment = Comment(
            "No invoice on this date in the patient's full invoice list (clinic system). "
            "Possibly a staff-sheet error - awaiting confirmation.", "Claude")
    wb.save(path)
    print("visits added from invoices:", len(NEW_VISITS), "| no-invoice rows flagged:", len(NO_INVOICE))


if __name__ == "__main__":
    main(sys.argv[1])
