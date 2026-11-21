"""Bill amounts from the clinic system's invoice screens (photos sent 21 Nov 2026).

Each visit row gets the SUM of that patient's invoices on that date (one row per patient per day).
Rows are found by sheet + date + file number. Every change asserts the current value first, so a
row is only filled if it is blank, and only corrected if it still holds the expected old value.

Run after scripts/add_bill_amounts_nov2026.py:
    python scripts/add_bill_amounts_invoices_nov2026.py <master.xlsx>
"""
import sys
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

# (sheet, date, file no, expected current value (None = blank), new amount, invoice breakdown)
CHANGES = [
    # --- September 2026: pending amounts filled ---
    ("Sep-2026", "01.09.2026", "TVIP00946055", None, 128.00, "Dermatology (Dr Shiraz) 128.00"),
    ("Sep-2026", "02.09.2026", "TVIP01020920", None, 517.50, "Dental (Dr Ahmed Yousif Ali) 517.50"),
    ("Sep-2026", "02.09.2026", "TVIP01021338", None, 104.50, "Pediatrics (Basel Abdulhamid) 104.50"),
    ("Sep-2026", "03.09.2026", "TVIP01015196", None, 489.00, "Internal Medicine (Dr Ali) 489.00"),
    ("Sep-2026", "05.09.2026", "TVIP00979556", None, 138.00, "Pediatrics (Dr Zafer) 138.00"),
    ("Sep-2026", "05.09.2026", "TVIP00950460", None, 113.00, "Dental (Dr Rafeh) 113.00"),
    ("Sep-2026", "05.09.2026", "TVIP01021611", None, 733.00, "General (Dr Asmaa) 95.00 + Ophthalmology (Al Droos) 638.00"),
    ("Sep-2026", "09.09.2026", "TVIP01015196", None, 95.00, "General (Dr Roaa) 95.00"),
    ("Sep-2026", "11.09.2026", "TVIP00830448", None, 169.00, "Physiotherapy (Fensi) 169.00"),
    ("Sep-2026", "11.09.2026", "784-1991-1902026-3", None, 142.00,
     "Family Medicine (Dr Nighet) 142.00 - invoice file no. TVIP01022335"),
    ("Sep-2026", "12.09.2026", "TVIP00973129", None, 485.00, "Gynecology (Dr Walaa) 343.00 + ENT (Dr Nawras) 142.00"),
    ("Sep-2026", "13.09.2026", "TVIP00950460", None, 563.00, "Dental (Dr Rafeh) 563.00"),
    ("Sep-2026", "18.09.2026", "TVIP00830448", None, 169.00, "Physiotherapy (Fensi) 169.00"),
    ("Sep-2026", "20.09.2026", "TVIP01020920", None, 142.00, "Pediatrics (Basel Abdulhamid) 142.00"),
    ("Sep-2026", "26.09.2026", "TVIP00973129", None, 181.00, "Gynecology (Dr Walaa) 181.00"),
    ("Sep-2026", "30.09.2026", "TVIP00950460", None, 443.81, "ENT (Dr Nandini) 443.81 (partially paid)"),
    # --- August 2026: pending amounts filled ---
    ("Aug-2026", "16.08.2026", "TVIP00979556", None, 645.11, "General (Dr Mustafa) 95.00 + Pediatrics (Dr Enas) 550.11"),
    ("Aug-2026", "18.08.2026", "TVIP00830448", None, 95.00, "Orthopedic (Dr Ihab) 95.00"),
    ("Aug-2026", "20.08.2026", "TVIP00973129", None, 342.00, "Gynecology (Dr Walaa) 342.00"),
    ("Aug-2026", "25.08.2026", "TVIP01015196", None, 285.00,
     "Pulmonology (Dr Issa) 95.00 + Internal Medicine (Dr Ali) 95.00 + General (Dr Dimeh) 95.00"),
    ("Aug-2026", "29.08.2026", "TVIP00604064", None, 232.00, "Internal Medicine (Dr Ali) 232.00"),
    # --- corrections: existing amount differs from the invoices ---
    ("Jul-2026", "20.07.2026", "TVIP01015196", 430, 460.00,
     "Physiotherapy (Alamgir) 270.00 + General (Dr Mofeed) 95.00 + Orthopedic (Dr Ihab) 95.00. Was 430 (75/85/270)"),
    ("Jul-2026", "21.07.2026", "TVIP00979556", 289.2, 339.00,
     "General (Dr Roaa) 95.00 (not paid) + ENT (Dr Nandini) 244.00. Was 289.20 (85/204.2)"),
    ("Jul-2026", "25.07.2026", "TVIP00973129", 136.4, 141.00, "Gynecology (Dr Walaa) 141.00 (not paid). Was 136.40"),
    ("Jul-2026", "31.07.2026", "TVIP00604064", 849, 929.00,
     "General (Dr Roaa) 95.00 + Internal Medicine (Dr Ali) 834.00. Was 849"),
    ("Aug-2026", "03.08.2026", "TVIP00830448", 2889.2, 3031.00,
     "Orthopedic (Dr Ihab) 1272.00 + General (Afsheen) 95.00 + Endocrinology (Dr Issam) 1509.00 (partially paid) "
     "+ Dermatology (Dr Shiraz) 155.00 (not paid). Was 2889.20 (photo list)"),
    ("Aug-2026", "05.08.2026", "TVIP01015196", 85, 95.00, "Orthopedic (Dr Higazy) 95.00. Was 85 (photo list)"),
]


def main(path):
    wb = openpyxl.load_workbook(path)
    for sheet, d, fno, old, new, note in CHANGES:
        ws = wb[sheet]
        dd = datetime.strptime(d, "%d.%m.%Y")
        sec, hits = "visit", []
        for r in range(4, ws.max_row + 1):
            if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS":
                sec = "rad"
            if sec == "visit" and ws.cell(r, 2).value == dd and str(ws.cell(r, 3).value) == fno:
                hits.append(r)
        assert len(hits) == 1, (sheet, d, fno, hits)
        c = ws.cell(hits[0], 6)
        assert c.value == old, (sheet, d, fno, "expected", old, "found", c.value)
        c.value = new
        c.number_format = "#,##0.00"
        c.comment = Comment("Invoices: " + note, "Claude")
    wb.save(path)
    print("invoice amounts applied:", len(CHANGES))


if __name__ == "__main__":
    main(sys.argv[1])
