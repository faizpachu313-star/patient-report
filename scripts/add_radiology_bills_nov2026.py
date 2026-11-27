"""Radiology bill amounts from the radiology price lists sent 21 Nov 2026 (Aug 17 cases = 16,799;
Sep 23 cases = 17,117). Rows are matched in the radiology section by date + patient name (the list's
file numbers are new-system TVIP numbers; the main file keeps its own number, the list's number is
noted). Only blank cells are filled.

Run after add_invoice_visits_*:  python scripts/add_radiology_bills_nov2026.py <master.xlsx>
"""
import sys
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

# (sheet, date, name in main file (start), list file no., amount, exam on list)
BILLS = [
    ("Aug-2026", "03.08.2026", "Faisal Khaled", "TVIP01017700", 854, "MRI"),
    ("Aug-2026", "05.08.2026", "Noura Ramadan", "TVIP00383558", 782, "MRI"),
    ("Aug-2026", "06.08.2026", "Aysha Saeed", "TVIP00870626", 854, "MRI"),
    ("Aug-2026", "09.08.2026", "Amal Abdallah", "TVIP00063874", 854, "MRI"),
    ("Aug-2026", "09.08.2026", "Mai Suhail", "TVIP00540138", 795, "MR"),
    ("Aug-2026", "12.08.2026", "Shanda Abdulla", "TVIP00221641", 854, "MRI"),
    ("Aug-2026", "13.08.2026", "Rubayea Salem", "TVIP01018488", 782, "MRI"),
    ("Aug-2026", "16.08.2026", "Tanjina Akther", "TVIP01019155", 260, "US TRANS"),
    ("Aug-2026", "18.08.2026", "Aisha Mohamed", "TVIP01019297", 2484, "3-MRI"),
    ("Aug-2026", "18.08.2026", "Aidah Eisa", "TVIP01019300", 2524, "3-MRI"),
    ("Aug-2026", "19.08.2026", "Abdulla Ahmed", "TVIP01019417", 782, "MRI"),
    ("Aug-2026", "22.08.2026", "Saif Obaid", "TVIP01019899", 854, "MRI"),
    ("Aug-2026", "23.08.2026", "Khuloud Taleb", "TVIP00369954", 835, "MRI"),
    ("Aug-2026", "24.08.2026", "Sabha Ali", "TVIP01020095", 782, "MRI"),
    ("Aug-2026", "27.08.2026", "Sultan Amhi", "TVIP01020546", 795, "MRI"),
    ("Aug-2026", "29.08.2026", "Mariam Hamad", "TVIP00978329", 854, "MRI"),
    ("Aug-2026", "30.08.2026", "Fatima Saeed", "TVIP01020827", 854, "MRI"),
    ("Sep-2026", "01.09.2026", "Shabbir Ahmad", "TVIP01021143", 459, "CT SCAN"),
    ("Sep-2026", "06.09.2026", "Shaikha Ali", "TVIP01021758", 795, "MRI CRV"),
    ("Sep-2026", "09.09.2026", "Wajid Ali", "TVIP01022197", 782, "MRI LUMBR"),
    ("Sep-2026", "12.09.2026", "Musanif Shah", "TVIP01022494", 891, "MRI ARM"),
    ("Sep-2026", "12.09.2026", "Fatima Salama", "TVIP00402597", 795, "MRI CRV"),
    ("Sep-2026", "13.09.2026", "Haider Khan", "TVIP01022522", 782, "MRI LMBR"),
    ("Sep-2026", "14.09.2026", "Khizar Hayat", "TVIP01022788", 474, "CT SCAN"),
    ("Sep-2026", "14.09.2026", "Enayat Ullah", "TVIP00983134", 782, "MRI LUMBR"),
    ("Sep-2026", "14.09.2026", "Shahid Zainudheen", "TVIP00965236", 474, "CT MAX"),
    ("Sep-2026", "15.09.2026", "Muhammad Asif", "TVIP01022844", 474, "CT MAX"),
    ("Sep-2026", "15.09.2026", "Nasser Mohammed", "TVIP01022861", 1670, "MRI B SHOULDER"),
    ("Sep-2026", "17.09.2026", "Anusha Sandamali", "TVIP01023249", 782, "MRI LUMBR"),
    ("Sep-2026", "17.09.2026", "Amal Abdulla", "TVIP01023104", 854, "MRI KNEE"),
    ("Sep-2026", "19.09.2026", "Abhishek Gond", "TVIP01023351", 434, "CT SCAN"),  # user: 434 (price list had 924)
    ("Sep-2026", "22.09.2026", "Osman Rohani", "TVIP01023846", 459, "CT SCAN"),
    ("Sep-2026", "22.09.2026", "Safeer Kunnath", "TVIP01023860", 782, "MRI LUMBR"),
    ("Sep-2026", "22.09.2026", "Muhammad Waseem", "TVIP00851673", 782, "MRI LUMBR"),
    ("Sep-2026", "26.09.2026", "Kala Miah", "TVIP01024211", 459, "CT SCAN"),
    ("Sep-2026", "26.09.2026", "Alia Huwaiden", "TVIP00970990", 854, "MRI KNEE"),
    ("Sep-2026", "28.09.2026", "Abhishek Gond", "TVIP01023351", 924, "MRI HIP WC"),  # date per Spectrum list
    ("Sep-2026", "29.09.2026", "Emaan Blaisha", "TVIP00579183", 795, "MRI"),
    ("Sep-2026", "30.09.2026", "Tilak Bahadur", "TVIP01024887", 854, "MRI"),
    ("Sep-2026", "30.09.2026", "", "TVIP01024882", 270, "US RETRO"),  # name blank in main file
]
FILL_NAME = {("Sep-2026", "TVIP01024882"): "Jeshitha Mujeeb Kunjumon"}


def main(path):
    wb = openpyxl.load_workbook(path)
    for sheet, d, name, fno, amt, exam in BILLS:
        ws = wb[sheet]
        dd = datetime.strptime(d, "%d.%m.%Y")
        start = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS")
        hits = [r for r in range(start, ws.max_row + 1) if ws.cell(r, 2).value == dd
                and (str(ws.cell(r, 4).value or "").startswith(name) if name else str(ws.cell(r, 3).value) == fno)
                and (exam.split()[0][:2] in str(ws.cell(r, 5).value).upper())]
        assert len(hits) == 1, (sheet, d, name, exam, hits)
        r = hits[0]
        c = ws.cell(r, 6)
        assert c.value is None, (sheet, d, name, c.value)
        c.value = amt
        c.number_format = "#,##0.00"
        note = f"Radiology price list (21 Nov 2026): {exam} {amt}"
        if str(ws.cell(r, 3).value) != fno:
            note += f"; list file no. {fno}"
        c.comment = Comment(note, "Claude")
        key = (sheet, str(ws.cell(r, 3).value))
        if key in FILL_NAME and not ws.cell(r, 4).value:
            ws.cell(r, 4).value = FILL_NAME[key]
            ws.cell(r, 4).comment = Comment("Name from radiology price list (was blank).", "Claude")
    wb.save(path)
    print("radiology bills filled:", len(BILLS))


if __name__ == "__main__":
    main(sys.argv[1])
