"""Fill BLANK radiology bill amounts that have a standard price confirmed by the user.

  US PREGNANT (incl. spellings US PREGNANCY / US PREAGNANT) = AED 390   (user, 21 Nov 2026)

Only blank cells are filled; existing amounts are never changed. Run after the add_bill_amounts_*
steps:  python scripts/add_radiology_standard_prices.py <master.xlsx>
"""
import re
import sys
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

PRICES = [(re.compile(r"^\s*US\s+PRE[A-Z]*GNAN", re.I), 390.00, "US pregnant standard price AED 390 (user, 21 Nov 2026)")]


def main(path):
    wb = openpyxl.load_workbook(path)
    filled = {}
    for ws in wb.worksheets[1:]:
        in_rad = False
        for r in range(4, ws.max_row + 1):
            if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS":
                in_rad = True
            if not (in_rad and isinstance(ws.cell(r, 2).value, datetime)) or ws.cell(r, 6).value is not None:
                continue
            exam = str(ws.cell(r, 5).value or "")
            for pat, price, note in PRICES:
                if pat.search(exam):
                    c = ws.cell(r, 6, price)
                    c.number_format = "#,##0.00"
                    c.comment = Comment(note, "Claude")
                    filled[ws.title] = filled.get(ws.title, 0) + 1
                    break
    wb.save(path)
    print("standard radiology prices filled:", filled)


if __name__ == "__main__":
    main(sys.argv[1])
