"""ONE-TIME fill (user, 21 Nov 2026): the US PREGNANT radiology cells that were blank on that day = AED 390.

NOT a standing rule: US pregnant prices vary (discounts for cash patients: 260, 290, 388, ...). Any new
or future US pregnant row stays blank until its real amount is given. Only the 47 rows listed below
are filled, and only if still blank. Run after the add_bill_amounts_* steps:
    python scripts/add_radiology_standard_prices.py <master.xlsx>
"""
import sys
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

PRICE = 390.00
NOTE = "US pregnant AED 390: one-time fill of blank cells (user, 21 Nov 2026)"
# (sheet, date, file no.) of the radiology rows filled
ROWS = [
    ('Aug-2026', '05.08.2026', 'TVIP01012803'),
    ('Aug-2026', '05.08.2026', 'TVIP00899382'),
    ('Aug-2026', '05.08.2026', 'TVIP01017948'),
    ('Aug-2026', '05.08.2026', 'TVIP01017956'),
    ('Aug-2026', '05.08.2026', 'TVIP01012463'),
    ('Aug-2026', '06.08.2026', 'TVIP01018050'),
    ('Aug-2026', '06.08.2026', 'TVIP00922585'),
    ('Aug-2026', '08.08.2026', 'TVIP01018290'),
    ('Aug-2026', '08.08.2026', 'TVIP01018291'),
    ('Aug-2026', '12.08.2026', 'TVIP01018828'),
    ('Aug-2026', '12.08.2026', 'TVIP01015684'),
    ('Aug-2026', '13.08.2026', 'TVIP01018885'),
    ('Aug-2026', '13.08.2026', 'TVIP01018920'),
    ('Aug-2026', '16.08.2026', 'TVIP01019170'),
    ('Aug-2026', '16.08.2026', 'TVIP01019178'),
    ('Aug-2026', '16.08.2026', 'TVIP01012813'),
    ('Aug-2026', '17.08.2026', 'TVIP01019287'),
    ('Aug-2026', '19.08.2026', 'TVIP01019555'),
    ('Aug-2026', '19.08.2026', 'TVIP01013011'),
    ('Aug-2026', '23.08.2026', 'TVIP01019927'),
    ('Aug-2026', '26.08.2026', 'TVIP01020441'),
    ('Aug-2026', '26.08.2026', 'TVIP01020489'),
    ('Aug-2026', '30.08.2026', 'TVIP01020965'),
    ('Aug-2026', '30.08.2026', 'TVIP01020834'),
    ('Sep-2026', '02.09.2026', 'TVIP01014023'),
    ('Sep-2026', '02.09.2026', 'TVIP01021249'),
    ('Sep-2026', '02.09.2026', 'TVIP01015398'),
    ('Sep-2026', '03.09.2026', 'TVIP01021417'),
    ('Sep-2026', '03.09.2026', 'TVIP01013116'),
    ('Sep-2026', '03.09.2026', 'TVIP01014947'),
    ('Sep-2026', '04.09.2026', '1013143'),
    ('Sep-2026', '05.09.2026', 'TVIP01021581'),
    ('Sep-2026', '06.09.2026', 'TVIP00854813'),
    ('Sep-2026', '07.09.2026', 'TVIP01012534'),
    ('Sep-2026', '08.09.2026', 'TVIP01021984'),
    ('Sep-2026', '12.09.2026', 'TVIP01022461'),
    ('Sep-2026', '13.09.2026', 'TVIP01022562'),
    ('Sep-2026', '13.09.2026', 'TVIP01013012'),
    ('Sep-2026', '15.09.2026', 'TVIP01013619'),
    ('Sep-2026', '16.09.2026', 'TVIP01023036'),
    ('Sep-2026', '16.09.2026', 'TVIP01023039'),
    ('Sep-2026', '17.09.2026', 'TVIP01014624'),
    ('Sep-2026', '19.09.2026', 'TVIP01023366'),
    ('Sep-2026', '19.09.2026', 'TVIP01023324'),
    ('Sep-2026', '23.09.2026', 'TVIP01023924'),
    ('Sep-2026', '26.09.2026', 'TVIP01024200'),
    ('Sep-2026', '30.09.2026', 'TVIP01023036'),
]


def main(path):
    wb = openpyxl.load_workbook(path)
    for sheet, d, fno in ROWS:
        ws = wb[sheet]
        dd = datetime.strptime(d, "%d.%m.%Y")
        in_rad, hits = False, []
        for r in range(4, ws.max_row + 1):
            if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS":
                in_rad = True
            if in_rad and ws.cell(r, 2).value == dd and str(ws.cell(r, 3).value) == fno \
                    and "PREG" in str(ws.cell(r, 5).value).upper().replace("PREAG", "PREG"):
                hits.append(r)
        assert len(hits) == 1, (sheet, d, fno, hits)
        c = ws.cell(hits[0], 6)
        if c.value is None:
            c.value = PRICE
            c.number_format = "#,##0.00"
            c.comment = Comment(NOTE, "Claude")
    wb.save(path)
    print("US pregnant one-time fill:", len(ROWS), "rows")


if __name__ == "__main__":
    main(sys.argv[1])
