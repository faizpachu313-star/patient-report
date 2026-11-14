"""Add the rows from the Nov-2026 uploads that were NOT already in the master workbook.

Run on the design-system base (before scripts/add_radiology_sections.py):
    python scripts/merge_uploads_nov2026.py <master.xlsx>

Every row below was checked against the master by date + file number and by name;
everything else in the uploads was already present (see PR description).
"""
import re
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402

# sheet -> rows (date, file, name, consultation). Source: faiz_marketing_AUG_data.xlsx,
# sheet "NOV25-TO-JUNE26 (2)" - the only rows of that sheet missing from the master.
NEW_VISITS = {
    "Jul-2026": [
        ("01.07.2026", "TVIP009213588", "Tawoor Khan", "Physio"),
        ("04.07.2026", "TVIP009213588", "Tawoor Khan", "Physio"),
    ],
    "Aug-2026": [
        ("06.08.2026", "TVIP00898744", "Nihalasherin Kaipadan Chalil", "Radiology (USG)"),
    ],
}
# Source: Spectrum_Referrals_Town_Centre_Branch.xlsx - the only 'Done' referral not already
# in the Oct-2026 radiology list (others matched by name at their scan date).
OCT_RADIOLOGY = ("24.10.2026", "18639", "Mohamed Abbas", "CT LEFT FOOT",
                 "From Spectrum Referrals - Town Centre Branch (status: Done). File no. is the "
                 "Town Centre EMR #; bill amount not given in source.")

REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def shift_formula(f, idx, only_sheet=None):
    """Shift row refs >= idx down by one (refs to `only_sheet` only, when given)."""
    if only_sheet is None:
        return REF.sub(lambda m: m.group(1) + m.group(2) + m.group(3)
                       + str(int(m.group(4)) + (int(m.group(4)) >= idx)), f)
    pat = re.compile(r"('%s'!)(\$?[A-Z]{1,3}\$?)(\d+)" % re.escape(only_sheet))
    return pat.sub(lambda m: m.group(1) + m.group(2) + str(int(m.group(3)) + (int(m.group(3)) >= idx)), f)


def insert_row(wb, ws, idx, style_row):
    merges = [r.coord for r in ws.merged_cells.ranges]
    heights = {k: d.height for k, d in ws.row_dimensions.items() if d.height is not None}
    for m in merges:
        ws.unmerge_cells(m)
    ws.insert_rows(idx)
    for m in merges:
        ws.merge_cells(shift_formula(m, idx))
    for k in list(ws.row_dimensions.keys()):
        ws.row_dimensions[k].height = None
    for k, h in heights.items():
        ws.row_dimensions[k + (k >= idx)].height = h
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = shift_formula(c.value, idx)
    for other in wb.worksheets:
        if other is ws:
            continue
        for row in other.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and f"'{ws.title}'!" in c.value:
                    c.value = shift_formula(c.value, idx, ws.title)
    src = style_row + (style_row >= idx)
    for col in range(1, ws.max_column + 1):
        s, t = ws.cell(src, col), ws.cell(idx, col)
        t.font, t.fill, t.border = copy(s.font), copy(s.fill), copy(s.border)
        t.alignment, t.number_format = copy(s.alignment), s.number_format
    ws.row_dimensions[idx].height = heights.get(style_row, ds.ROW_H["data"])


def renumber(ws, first, last):
    for i, r in enumerate(range(first, last + 1), 1):
        ws.cell(r, 1).value = i


def add_visits(wb, sheet, rows):
    ws = wb[sheet]
    total = ws.max_row
    spacer = total - 1
    for d, fno, name, svc in rows:
        dd = datetime.strptime(d, "%d.%m.%Y")
        data = list(range(4, spacer))
        used = [r for r in data if ws.cell(r, 2).value is not None]
        later = [r for r in used if ws.cell(r, 2).value > dd]
        nxt = (used[-1] + 1) if used else 4
        if not later and nxt < spacer and ws.cell(nxt, 2).value is None:
            r = nxt  # blank template row right after the last entry
        else:
            r = later[0] if later else nxt
            insert_row(wb, ws, r, 4)
            total += 1
            spacer += 1
            data = list(range(4, spacer))
        for col, v in zip("BCDE", (dd, fno, name, svc)):
            ws[f"{col}{r}"] = v
        for col in "ABCDE":
            ds.style_data(ws[f"{col}{r}"], col)
        used = [x for x in data if ws.cell(x, 2).value is not None]
    renumber(ws, 4, used[-1])
    for r in data:
        if ws.cell(r, 2).value is None:
            ws.cell(r, 1).value = None


def add_oct_radiology(wb):
    ws = wb["Oct-2026"]
    d, fno, name, exam, note = OCT_RADIOLOGY
    dd = datetime.strptime(d, "%d.%m.%Y")
    hdr = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Sl No" and r > 3)
    rows = [r for r in range(hdr + 1, ws.max_row + 1) if isinstance(ws.cell(r, 2).value, datetime)]
    idx = next(r for r in rows if ws.cell(r, 2).value > dd)
    insert_row(wb, ws, idx, rows[0])
    for col, v in zip("BCDE", (dd, fno, name, exam)):
        ws[f"{col}{idx}"] = v
    ws[f"D{idx}"].comment = Comment(note, "Claude")
    renumber(ws, hdr + 1, rows[-1] + 1)


def main(path):
    wb = openpyxl.load_workbook(path)
    for sheet, rows in NEW_VISITS.items():
        add_visits(wb, sheet, rows)
    add_oct_radiology(wb)
    s = wb["Summary"]
    s["A19"] = s["A19"].value.replace("Jul 2026 has no data yet (blank template)",
                                      "Jul 2026 and the 6-Aug USG visit from the marketing AUG file")
    wb.save(path)


if __name__ == "__main__":
    main(sys.argv[1])
