"""Remove double entries: same patient + same date (+ same doctor/visit) in the same sheet section.

Different dates are separate visits and are always kept.
Run on the design-system base, before merge_uploads_*.py and add_radiology_sections.py:
    python scripts/remove_duplicates.py <master.xlsx>
"""
import re
import sys
from datetime import datetime

import openpyxl

REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")

# (sheet, date, file no) -> how to resolve. Rows are found by content, not position.
DUPLICATES = [
    # identical rows: keep the first, delete the repeat
    ("Dec-2025", "14.12.2025", "125131", None),
    ("Dec-2025", "14.12.2025", "125134", None),  # repeat has the cut-off name "Mohamma"
    # same visit entered twice (doctor names vs departments): keep one row with both
    ("Feb-2026", "23.02.2026", "TMC128585", "GP, Orthopedic, Lab, X-Ray (Dr Afsheen, Dr Agnel)"),
]
# same patient, same day, entered once per file number: keep one row listing both numbers
NAME_DUPLICATES = [
    ("Dec-2025", "14.12.2025", "Mohd Zaseem", "125122; 125123"),
]


def shift(f, idx, only_sheet=None):
    """Rows after the deleted row `idx` move up by one."""
    fix = lambda n: str(n - 1 if n > idx else n)  # noqa: E731
    if only_sheet is None:
        return REF.sub(lambda m: m.group(1) + m.group(2) + m.group(3) + fix(int(m.group(4))), f)
    pat = re.compile(r"('%s'!)(\$?[A-Z]{1,3}\$?)(\d+)" % re.escape(only_sheet))
    return pat.sub(lambda m: m.group(1) + m.group(2) + fix(int(m.group(3))), f)


def delete_row(wb, ws, idx):
    merges = [r.coord for r in ws.merged_cells.ranges]
    heights = {k: d.height for k, d in ws.row_dimensions.items() if d.height is not None}
    for m in merges:
        ws.unmerge_cells(m)
    ws.delete_rows(idx)
    for m in merges:
        ws.merge_cells(shift(m, idx))
    for k in list(ws.row_dimensions.keys()):
        ws.row_dimensions[k].height = None
    for k, h in heights.items():
        if k != idx:
            ws.row_dimensions[k - (k > idx)].height = h
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = shift(c.value, idx)
    for other in wb.worksheets:
        if other is ws:
            continue
        for row in other.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and f"'{ws.title}'!" in c.value:
                    c.value = shift(c.value, idx, ws.title)


def main(path):
    wb = openpyxl.load_workbook(path)
    for sheet, d, fno, merged_service in DUPLICATES:
        ws = wb[sheet]
        dd = datetime.strptime(d, "%d.%m.%Y")
        rows = [r for r in range(4, ws.max_row + 1)
                if ws.cell(r, 2).value == dd and str(ws.cell(r, 3).value) == fno]
        assert len(rows) == 2, (sheet, d, fno, rows)
        keep, drop = rows
        if merged_service:
            keep, drop = drop, keep  # keep the row with departments, drop the doctor-only row
            ws.cell(keep, 5).value = merged_service
        delete_row(wb, ws, drop)
        print(f"{sheet}: removed duplicate of {fno} on {d} (row {drop})")
    for sheet, d, name, files in NAME_DUPLICATES:
        ws = wb[sheet]
        dd = datetime.strptime(d, "%d.%m.%Y")
        rows = [r for r in range(4, ws.max_row + 1)
                if ws.cell(r, 2).value == dd and ws.cell(r, 4).value == name]
        assert len(rows) == 2, (sheet, d, name, rows)
        ws.cell(rows[0], 3).value = files
        delete_row(wb, ws, rows[1])
        print(f"{sheet}: merged duplicate {name} on {d} (row {rows[1]})")
    for sheet in {s for s, *_ in DUPLICATES + NAME_DUPLICATES}:
        ws = wb[sheet]
        n = 0
        for r in range(4, ws.max_row + 1):
            if isinstance(ws.cell(r, 2).value, datetime):
                n += 1
                ws.cell(r, 1).value = n
            else:
                break
    wb.save(path)


if __name__ == "__main__":
    main(sys.argv[1])
