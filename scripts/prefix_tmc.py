"""Add the 'TMC' prefix to plain file numbers from Nov 2025 to Jun 2026 (visits and radiology).

Only plain 4-6 digit numbers get the prefix (e.g. 124285 -> TMC124285, "125122; 125123" ->
"TMC125122; TMC125123"). Numbers already starting with TMC, Emirates IDs (784-...) and other
IDs (e.g. 9-digit 506431213) are left unchanged.

Run after scripts/add_radiology_sections.py:
    python scripts/prefix_tmc.py <master.xlsx>
"""
import re
import sys
from datetime import datetime

import openpyxl

SHEETS = ["Nov-2025", "Dec-2025", "Jan-2026", "Feb-2026", "Mar-2026", "Apr-2026", "May-2026", "Jun-2026"]
PLAIN = re.compile(r"^\d{4,6}$")


def prefix(value):
    parts = [p.strip() for p in str(value).split(";")]
    if not all(PLAIN.match(p) for p in parts):
        return value
    return "; ".join("TMC" + p for p in parts)


def main(path):
    wb = openpyxl.load_workbook(path)
    changed = 0
    for sheet in SHEETS:
        ws = wb[sheet]
        for r in range(4, ws.max_row + 1):
            if isinstance(ws.cell(r, 1).value, int) and isinstance(ws.cell(r, 2).value, datetime):
                c = ws.cell(r, 3)
                new = prefix(c.value)
                if new != c.value:
                    c.value = new
                    changed += 1
    wb.save(path)
    print("TMC prefix added to", changed, "file numbers")


if __name__ == "__main__":
    main(sys.argv[1])
