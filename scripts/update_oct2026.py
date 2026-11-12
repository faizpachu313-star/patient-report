"""Replace Oct-2026 with the 'MUHAMMED DATA - OCT 2026' workbook and restyle every
sheet of the master workbook to the design system (scripts/design_system.py).

Usage: python scripts/update_oct2026.py <oct_source.xlsx> <master.xlsx>
"""
import sys
from copy import copy
from datetime import datetime

import openpyxl

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402

src_path, master_path = sys.argv[1], sys.argv[2]
src = openpyxl.load_workbook(src_path)["Oct-2026"]
wb = openpyxl.load_workbook(master_path)

# ---- 1. Replace Oct-2026 with an exact copy of the source sheet -------------
idx = wb.sheetnames.index("Oct-2026")
wb.remove(wb["Oct-2026"])
oct_ws = wb.create_sheet("Oct-2026", idx)
for row in src.iter_rows():
    for c in row:
        t = oct_ws.cell(row=c.row, column=c.column, value=c.value)
        if c.has_style:
            t.font, t.fill, t.border = copy(c.font), copy(c.fill), copy(c.border)
            t.alignment, t.number_format = copy(c.alignment), c.number_format
for rng in src.merged_cells.ranges:
    oct_ws.merge_cells(str(rng))
for k, d in src.column_dimensions.items():
    oct_ws.column_dimensions[k].width = d.width
for k, d in src.row_dimensions.items():
    if d.height is not None:
        oct_ws.row_dimensions[k].height = d.height
oct_ws.freeze_panes = src.freeze_panes
ds.apply_page_setup(oct_ws)

# ---- 2. Restyle the other month sheets ---------------------------------------
for ws in wb.worksheets:
    if ws.title in ("Summary", "Oct-2026"):
        continue
    month_label = ws["A2"].value.replace("PATIENT DATA OF ", "")
    # locate spacer + total rows (total row = last row, spacer just above)
    total_r = ws.max_row
    spacer_r = total_r - 1
    last_col = "E"  # no bill amounts for these months

    ds.style_name_line(ws["A1"])
    ds.style_title(ws["A2"])
    for c in ws[3][:5]:
        ds.style_header(c)
    for r in range(4, spacer_r):
        b = ws.cell(r, 2)
        if isinstance(b.value, str) and b.value.strip():
            b.value = datetime.strptime(b.value.strip(), "%d %b %Y")
        for c in ws[r][:5]:
            ds.style_data(c, c.column_letter)
    for c in ws[spacer_r][:5]:
        ds.style_spacer(c)
    ws.row_dimensions[spacer_r].height = ds.ROW_H["spacer"]
    ws.cell(total_r, 1).value = f"TOTAL PATIENTS (VISITS) - {month_label}"
    for c in ws[total_r][:5]:
        ds.style_total(c, h="right" if c.column == 1 else "center")
    ws.row_dimensions[total_r].height = ds.ROW_H["total"]
    for col in "ABCDE":
        ws.column_dimensions[col].width = ds.COL_WIDTHS[col]
    ws.freeze_panes = "A4"
    ds.apply_page_setup(ws)

# ---- 3. Summary --------------------------------------------------------------
s = wb["Summary"]
for r in range(4, 16):
    for c in s[r][:3]:
        c.fill = openpyxl.styles.PatternFill()  # design system: no zebra banding
s["C15"] = "='Oct-2026'!E50"
s["A19"] = ("Note: Nov 2025 - Jun 2026 from the MAIN OG 2 workbook; Jul 2026 has no data yet "
            "(blank template); Aug & Sep from the monthly files; Oct from the "
            "'MUHAMMED DATA - OCT 2026' workbook (visits count; radiology referrals and bill "
            "amounts are on the Oct-2026 sheet).")
ds.apply_page_setup(s, print_titles=False)

wb.save(master_path)
print("saved", master_path)
