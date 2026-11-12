"""Patient-report design system.

Source of truth: the "MUHAMMED DATA - OCT 2026" workbook (Oct-2026 sheet).
Every sheet in every patient workbook must use these styles. See DESIGN_SYSTEM.md.
"""
from datetime import datetime

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

# ---- Tokens -----------------------------------------------------------------
FONT_NAME = "Times New Roman"
NAVY = "FF1F3864"        # name line text, title bar, month total bars
HEADER_BLUE = "FF2F4B7C"  # column header row
TEAL = "FF1F7A8C"        # spacer strip under data, section title bars
YELLOW = "FFFFE699"      # combined-total bars
WHITE = "FFFFFFFF"
BLACK = "FF000000"
GRID = "FF9AA5B1"        # cell border colour

DATE_FMT = r"dd\-mmm\-yyyy"
AMOUNT_FMT = "#,##0.00"

# Columns: Sl No | Date | File Number | Patient Name | Consultation | Bill Amount (AED)
COL_WIDTHS = {"A": 8, "B": 16, "C": 20, "D": 52, "E": 50, "F": 18}

ROW_H = {"name": 21.75, "title": 27.75, "header": 30.0, "data": 18.0,
         "spacer": 7.5, "total": 24.0, "section": 31.5}

_side = Side(style="thin", color=GRID)
BORDER = Border(left=_side, right=_side, top=_side, bottom=_side)


def fill(rgb):
    return PatternFill("solid", fgColor=rgb)


def font(size=11, bold=False, color=None):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color)


def align(h="center", wrap=True, v="center"):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ---- Element styles ---------------------------------------------------------
def style_name_line(c):
    c.font, c.alignment, c.border = font(12, True, NAVY), align("left"), BORDER


def style_title(c):
    c.font, c.fill, c.alignment, c.border = font(15, True, WHITE), fill(NAVY), align(), BORDER


def style_header(c):
    c.font, c.fill, c.alignment, c.border = font(11, True, WHITE), fill(HEADER_BLUE), align(), BORDER


def style_data(c, col_letter):
    """Data cell: no zebra banding, Patient Name left, amounts right, rest centred."""
    c.font, c.border, c.fill = font(11), BORDER, PatternFill()
    if col_letter == "D":
        c.alignment = align("left")
    elif col_letter == "F":
        c.alignment = align("right", wrap=False)
        c.number_format = AMOUNT_FMT
    else:
        c.alignment = align()
    if col_letter == "B" and isinstance(c.value, datetime):
        c.number_format = DATE_FMT


def style_spacer(c):
    c.fill, c.border = fill(TEAL), BORDER


def style_total(c, h="center", number_format=None):
    c.font, c.fill, c.border = font(12, True, WHITE), fill(NAVY), BORDER
    c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=(h == "center"))
    if number_format:
        c.number_format = number_format


def apply_page_setup(ws, print_titles=True):
    """A4 landscape, fit to 1 page wide, page footer, no gridlines."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.page_margins.footer = 0.5
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "Page &P of &N"
    ws.sheet_view.showGridLines = False
    if print_titles:
        ws.print_title_rows = "1:3"
