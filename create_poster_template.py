"""
Creates assets/poster_template.xlsx — the Poster DSR template expected by dsr_pro.py.

Columns (A–M) as referenced in apply_poster_styling():
  A: S.N.
  B: Date
  C: HQ (Headquarter)
  D: Staff Name
  E: Outlet Name
  F: Address
  G: Contact No
  H: Owner Name
  I: Xtreme Poster
  J: MT 250 Poster
  K: MT 330 Poster
  L: Bam Bam
  M: DPS Board
  N: (Verified — conditional formatted yes/no)
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = os.path.join("assets", "poster_template.xlsx")

HEADERS = [
    "S.N.",
    "Date",
    "HQ",
    "Staff Name",
    "Outlet Name",
    "Address",
    "Contact No",
    "Owner Name",
    "Xtreme Poster",
    "MT 250 Poster",
    "MT 330 Poster",
    "Bam Bam",
    "DPS Board",
    "Verified",
]

COL_WIDTHS = {
    "A": 6,   # S.N.
    "B": 12,  # Date
    "C": 18,  # HQ
    "D": 22,  # Staff Name
    "E": 30,  # Outlet Name
    "F": 28,  # Address
    "G": 15,  # Contact No
    "H": 22,  # Owner Name
    "I": 15,  # Xtreme Poster
    "J": 15,  # MT 250 Poster
    "K": 15,  # MT 330 Poster
    "L": 12,  # Bam Bam
    "M": 18,  # DPS Board
    "N": 12,  # Verified
}

def make_border():
    thin = Side(border_style="thin", color="000000")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def make_header_fill():
    return PatternFill(start_color="1D4E89", end_color="1D4E89", fill_type="solid")

def create_template():
    os.makedirs("assets", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "POSTER_TEMPLATE"

    border = make_border()
    header_fill = make_header_fill()
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Row 1: Title merge
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    title_cell.value = "POSTER DSR REPORT"
    title_cell.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0D2D5E", end_color="0D2D5E", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Row 2: Column headers
    ws.row_dimensions[2].height = 36
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    # Set column widths
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # Freeze panes below header row
    ws.freeze_panes = "A3"

    wb.save(OUTPUT_PATH)
    print(f"✅ Template created: {os.path.abspath(OUTPUT_PATH)}")

if __name__ == "__main__":
    create_template()
