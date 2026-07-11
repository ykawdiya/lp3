#!/usr/bin/env python3
"""Build a Cricket Auction Tracker Excel template.

Sheet 1 "Auction"  : master list -> Player | Price (Cr) | Team
Sheet 2 "Balance"   : per-team purse / spent / balance / players (source of truth for team names)
Sheets 3..17        : one sheet per team, auto-filled from the Auction sheet

All numbers are in CRORES. 1.20 = 1 crore 20 lakh, 2.00 = 2 crore, etc.
Every team starts with a purse of 100.00.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.workbook.properties import CalcProperties
from openpyxl.utils import get_column_letter

NUM_TEAMS = 15
PURSE = 100.00
DATA_LAST = 1000          # auction rows the formulas scan (2..1000)
TEAM_ROWS = 40            # max players shown per team sheet (cricket squad ~25)

# ---- palette -------------------------------------------------------------
NAVY      = "1F3A5F"
BLUE      = "2E5E8C"
LIGHTBLUE = "DCE6F1"
GREEN     = "2E7D32"
GREENFILL = "E3F2E4"
RED       = "C62828"
GOLD      = "B8860B"
GREY      = "F2F2F2"
WHITE     = "FFFFFF"

def fill(hex_):   return PatternFill("solid", fgColor=hex_)
def font(**kw):   return Font(**kw)

thin = Side(style="thin", color="BFBFBF")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
left   = Alignment(horizontal="left",   vertical="center")
right  = Alignment(horizontal="right",  vertical="center")

MONEY = "0.00"

wb = Workbook()
wb.calculation = CalcProperties(fullCalcOnLoad=True)   # force recalc on open

# =========================================================================
# SHEET 1 : AUCTION
# =========================================================================
au = wb.active
au.title = "Auction"
au.sheet_properties.tabColor = NAVY

# Title band
au.merge_cells("A1:C1")
c = au["A1"]
c.value = "🏏  CRICKET AUCTION  —  MASTER LIST"
c.font = font(bold=True, size=16, color=WHITE)
c.fill = fill(NAVY)
c.alignment = center
au.row_dimensions[1].height = 30

au.merge_cells("A2:C2")
c = au["A2"]
c.value = ("Enter every sold player below. Prices are in CRORES  (1.20 = 1cr 20L, "
           "2.00 = 2cr).  Pick the team from the drop-down — the team sheets & "
           "Balance update automatically.")
c.font = font(italic=True, size=9, color="555555")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
c.fill = fill(GREY)
au.row_dimensions[2].height = 28

# Header row (row 4)
hdr = ["Player", "Price (Cr)", "Team"]
for i, h in enumerate(hdr, start=1):
    cell = au.cell(row=4, column=i, value=h)
    cell.font = font(bold=True, color=WHITE)
    cell.fill = fill(BLUE)
    cell.alignment = center
    cell.border = border_all

au.freeze_panes = "A5"
au.column_dimensions["A"].width = 30
au.column_dimensions["B"].width = 14
au.column_dimensions["C"].width = 16

# Example rows so the template visibly works (safe to overwrite / delete)
examples = [
    ("Player_X", 2.00, "Team 4"),
    ("Rohit S.", 3.50, "Team 1"),
    ("Bumrah",   4.20, "Team 4"),
    ("Kohli",    5.00, "Team 7"),
    ("Jadeja",   1.20, "Team 1"),
]
first_data = 5
for r, (p, price, team) in enumerate(examples, start=first_data):
    au.cell(row=r, column=1, value=p)
    pc = au.cell(row=r, column=2, value=price); pc.number_format = MONEY
    au.cell(row=r, column=3, value=team)

# style the whole data-entry area
for r in range(first_data, DATA_LAST + 1):
    a = au.cell(row=r, column=1); a.border = border_all; a.alignment = left
    b = au.cell(row=r, column=2); b.border = border_all; b.alignment = right; b.number_format = MONEY
    cc = au.cell(row=r, column=3); cc.border = border_all; cc.alignment = center
    if r % 2 == 0:
        for col in (1, 2, 3):
            au.cell(row=r, column=col).fill = fill(GREY)

# Data-validation drop-down for the Team column (list = Balance!A2:A16)
dv = DataValidation(type="list",
                    formula1="=Balance!$A$2:$A$%d" % (NUM_TEAMS + 1),
                    allow_blank=True, showDropDown=False)
dv.error = "Pick a team from the list."
dv.errorTitle = "Invalid team"
dv.prompt = "Choose the buying team"
dv.promptTitle = "Team"
au.add_data_validation(dv)
dv.add("C%d:C%d" % (first_data, DATA_LAST))

# little running totals top-right for convenience
au["E4"] = "Total sold"
au["E4"].font = font(bold=True)
au["F4"] = "=COUNTA(A%d:A%d)" % (first_data, DATA_LAST)
au["E5"] = "Total spend (Cr)"
au["E5"].font = font(bold=True)
au["F5"] = "=SUM(B%d:B%d)" % (first_data, DATA_LAST)
au["F5"].number_format = MONEY
au.column_dimensions["E"].width = 16
au.column_dimensions["F"].width = 12

# =========================================================================
# SHEET 2 : BALANCE  (source of truth for team names)
# =========================================================================
bal = wb.create_sheet("Balance")
bal.sheet_properties.tabColor = GOLD

bal.merge_cells("A1:E1")
c = bal["A1"]
c.value = "💰  TEAM PURSE & BALANCE"
c.font = font(bold=True, size=16, color=WHITE)
c.fill = fill(GOLD)
c.alignment = center
bal.row_dimensions[1].height = 30

bal.merge_cells("A2:E2")
c = bal["A2"]
c.value = ("Rename a team here and it updates everywhere (drop-down + that team's sheet). "
           "Purse is editable per team. All figures in crores.")
c.font = font(italic=True, size=9, color="555555")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
c.fill = fill(GREY)
bal.row_dimensions[2].height = 24

bal_hdr = ["Team", "Purse (Cr)", "Spent (Cr)", "Balance (Cr)", "Players"]
for i, h in enumerate(bal_hdr, start=1):
    cell = bal.cell(row=3, column=i, value=h)
    cell.font = font(bold=True, color=WHITE)
    cell.fill = fill(BLUE)
    cell.alignment = center
    cell.border = border_all

for t in range(1, NUM_TEAMS + 1):
    r = 3 + t
    bal.cell(row=r, column=1, value="Team %d" % t).alignment = left
    p = bal.cell(row=r, column=2, value=PURSE); p.number_format = MONEY; p.alignment = right
    s = bal.cell(row=r, column=3,
                 value="=SUMIF(Auction!$C$%d:$C$%d,$A%d,Auction!$B$%d:$B$%d)"
                       % (first_data, DATA_LAST, r, first_data, DATA_LAST))
    s.number_format = MONEY; s.alignment = right
    b = bal.cell(row=r, column=4, value="=B%d-C%d" % (r, r))
    b.number_format = MONEY; b.alignment = right
    n = bal.cell(row=r, column=5,
                 value="=COUNTIF(Auction!$C$%d:$C$%d,$A%d)" % (first_data, DATA_LAST, r))
    n.alignment = center
    for col in range(1, 6):
        bal.cell(row=r, column=col).border = border_all
    if t % 2 == 0:
        for col in range(1, 6):
            bal.cell(row=r, column=col).fill = fill(GREY)

# totals row
tr = 3 + NUM_TEAMS + 1
bal.cell(row=tr, column=1, value="TOTAL").font = font(bold=True)
for col, letter in ((2, "B"), (3, "C"), (4, "D"), (5, "E")):
    cell = bal.cell(row=tr, column=col,
                    value="=SUM(%s4:%s%d)" % (letter, letter, tr - 1))
    cell.font = font(bold=True)
    cell.number_format = MONEY if col != 5 else "0"
    cell.alignment = right if col != 5 else center
    cell.fill = fill(LIGHTBLUE)
    cell.border = border_all
bal.cell(row=tr, column=1).fill = fill(LIGHTBLUE)
bal.cell(row=tr, column=1).border = border_all

bal.freeze_panes = "A4"
bal.column_dimensions["A"].width = 20
for col in ("B", "C", "D", "E"):
    bal.column_dimensions[col].width = 13

# conditional-ish highlight: colour the Balance column via number format only
# (leave real conditional formatting out for max compatibility)

# =========================================================================
# SHEETS 3..17 : ONE PER TEAM
# =========================================================================
for t in range(1, NUM_TEAMS + 1):
    ws = wb.create_sheet("Team %d" % t)
    ws.sheet_properties.tabColor = BLUE
    bal_row = 3 + t   # matching row on Balance sheet

    # Title = live reference to Balance team name
    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value = "=Balance!A%d" % bal_row
    c.font = font(bold=True, size=15, color=WHITE)
    c.fill = fill(NAVY)
    c.alignment = center
    ws.row_dimensions[1].height = 28

    # stat block (pulled from Balance -> single source of truth)
    stats = [
        ("Purse (Cr)",   "=Balance!B%d" % bal_row, LIGHTBLUE),
        ("Spent (Cr)",   "=Balance!C%d" % bal_row, LIGHTBLUE),
        ("Balance (Cr)", "=Balance!D%d" % bal_row, GREENFILL),
        ("Players",      "=Balance!E%d" % bal_row, GREY),
    ]
    for i, (label, formula, bg) in enumerate(stats):
        r = 3 + i
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = font(bold=True); lc.fill = fill(bg)
        lc.alignment = left; lc.border = border_all
        vc = ws.cell(row=r, column=2, value=formula)
        vc.fill = fill(bg); vc.alignment = right; vc.border = border_all
        vc.number_format = MONEY if label != "Players" else "0"
        if label == "Balance (Cr)":
            vc.font = font(bold=True, color=GREEN)

    # squad table header
    hr = 8
    for i, h in enumerate(["Player", "Price (Cr)"], start=1):
        cell = ws.cell(row=hr, column=i, value=h)
        cell.font = font(bold=True, color=WHITE)
        cell.fill = fill(BLUE); cell.alignment = center; cell.border = border_all

    # auto-fill players bought by this team via CSE array formulas
    # (INDEX/SMALL/IF -> works in Excel 2007+, Google Sheets, LibreOffice)
    start = hr + 1
    for k in range(TEAM_ROWS):
        r = start + k
        cnt = "ROWS($A$%d:A%d)" % (start, r)
        rng_c = "Auction!$C$%d:$C$%d" % (first_data, DATA_LAST)
        base  = "ROW(Auction!$C$%d:$C$%d)-ROW(Auction!$C$%d)+1" % (first_data, DATA_LAST, first_data)
        small = "SMALL(IF(%s=$A$1,%s),%s)" % (rng_c, base, cnt)
        fa = "=IFERROR(INDEX(Auction!$A$%d:$A$%d,%s),\"\")" % (first_data, DATA_LAST, small)
        fb = "=IFERROR(INDEX(Auction!$B$%d:$B$%d,%s),\"\")" % (first_data, DATA_LAST, small)
        ca = ws.cell(row=r, column=1)
        ca.value = ArrayFormula("A%d" % r, fa)
        ca.alignment = left; ca.border = border_all
        cb = ws.cell(row=r, column=2)
        cb.value = ArrayFormula("B%d" % r, fb)
        cb.alignment = right; cb.border = border_all; cb.number_format = MONEY
        if k % 2 == 1:
            ca.fill = fill(GREY); cb.fill = fill(GREY)

    ws.freeze_panes = "A%d" % start
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 14

# =========================================================================
out = "Cricket_Auction_Tracker.xlsx"
wb.save(out)
print("saved:", out)
print("sheets:", wb.sheetnames)
