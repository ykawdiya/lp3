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

# ---- real bid data (from Auction_Sheet_Invictus) -------------------------
# (team name, [(player, amount_in_Cr), ...])   amounts verified vs source totals
TEAMS_DATA = [
    ("Mumbai Indians", [
        ("Rohit Sharma", 26.5), ("Kartik Tyagi", 12), ("Kuldeep Yadav", 20),
        ("Hugh Weibgen", 1), ("Matthew Potts", 5), ("Chamika Karunaratne", 0.5),
        ("Karun Nair", 5.75), ("Sachin Baby", 1.75), ("Anrich Nortje", 7)]),
    ("Royal Challangers Banglore", [
        ("Navdeep Saini", 2.25), ("Harnoor Singh", 2.25), ("Swastik Chikara", 7.5),
        ("Tilak Varma", 22.5), ("Dawid Malan", 7), ("Haris Rauf", 8.5),
        ("George Scrimshaw", 0.5), ("Hasan Ali", 3.5), ("Usma Mir", 2.25),
        ("Jimmy Neesham", 6), ("Parth Rekhade", 1.25), ("Dinesh Bana", 1),
        ("Yuzvendra Chahal", 12)]),
    ("Chennai Super Kings", [
        ("Sandeep Sharma", 8), ("Sarafaz Khan", 13), ("Arjun Tendulkar", 6.25),
        ("Ruturaj Gaikwad", 18.5), ("Romario Shepherd", 8.5), ("Tom Latham", 4.25),
        ("Shaheen Shah Afridi", 6.25), ("Naseem Shah", 7), ("Mohammad Amir", 2.25),
        ("Akhil", 0.5)]),
    ("Kolkata Knight Riders", [
        ("Devdutt Padikkal", 19.5), ("Vaibhav Suryavanshi", 16.5), ("Sikandar Raza", 7),
        ("Ashutosh Sharma", 10.5), ("Dharmendrasinh Jadeja", 3.25), ("Sameer Rizvi", 8.75),
        ("T. Natarajan", 8.25)]),
    ("Patna Panthers", [
        ("Umesh Yadav", 4.25), ("Liam Livingstone", 13), ("Rahul Tripathi", 7.25),
        ("Suryakumar Yadav", 21.5), ("Tim Seifert", 9), ("Spencer Johnson", 5.75)]),
    ("Sunrisers Hydrabad", [
        ("Lhuan-dré Pretorius", 3), ("KL Rahul", 29), ("Angkrish Raghuvanshi", 12.5),
        ("Mohammed Siraj", 14.5), ("Baba Indrajith", 2.5)]),
    ("Delhi Capitals", [
        ("Nitish Rana", 11), ("Ben McKinney", 6.5), ("Keshav Maharaj", 8.25),
        ("Abrar Ahmed", 8), ("Faf Du Plesis", 12.5), ("Rajat Patidar", 18.5)]),
    ("Gujarat Titans", [
        ("Rinku Singh", 14.5), ("Sai Sudharsan", 23.5), ("Ishaan Kishaan", 24.5),
        ("Govinda Poddar", 0.5)]),
    ("Lucknow Super Giants", [
        ("Sanju Samson", 19.5), ("Dewald Brevis", 10), ("Mayank Dagar", 1.5),
        ("Abhishek Sharma", 16)]),
    ("Imphal Igniter", [
        ("Shaik Rasheed", 5.5), ("Chetan Sakariya", 6.25), ("Quinton de Kock", 19),
        ("Joe Root", 5.5), ("Virat Kohli", 27.5), ("Vidwath Kaverappa", 2),
        ("Siddharth Desai", 2.75)]),
    ("Punjab Kings", [
        ("Ishant Sharma", 5.5), ("Deepak Hooda", 4.5), ("Sam Konstas", 6),
        ("Jitesh Sharma", 10.25), ("Shubman Gill", 28), ("Mitchell Starc", 13)]),
    ("Rising Pune Super Giants", [
        ("Jasprit Bumrah", 24.25), ("Moeen Ali", 10), ("Will Jacks", 15.5),
        ("Dasun Shanaka", 5.75), ("Amit Shukla", 1.75), ("Abhishek Powel", 8.5)]),
    ("Rajasthan Royals", [
        ("Harpreet Bhatia", 1), ("Musheer Khan", 6.5), ("Avesh Khan", 14.5),
        ("MS Dhoni", 15), ("Glenn Maxwell", 12.5), ("Tom Bruce", 0.5),
        ("Harry Dixon", 0.5), ("Ajinkya Rahane", 8.75), ("Priyansh Arya", 16.5)]),
    ("Kochi Tuskers", [
        ("Jason Holder", 14.5), ("Kagiso Rabada", 22), ("Heinrich Klassen", 30),
        ("Nikhil Gangta", 0.5)]),
    ("Guwahati Chargers", [
        ("Hardik Pandya", 26), ("Shakib Al Hasan", 9), ("Nehal Wadhera", 6),
        ("Rishabh Pant", 17), ("Wiaan Mulder", 4.75), ("Maheesh Theekshana", 5.75)]),
]
TEAMS = [t[0] for t in TEAMS_DATA]
assert len(TEAMS_DATA) == NUM_TEAMS
# flattened auction rows: (player, price, team)
AUCTION_ROWS = [(pl, amt, name) for name, roster in TEAMS_DATA for pl, amt in roster]

# ---- player metadata (from IPL PLAYERS LIST) -----------------------------
# player -> (Role, Hand, Country, City).  Role/Hand from Grade A/B/C batting-
# or bowling-style; Country/City only exist in the list's Grade D section.
# Blanks = not present in the provided players list.
PLAYER_META = {
    'Rohit Sharma': ('Batsman', 'Right', '', ''),
    'Kartik Tyagi': ('Bowler', 'Right', '', ''),
    'Kuldeep Yadav': ('Bowler', 'Left', '', ''),
    'Hugh Weibgen': ('Batsman', '', 'Australia', ''),
    'Matthew Potts': ('Bowler', '', 'England', ''),
    'Chamika Karunaratne': ('All Rounder', 'Right', 'Sri Lanka', ''),
    'Karun Nair': ('Batsman', 'Right', '', ''),
    'Sachin Baby': ('Batsman', '', 'India', 'Kerala'),
    'Anrich Nortje': ('Bowler', 'Right', '', ''),
    'Navdeep Saini': ('Bowler', 'Right', '', ''),
    'Harnoor Singh': ('Batsman', 'Left', '', ''),
    'Swastik Chikara': ('Batsman', 'Right', '', ''),
    'Tilak Varma': ('Batsman', 'Left', '', ''),
    'Dawid Malan': ('Batsman', '', 'England', ''),
    'Haris Rauf': ('Bowler', '', 'Pakistan', ''),
    'George Scrimshaw': ('Bowler', '', 'England', ''),
    'Hasan Ali': ('Bowler', '', 'Pakistan', ''),
    'Usma Mir': ('Bowler', '', 'Pakistan', ''),
    'Jimmy Neesham': ('All Rounder', 'Left', 'New Zealand', ''),
    'Parth Rekhade': ('All Rounder', '', 'India', 'Vidarbha'),
    'Dinesh Bana': ('Wicketkeeper', '', 'India', 'Haryana'),
    'Yuzvendra Chahal': ('Bowler', 'Right', '', ''),
    'Sandeep Sharma': ('Bowler', 'Right', '', ''),
    'Sarafaz Khan': ('Batsman', 'Right', '', ''),
    'Arjun Tendulkar': ('Bowler', 'Left', '', ''),
    'Ruturaj Gaikwad': ('Batsman', 'Right', '', ''),
    'Romario Shepherd': ('All Rounder', 'Right', '', ''),
    'Tom Latham': ('Wicketkeeper', '', 'New Zealand', ''),
    'Shaheen Shah Afridi': ('Bowler', '', 'Pakistan', ''),
    'Naseem Shah': ('Bowler', '', 'Pakistan', ''),
    'Mohammad Amir': ('', '', '', ''),
    'Akhil': ('', '', '', ''),
    'Devdutt Padikkal': ('Batsman', 'Left', '', ''),
    'Vaibhav Suryavanshi': ('Batsman', 'Left', '', ''),
    'Sikandar Raza': ('All Rounder', 'Right', 'Zimbabwe', ''),
    'Ashutosh Sharma': ('Batsman', 'Right', '', ''),
    'Dharmendrasinh Jadeja': ('Bowler', '', 'India', 'Saurashtra'),
    'Sameer Rizvi': ('Batsman', 'Right', '', ''),
    'T. Natarajan': ('Bowler', 'Left', '', ''),
    'Umesh Yadav': ('Bowler', 'Right', '', ''),
    'Liam Livingstone': ('All Rounder', 'Right', '', ''),
    'Rahul Tripathi': ('Batsman', 'Right', '', ''),
    'Suryakumar Yadav': ('Batsman', 'Right', '', ''),
    'Tim Seifert': ('Wicketkeeper', '', 'New Zealand', ''),
    'Spencer Johnson': ('Bowler', 'Left', '', ''),
    'Lhuan-dré Pretorius': ('Batsman', '', 'South Africa', ''),
    'KL Rahul': ('Wicketkeeper', 'Right', '', ''),
    'Angkrish Raghuvanshi': ('Batsman', 'Right', '', ''),
    'Mohammed Siraj': ('Bowler', 'Right', '', ''),
    'Baba Indrajith': ('Batsman', '', 'India', 'Tamil Nadu'),
    'Nitish Rana': ('Batsman', 'Left', '', ''),
    'Ben McKinney': ('Batsman', '', 'England', ''),
    'Keshav Maharaj': ('Bowler', '', 'South Africa', ''),
    'Abrar Ahmed': ('Bowler', '', 'Pakistan', ''),
    'Faf Du Plesis': ('Batsman', 'Right', '', ''),
    'Rajat Patidar': ('Batsman', 'Right', '', ''),
    'Rinku Singh': ('Batsman', 'Left', '', ''),
    'Sai Sudharsan': ('Batsman', 'Left', '', ''),
    'Ishaan Kishaan': ('Wicketkeeper', 'Left', '', ''),
    'Govinda Poddar': ('All Rounder', '', 'India', 'Odisha'),
    'Sanju Samson': ('Wicketkeeper', 'Right', '', ''),
    'Dewald Brevis': ('Batsman', 'Right', '', ''),
    'Mayank Dagar': ('All Rounder', '', 'India', 'Himachal Pradesh'),
    'Abhishek Sharma': ('Batsman', 'Left', '', ''),
    'Shaik Rasheed': ('Batsman', 'Right', '', ''),
    'Chetan Sakariya': ('Bowler', 'Left', '', ''),
    'Quinton de Kock': ('Wicketkeeper', 'Left', '', ''),
    'Joe Root': ('Batsman', '', 'England', ''),
    'Virat Kohli': ('Batsman', 'Right', '', ''),
    'Vidwath Kaverappa': ('Bowler', 'Right', '', ''),
    'Siddharth Desai': ('Bowler', '', 'India', 'Gujarat'),
    'Ishant Sharma': ('Bowler', 'Right', '', ''),
    'Deepak Hooda': ('All Rounder', 'Right', '', ''),
    'Sam Konstas': ('Batsman', '', 'Australia', ''),
    'Jitesh Sharma': ('Wicketkeeper', 'Right', '', ''),
    'Shubman Gill': ('Batsman', 'Right', '', ''),
    'Mitchell Starc': ('Bowler', 'Left', '', ''),
    'Jasprit Bumrah': ('Bowler', 'Right', '', ''),
    'Moeen Ali': ('All Rounder', 'Left', '', ''),
    'Will Jacks': ('All Rounder', 'Right', '', ''),
    'Dasun Shanaka': ('All Rounder', 'Right', 'Sri Lanka', ''),
    'Amit Shukla': ('All Rounder', '', 'India', 'Services'),
    'Abhishek Powel': ('', '', '', ''),
    'Harpreet Bhatia': ('Batsman', 'Left', '', ''),
    'Musheer Khan': ('Batsman', 'Right', '', ''),
    'Avesh Khan': ('Bowler', 'Right', '', ''),
    'MS Dhoni': ('Wicketkeeper', 'Right', '', ''),
    'Glenn Maxwell': ('All Rounder', 'Right', '', ''),
    'Tom Bruce': ('Batsman', '', 'New Zealand', ''),
    'Harry Dixon': ('Batsman', '', 'Australia', ''),
    'Ajinkya Rahane': ('Batsman', 'Right', '', ''),
    'Priyansh Arya': ('Batsman', 'Left', '', ''),
    'Jason Holder': ('All Rounder', 'Right', 'West Indies', ''),
    'Kagiso Rabada': ('Bowler', 'Right', '', ''),
    'Heinrich Klassen': ('Wicketkeeper', 'Right', '', ''),
    'Nikhil Gangta': ('Batsman', '', 'India', 'Himachal Pradesh'),
    'Hardik Pandya': ('All Rounder', 'Right', '', ''),
    'Shakib Al Hasan': ('All Rounder', 'Left', 'Bangladesh', ''),
    'Nehal Wadhera': ('Batsman', 'Left', '', ''),
    'Rishabh Pant': ('Wicketkeeper', 'Left', '', ''),
    'Wiaan Mulder': ('All Rounder', 'Right', 'South Africa', ''),
    'Maheesh Theekshana': ('Bowler', 'Right', '', ''),
}

wb = Workbook()
wb.calculation = CalcProperties(fullCalcOnLoad=True)   # force recalc on open

# =========================================================================
# SHEET 1 : AUCTION
# =========================================================================
au = wb.active
au.title = "Auction"
au.sheet_properties.tabColor = NAVY

# Title band
au.merge_cells("A1:D1")
c = au["A1"]
c.value = "🏏  CRICKET AUCTION  —  MASTER LIST"
c.font = font(bold=True, size=16, color=WHITE)
c.fill = fill(NAVY)
c.alignment = center
au.row_dimensions[1].height = 30

au.merge_cells("A2:D2")
c = au["A2"]
c.value = ("Enter every sold player below. Prices are in CRORES  (1.20 = 1cr 20L, "
           "2.00 = 2cr).  Pick the team from the drop-down. Role fills in from the "
           "Players sheet. Use the filter arrows to group by Role or Team.")
c.font = font(italic=True, size=9, color="555555")
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
c.fill = fill(GREY)
au.row_dimensions[2].height = 28

# Header row (row 4)
hdr = ["Player", "Price (Cr)", "Team", "Role"]
for i, h in enumerate(hdr, start=1):
    cell = au.cell(row=4, column=i, value=h)
    cell.font = font(bold=True, color=WHITE)
    cell.fill = fill(BLUE)
    cell.alignment = center
    cell.border = border_all

au.freeze_panes = "A5"
au.column_dimensions["A"].width = 30
au.column_dimensions["B"].width = 14
au.column_dimensions["C"].width = 22
au.column_dimensions["D"].width = 15

# Real bid data (grouped by team, in roster order)
first_data = 5
for r, (p, price, team) in enumerate(AUCTION_ROWS, start=first_data):
    au.cell(row=r, column=1, value=p)
    pc = au.cell(row=r, column=2, value=price); pc.number_format = MONEY
    au.cell(row=r, column=3, value=team)

last_data = first_data + len(AUCTION_ROWS) - 1

# style the data-entry area; Role (col D) is looked up from the Players sheet
for r in range(first_data, DATA_LAST + 1):
    a = au.cell(row=r, column=1); a.border = border_all; a.alignment = left
    b = au.cell(row=r, column=2); b.border = border_all; b.alignment = right; b.number_format = MONEY
    cc = au.cell(row=r, column=3); cc.border = border_all; cc.alignment = center
    d = au.cell(row=r, column=4,
                value='=IFERROR(T(VLOOKUP($A%d,Players!$A:$E,2,FALSE)),"")' % r)
    d.border = border_all; d.alignment = center
    if r % 2 == 0:
        for col in (1, 2, 3, 4):
            au.cell(row=r, column=col).fill = fill(GREY)

# filter arrows over the header so bids can be grouped/sorted by Role or Team
au.auto_filter.ref = "A4:D%d" % last_data

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
    bal.cell(row=r, column=1, value=TEAMS[t - 1]).alignment = left
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
def tab_name(name):
    # Excel tab: <=31 chars, none of : \ / ? * [ ]
    safe = name.translate({ord(ch): " " for ch in ':\\/?*[]'})
    return safe[:31]

for t in range(1, NUM_TEAMS + 1):
    ws = wb.create_sheet(tab_name(TEAMS[t - 1]))
    ws.sheet_properties.tabColor = BLUE
    bal_row = 3 + t   # matching row on Balance sheet

    # Title = live reference to Balance team name
    ws.merge_cells("A1:F1")
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
    squad_hdr = ["Player", "Price (Cr)", "Role", "Hand", "Country", "City"]
    for i, h in enumerate(squad_hdr, start=1):
        cell = ws.cell(row=hr, column=i, value=h)
        cell.font = font(bold=True, color=WHITE)
        cell.fill = fill(BLUE); cell.alignment = center; cell.border = border_all

    # auto-fill players bought by this team via CSE array formulas
    # (INDEX/SMALL/IF -> works in Excel 2007+, Google Sheets, LibreOffice)
    # Role/Hand/Country/City are VLOOKUP'd from the Players sheet by player name.
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
        # Role / Hand / Country / City from Players sheet (cols 2..5)
        for col, vidx, algn in ((3, 2, center), (4, 3, center), (5, 4, left), (6, 5, left)):
            cc = ws.cell(row=r, column=col,
                         value='=IFERROR(T(VLOOKUP($A%d,Players!$A:$E,%d,FALSE)),"")' % (r, vidx))
            cc.alignment = algn; cc.border = border_all
        if k % 2 == 1:
            for col in range(1, 7):
                ws.cell(row=r, column=col).fill = fill(GREY)

    ws.freeze_panes = "A%d" % start
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 15

# =========================================================================
# SUMMARY : all 15 teams in one sheet, laid out like the uploaded file
#           (3 row-blocks x 5 column-blocks), fully linked to the Auction data
# =========================================================================
summ = wb.create_sheet("Summary", index=2)   # after Auction & Balance
summ.sheet_properties.tabColor = GREEN
SLOTS = 15  # player rows per block (matches the uploaded sheet)

# column trios for the 5 blocks: (S.No col, Player col, Amount col)
COL_TRIOS = [(2, 3, 4), (6, 7, 8), (10, 11, 12), (14, 15, 16), (18, 19, 20)]
# per row-block: (name_row, header_row, first_player_row, spent_row, left_row, total_row)
ROW_BLOCKS = [(2, 3, 4, 19, 20, 21),
              (23, 24, 25, 40, 41, 42),
              (44, 45, 46, 61, 62, 63)]

rng_c = "Auction!$C$%d:$C$%d" % (first_data, DATA_LAST)
base  = "ROW(Auction!$C$%d:$C$%d)-ROW(Auction!$C$%d)+1" % (first_data, DATA_LAST, first_data)

# column widths (narrow spacers between blocks)
summ.column_dimensions["A"].width = 2
for (sno, pl, amt) in COL_TRIOS:
    summ.column_dimensions[get_column_letter(sno)].width = 5
    summ.column_dimensions[get_column_letter(pl)].width = 22
    summ.column_dimensions[get_column_letter(amt)].width = 9
for spacer in (5, 9, 13, 17):
    summ.column_dimensions[get_column_letter(spacer)].width = 2

for idx in range(NUM_TEAMS):
    g, cpos = idx // 5, idx % 5
    sno_c, pl_c, amt_c = COL_TRIOS[cpos]
    name_r, hdr_r, p0, spent_r, left_r, total_r = ROW_BLOCKS[g]
    bal_row = 4 + idx
    L_sno, L_pl, L_amt = (get_column_letter(sno_c), get_column_letter(pl_c),
                          get_column_letter(amt_c))
    name_cell = "$%s$%d" % (L_sno, name_r)   # criteria = team name cell

    # --- team name (merged across the 3 columns), linked to Balance ---
    summ.merge_cells(start_row=name_r, start_column=sno_c, end_row=name_r, end_column=amt_c)
    nc = summ.cell(name_r, sno_c, "=Balance!$A$%d" % bal_row)
    nc.font = font(bold=True, size=12, color=WHITE); nc.fill = fill(NAVY); nc.alignment = center
    summ.row_dimensions[name_r].height = 22

    # --- header row ---
    for cc, htxt in ((sno_c, "S.No"), (pl_c, "Player Name"), (amt_c, "Amount")):
        h = summ.cell(hdr_r, cc, htxt)
        h.font = font(bold=True, color=WHITE); h.fill = fill(BLUE)
        h.alignment = center; h.border = border_all

    # --- player rows (S.No static 1..15; name/amount auto-filled) ---
    for k in range(SLOTS):
        r = p0 + k
        small = "SMALL(IF(%s=%s,%s),%d)" % (rng_c, name_cell, base, k + 1)
        fa = '=IFERROR(INDEX(Auction!$A$%d:$A$%d,%s),"")' % (first_data, DATA_LAST, small)
        fb = '=IFERROR(INDEX(Auction!$B$%d:$B$%d,%s),"")' % (first_data, DATA_LAST, small)
        sc = summ.cell(r, sno_c, k + 1); sc.alignment = center; sc.border = border_all
        pc = summ.cell(r, pl_c); pc.value = ArrayFormula("%s%d" % (L_pl, r), fa)
        pc.alignment = left; pc.border = border_all
        ac = summ.cell(r, amt_c); ac.value = ArrayFormula("%s%d" % (L_amt, r), fb)
        ac.alignment = right; ac.border = border_all; ac.number_format = MONEY
        if k % 2 == 1:
            sc.fill = fill(GREY); pc.fill = fill(GREY); ac.fill = fill(GREY)

    # --- totals block (label spans S.No+Player cols; value in Amount col) ---
    totals = [(spent_r, "Total Amount Spent", "=Balance!$C$%d" % bal_row, GREY),
              (left_r,  "Total Amount Left",  "=Balance!$D$%d" % bal_row, GREENFILL),
              (total_r, "Total Amount",       "=Balance!$B$%d" % bal_row, LIGHTBLUE)]
    for rr, label, formula, bg in totals:
        summ.merge_cells(start_row=rr, start_column=sno_c, end_row=rr, end_column=pl_c)
        lc = summ.cell(rr, sno_c, label)
        lc.font = font(bold=True); lc.alignment = right; lc.fill = fill(bg); lc.border = border_all
        vc = summ.cell(rr, amt_c, formula)
        vc.font = font(bold=True); vc.number_format = MONEY; vc.alignment = right
        vc.fill = fill(bg); vc.border = border_all

summ.sheet_view.showGridLines = False

# =========================================================================
# PLAYERS : metadata master (single source for Role / Hand / Country / City)
#           referenced by VLOOKUP from the Auction & team sheets.
# =========================================================================
pl = wb.create_sheet("Players")
pl.sheet_properties.tabColor = "555555"
pl.merge_cells("A1:E1")
c = pl["A1"]
c.value = "PLAYER INFO  (edit here — Role/Hand/Country/City feed the other sheets)"
c.font = font(bold=True, size=12, color=WHITE); c.fill = fill("555555"); c.alignment = center
pl.row_dimensions[1].height = 24
for i, h in enumerate(["Player", "Role", "Hand", "Country", "City"], start=1):
    cell = pl.cell(row=2, column=i, value=h)
    cell.font = font(bold=True, color=WHITE); cell.fill = fill(BLUE)
    cell.alignment = center; cell.border = border_all
# rows in auction order; blanks where the players list had no info
pr = 3
for player, _price, _team in AUCTION_ROWS:
    role, hand, country, city = PLAYER_META.get(player, ("", "", "", ""))
    for col, v, algn in ((1, player, left), (2, role, center), (3, hand, center),
                         (4, country, left), (5, city, left)):
        cell = pl.cell(row=pr, column=col, value=v)
        cell.alignment = algn; cell.border = border_all
    if pr % 2 == 0:
        for col in range(1, 6):
            pl.cell(row=pr, column=col).fill = fill(GREY)
    pr += 1
pl.freeze_panes = "A3"
pl.column_dimensions["A"].width = 26
pl.column_dimensions["B"].width = 13
pl.column_dimensions["C"].width = 8
pl.column_dimensions["D"].width = 14
pl.column_dimensions["E"].width = 16

# =========================================================================
out = "Cricket_Auction_Tracker.xlsx"
wb.save(out)
print("saved:", out)
print("sheets:", wb.sheetnames)
