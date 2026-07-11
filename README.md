# 🏏 Cricket Auction Tracker

An Excel workbook (`Cricket_Auction_Tracker.xlsx`) that tracks a 15-team player
auction. Enter each sold player once on the **Auction** sheet and every team
sheet + the balance table update automatically.

## How it works

All prices are in **crores**. The decimals are lakhs: `1.20` = 1 cr 20 lakh,
`2.00` = 2 cr. Every team starts with a purse of **100.00**.

| Sheet | What it holds |
|-------|---------------|
| **Auction** | Master list — one row per sold player: `Player · Price (Cr) · Team · Role`. The Team column has a drop-down of the 15 teams; **Role** (Batsman / Bowler / All Rounder / Wicketkeeper) fills in automatically from the Players sheet. Filter arrows let you group/sort by Role or Team. |
| **Players** | Metadata master: `Player · Role · Hand · Country · City`. Edit here and the Auction + team sheets update. Country/City come only from the Grade D part of the source list, so they're filled where available. |
| **Balance** | Purse / Spent / Balance / Players for every team, plus a totals row. This sheet is the source of truth for team names. |
| **Summary** | All 15 teams on one page in the original grid layout (3×5 blocks of `S.No · Player Name · Amount` with Spent / Left / Total per team). Fully linked — updates automatically from the Auction sheet. |
| One sheet per franchise | Each team's purse, spend, remaining balance, and the auto-filled squad — `Player · Price · Role · Hand · Country · City`. |

The workbook is **pre-loaded** with the Invictus auction bid data — 15 teams
and 102 players. The 15 teams are: Mumbai Indians, Royal Challangers Banglore,
Chennai Super Kings, Kolkata Knight Riders, Patna Panthers, Sunrisers Hydrabad,
Delhi Capitals, Gujarat Titans, Lucknow Super Giants, Imphal Igniter, Punjab
Kings, Rising Pune Super Giants, Rajasthan Royals, Kochi Tuskers, and Guwahati
Chargers.

## Using it

1. Open the file in Excel (2016+/365), Google Sheets, or LibreOffice Calc.
2. On **Auction**, add a row: player name, price sold for, and pick the buying
   team from the drop-down.
3. That's it — the team's sheet lists the player and its balance drops by the
   price. Example: `Virat Kohli` sold for `27.5` to `Imphal Igniter` appears on
   the *Imphal Igniter* sheet, and that team's balance drops accordingly.

## Notes

- Rename a team in column A of the **Balance** sheet and the name updates
  everywhere (the drop-down and that team's sheet title/filter follow it).
- Each team's purse is editable on the **Balance** sheet if you want unequal
  budgets.
- `build_auction.py` is the generator script; re-run it to rebuild the workbook.
