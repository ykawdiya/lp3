# 🏏 Cricket Auction Tracker

An Excel workbook (`Cricket_Auction_Tracker.xlsx`) that tracks a 15-team player
auction. Enter each sold player once on the **Auction** sheet and every team
sheet + the balance table update automatically.

## How it works

All prices are in **crores**. The decimals are lakhs: `1.20` = 1 cr 20 lakh,
`2.00` = 2 cr. Every team starts with a purse of **100.00**.

| Sheet | What it holds |
|-------|---------------|
| **Auction** | Master list — one row per sold player: `Player · Price (Cr) · Team`. The Team column has a drop-down of the 15 teams. |
| **Balance** | Purse / Spent / Balance / Players for every team, plus a totals row. This sheet is the source of truth for team names. |
| **Team 1 … Team 15** | Each team's purse, spend, remaining balance, and the auto-filled list of players it bought. |

## Using it

1. Open the file in Excel (2016+/365), Google Sheets, or LibreOffice Calc.
2. On **Auction**, type the player name, the price it sold for, and pick the
   buying team from the drop-down.
3. That's it — the team's sheet lists the player, and its balance drops by the
   price. Example: `Player_X` sold for `2.0` to `Team 4` appears on the *Team 4*
   sheet and `Team 4`'s balance becomes `98.00`.

## Notes

- Rename a team in column A of the **Balance** sheet and the name updates
  everywhere (the drop-down and that team's sheet title/filter follow it).
- Each team's purse is editable on the **Balance** sheet if you want unequal
  budgets.
- The five rows already on the Auction sheet are examples — clear them and add
  your own players.
- `build_auction.py` is the generator script; re-run it to rebuild the workbook.
