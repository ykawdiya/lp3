# 🏏 Cricket Auction Tracker

An Excel workbook (`Cricket_Auction_Tracker.xlsx`) that tracks a 15-team player
auction. Enter each sold player once on the **Auction** sheet and every team
sheet + the balance table update automatically.

## How it works

All prices are in **crores**. The decimals are lakhs: `1.20` = 1 cr 20 lakh,
`2.00` = 2 cr. Every team starts with a purse of **100.00**.

| Sheet | What it holds |
|-------|---------------|
| **Auction** | The **full player pool** — one row per player: `Player · Price (Cr) · Team · Role · Base (Cr) · Grade`. For a player who has been sold, type the **Price** and pick the buying **Team** from the drop-down; the rest stay blank until they go under the hammer. **Role / Base / Grade** are pre-filled. Filter arrows group/sort by Grade, Role or Team. |
| **Players** | Metadata master: `Player · Role · Hand · Country · City · Base · Grade`. Edit here and the Auction + team sheets update. Country/City come only from the Grade D part of the source list, so they're filled where available. |
| **Balance** | Purse / Spent / Balance / Players for every team, plus a totals row. This sheet is the source of truth for team names. |
| **Summary** | All 15 teams on one page in the original grid layout (3×5 blocks of `S.No · Player Name · Amount` with Spent / Left / Total per team). Fully linked — updates automatically from the Auction sheet. |
| One sheet per franchise | Each team's purse, spend, remaining balance, and the auto-filled squad — `Player · Price · Role · Hand · Country · City`. |

The Auction sheet is **pre-loaded** with the full player pool (**283 players**:
280 from the IPL players list + 3 sold players not in that list). The **102
already-sold** players from the Invictus data have their Price and Team filled;
the rest are ready for you to assign. The 15 teams are: Mumbai Indians, Royal
Challangers Banglore, Chennai Super Kings, Kolkata Knight Riders, Patna
Panthers, Sunrisers Hydrabad, Delhi Capitals, Gujarat Titans, Lucknow Super
Giants, Imphal Igniter, Punjab Kings, Rising Pune Super Giants, Rajasthan
Royals, Kochi Tuskers, and Guwahati Chargers.

## Using it

1. Open the file in Excel (2016+/365), Google Sheets, or LibreOffice Calc.
2. On **Auction**, find the player, type the **Price** they sold for, and pick
   the buying **Team** from the drop-down.
3. That's it — the team's sheet lists the player and its balance drops by the
   price. Example: `Virat Kohli` sold for `27.5` to `Imphal Igniter` appears on
   the *Imphal Igniter* sheet, and that team's balance drops accordingly.
4. The top-right box on the Auction sheet shows pool size, players sold, and
   total spend.

## Notes

- Rename a team in column A of the **Balance** sheet and the name updates
  everywhere (the drop-down and that team's sheet title/filter follow it).
- Each team's purse is editable on the **Balance** sheet if you want unequal
  budgets.
- `build_auction.py` is the generator script; re-run it to rebuild the workbook.
