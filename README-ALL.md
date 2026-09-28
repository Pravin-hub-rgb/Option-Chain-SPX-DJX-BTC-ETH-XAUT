====================================
 DELTA OPTION CHAIN — FEATURES
====================================

This tool streams live BTC, ETH and XAUT (Gold) option chain data
from Delta Exchange (India) into Microsoft Excel.

────────────────────────────────────
  1. THREE EXPIRY BLOCKS
────────────────────────────────────

Each sheet (btc_chain.xlsx / eth_chain.xlsx) has
3 blocks side by side, each showing a different
expiry date:

  Block 1  |  Gap  |  Block 2  |  Gap  |  Block 3
  (A-I)    | (J-K) |  (L-T)    | (U-V) |  (W-AE)

────────────────────────────────────
  2. CHANGE EXPIRY DATES
────────────────────────────────────

Edit these cells in Excel (they update instantly):

  D3  → Block 1 expiry date
  O3  → Block 2 expiry date
  Z3  → Block 3 expiry date

Format: DD/MM/YYYY  (or pick from the dropdown
that appears in the cell)

Examples:
  26/09/2026  → 26 Sep 2026
  02/10/2026  → 02 Oct 2026

Default: today, tomorrow, day-after-tomorrow

────────────────────────────────────
   3. CHANGE STRIKE RANGE
────────────────────────────────────

Edit cell A2 to control how many strikes to show:

  A2 empty → shows all available strikes from exchange
  A2 = 20  → shows 20 strikes (10 above ATM + 10 below ATM)
  A2 = 10  → shows 10 strikes (5 above + 5 below)

Note: if you type a number higher than what exchange provides,
tool automatically shows maximum available strikes.

────────────────────────────────────
   4. STRIKE LIMITATIONS (Important)
────────────────────────────────────

Not all expiries have the same number of strikes.

  • Today's expiry (Block 1) — most strikes
  • Tomorrow's expiry (Block 2) — fewer, starts sparse
  • Day-after expiry (Block 3) — fewest

This is how Delta Exchange works: new expiries launch
with fewer strikes and add more dynamically every ~5
minutes as the market moves.

Strike spacing also varies:
  • No fixed spacing — shows whatever the exchange lists (gaps vary)

The header row (Row 1) shows the actual count:
  "BTC Spot: 73290 | 34 strikes"

You cannot force more strikes via the API — this is
controlled by the exchange's listing rules.

────────────────────────────────────
   5. COLOR CODING
────────────────────────────────────

  Golden (row highlight)  → ATM strike (closest to spot price)
  Light green (A:D)       → ITM calls (strike below spot)
  Light red (F:I)         → ITM puts (strike above spot)

────────────────────────────────────
   6. WHAT EACH BLOCK SHOWS
────────────────────────────────────

Each block has 9 columns (same layout):

  Call LTP | Bid | Ask | OI | Strike | Put OI | Bid | Ask | LTP

  LTP   = Last traded price
  Bid   = Highest buy order
  Ask   = Lowest sell order
  OI    = Open interest (in USD value, e.g. $2.34M)

────────────────────────────────────
   7. HOW TO USE
────────────────────────────────────

  1. Open btc_chain.xlsx and eth_chain.xlsx in Excel
  2. Double-click OptionChain.exe
  3. All blocks start updating automatically
  4. Edit D3/O3/Z3 to switch expiries
  5. Edit A2 to change strike range
  6. Close Excel to stop

────────────────────────────────────
DONE
────────────────────────────────────
