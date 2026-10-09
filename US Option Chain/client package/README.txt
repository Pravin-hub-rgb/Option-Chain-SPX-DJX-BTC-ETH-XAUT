====================================
 US STOCK / S&P500 OPTION CHAIN
====================================

Fetches option chain data into THREE
Excel files (auto-open, Delta-style):

  spx_option_chain.xlsx
    -> always SPX (S&P 500 index), locked
  djx_option_chain.xlsx
    -> always DJX (Dow Jones index), locked
  us_stock_option_chain.xlsx
    -> any US ticker, change in K1
    -> default AAPL

All update live at the same time.
Same symbol across files = one
shared cache (one API fetch).

Data source: BigClawd free public API
(near real-time brokerage feed, no key).
Fallback: CBOE delayed quotes (~15 min)
if BigClawd is down.

------------------------------------
 HOW TO USE
------------------------------------
 Option A — EXE (no Python needed):
  1. Extract the complete OptionChain onedir folder if it is delivered in a ZIP.
     Keep the _internal folder beside OptionChain.exe.
  2. Double-click OptionChain.exe inside that folder.
  3. THREE Excel files open + live data

 Option B — source (Python installed):
  1. Run:  python main.py
  2. All xlsx files auto-open
  3. SPX file: always SPX
     DJX file: always DJX
  4. Stock file: type ticker in K1
     (label Symbol in J1)
     AAPL -> QQQ -> MSFT -> any US stock
  5. Blocks update automatically
  6. Close Excel / Ctrl+C to stop

 Rebuild EXE later:

  build_exe.bat

  This builds an onedir client folder only. No ZIP is created unless requested.

  Keep the _internal folder beside OptionChain.exe when running.

Excel desktop is required (Excel for the web is not supported). The US tool
uses xlwings/Excel COM; .NET for Excel-DNA XLLs is not a US tool prerequisite.

Startup and performance diagnostics:
  %LOCALAPPDATA%\USOptionChain\startup_diag.log
Excel/COM tracebacks:
  %LOCALAPPDATA%\USOptionChain\app.log
Logs include workbook connection status and aggregate refresh, chain-build,
Excel-write, formatting and missed-refresh timings. They do not record prices
or strategy/leg contents and rotate at 1 MB.

Read-only Excel add-in check:
  powershell -NoProfile -ExecutionPolicy Bypass -File .\diagnose_excel_addins.ps1
Writes a report to the Desktop. It checks common Excel OPEN entries and
XLSTART files and never changes registry settings or disables add-ins.

------------------------------------
 CHANGE SYMBOL (stock file only)
------------------------------------
 In us_stock_option_chain.xlsx:
  Yellow K1 = active ticker
  Label J1: "Symbol (change):"

  Type AAPL in K1 -> 3 blocks switch
  to AAPL chain (next 3 expiries).

 In spx_option_chain.xlsx:
  K1 is LOCKED to SPX (gray).
  Label J1: "Symbol (locked):"
  Edits are overwritten back to SPX.

 In djx_option_chain.xlsx:
  K1 is LOCKED to DJX (gray).
  Label J1: "Symbol (locked):"
  Edits are overwritten back to DJX.

  Previous symbol data stays cached
  so switching back is instant.

------------------------------------
 CHANGE EXPIRY DATES
------------------------------------
   Yellow cells = editable input
   (all files)

    D3  -> Block 1 expiry  (C3: edit D3)
    O3  -> Block 2 expiry  (N3: edit O3)
    Z3  -> Block 3 expiry  (Y3: edit Z3)

  Format (text): DD/MM/YYYY
    e.g. 25/09/2026, 30/10/2026
  Dropdown also lists ALL upcoming
  expiries for that file's symbol

  Default: next 3 available expiries
  Blank cell auto-fills active expiry
  Tool rolls unpinned slots to next expiry

------------------------------------
 CHANGE NUMBER OF STRIKES
------------------------------------
  Yellow cell A2 (label B2:
  "± strikes — edit A2"):

    A2 empty -> all available strikes
                (max 100 around ATM)
    A2 = 10  -> 10 above ATM + 10 below
                ATM (21 rows incl ATM)
    A2 = 20  -> 20 above + 20 below
                (41 rows incl ATM)
    A2 > 100 -> capped to 100
                (BigClawd API max)

  Stale rows below are auto-cleared
  when you reduce A2.

------------------------------------
  COLORS / FORMAT
------------------------------------
  Yellow row  -> ATM strike (bold)
  Green       -> ITM calls (strike < spot)
  Red         -> ITM puts (strike > spot)
  Dark blue   -> column headers (bold)
  Light yellow-> input cells
                 (A2, D3, O3, Z3,
                  and K1 on stock file)

  Money columns use $ only (USD).
  OI: $1.23K / $4.56M style.

------------------------------------
 COLUMNS (each block)
------------------------------------
 Call LTP | Bid | Ask | OI | Strike |
 Put OI | Bid | Ask | LTP

  OI / bid / ask / LTP from BigClawd
  near real-time brokerage feed (no key).
  Fallback CBOE ~15 min delayed.
  OI settles EOD at OCC — vendor may
  show broker estimate intraday.

------------------------------------
  RATE LIMIT / POLLING
------------------------------------
  config.json:
    refresh_interval_seconds : Excel write (0.1 = snappy UI)
    poll_gap_seconds         : gap between source calls (2)
    round_pause_seconds      : pause after full round (2)
    expiry_refresh_seconds   : auto-roll unpinned slots (60)
    data_source              : "auto" | "bigclawd" | "cboe"
    spx_symbol               : locked file symbol ("SPX")
    djx_symbol               : locked file symbol ("DJX")
    default_symbol           : stock file default ("AAPL")

  BigClawd observed limit: ~120 req/min/IP.
  Up to 3 unique symbols (SPX+DJX+stock)
  ~ under limit with 2s gap.
  Same ticker across files = 1 fetch.

------------------------------------
  PRODUCT KEY / LICENSE
------------------------------------
  First run: popup asks for a
  product key (not pre-activated).
  Tool runs till the last Friday of
  the month.
  After that the popup asks for the
  new key.

  Wrong key -> "try again" popup
  and the dialog reopens. Cancel
  -> app closes.

  Get the current key from owner.

  While running: re-check ~every 60s.

====================================
  NOTES
====================================
  - Primary BigClawd = near real-time
    free (brokerage feed, no key)
  - Fallback CBOE ~15 min delayed
  - Official OCC OI settles EOD; live
    number may be broker estimate
  - Free feeds: personal tool use —
    check ToS before commercial resell
  - True OPRA real-time: IBKR / paid

====================================
