====================================
 DELTA OPTION CHAIN - HOW TO USE
====================================

1. Extract this folder anywhere on the laptop

2. Microsoft Excel (desktop) should be installed

3. Double-click  OptionChain.exe

THAT'S IT!

- btc_chain.xlsx, eth_chain.xlsx and xaut_chain.xlsx
  (BTC, ETH & Gold) open automatically
- Live data starts updating within a few seconds
- To stop: just close Excel (tool exits on its own)

------------------------------------
 CHANGE EXPIRY DATES
------------------------------------
Edit these cells in each sheet (updates instantly):

  D3  ->  Block 1 expiry
  O3  ->  Block 2 expiry
  Z3  ->  Block 3 expiry (BTC/ETH only -
         XAUT sheet has 2 blocks)

Format: DD/MM/YYYY  e.g.  26/09/2026
  (a dropdown list also appears in the cell)

Default: today, tomorrow, day-after-tomorrow
(New expiries are picked up automatically as the
 exchange launches them)

------------------------------------
 CHANGE NUMBER OF STRIKES
------------------------------------
Edit cell A2:

  A2 empty  ->  all available strikes
  A2 = 20   ->  20 strikes (10 above + 10 below ATM)
  A2 = 10   ->  10 strikes (5 above + 5 below)

If you type more than exchange offers, it shows maximum.

------------------------------------
 COLORS
------------------------------------
  Yellow row  ->  ATM strike
  Green       ->  ITM calls
  Red         ->  ITM puts

------------------------------------
 LICENSE (MONTHLY KEY)
------------------------------------
First run asks for a Product Key -
get it from the owner and enter it.

Tool runs till the last Friday of
the month. After that the popup
asks for the new key.

------------------------------------
 SCREEN SHARE PROTECTION
------------------------------------
If AnyDesk / TeamViewer / similar remote
access apps are running on this PC, the
tool shows an error and closes itself.

Close those apps first, then run
OptionChain.exe again.

------------------------------------
 TROUBLESHOOT
------------------------------------
"Windows protected your PC" message?
  Click "More info" -> "Run anyway"

Data not updating?
  Check internet connection, close Excel,
  and double-click OptionChain.exe again

------------------------------------
 CONTACT
------------------------------------
For algo trading bots, screeners, scanners,
option chain data fetching & custom dashboards:

  Instagram: https://www.instagram.com/stoplossdiaries/

------------------------------------
DONE
------------------------------------
