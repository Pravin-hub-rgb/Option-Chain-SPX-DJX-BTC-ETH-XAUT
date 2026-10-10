====================================
 DELTA OPTION CHAIN - HOW TO USE
====================================

1. For the updated onedir release, extract the complete OptionChain folder
   from DeltaOptionChain.zip anywhere on the laptop. If Windows marked the
   downloaded ZIP as blocked, right-click the ZIP, choose Properties, select
   Unblock if shown, then extract it before running the app.
   Extract the WHOLE folder, not just OptionChain.exe. Do not extract into
   OneDrive or a Desktop that syncs to the cloud - those can leave the
   _internal folder as an empty placeholder and the app will not start.

2. Microsoft Excel (desktop) should be installed
   The app uses Excel COM; Excel for the web is not sufficient. Delta's
   xlwings path does not require the .NET Framework XLL runtime.

3. Double-click OptionChain.exe inside that folder.
   Keep the folder, including its _internal subfolder, together.
   Starting another Delta copy from a different folder replaces the older
   Delta process; it does not stop the separate US Option Chain tool.

THAT'S IT!

- btc_chain.xlsx, eth_chain.xlsx and xaut_chain.xlsx
  (BTC, ETH & Gold) open automatically
- Live data starts updating within a few seconds
- To stop: closing Excel only PAUSES the writes. The tool keeps running and
  carries on by itself the moment you reopen the workbooks. To stop it for
  good, open Task Manager, find OptionChain.exe and choose End task.
- Startup diagnostics are appended to:
  %LOCALAPPDATA%\DeltaOptionChain\startup_diag.log
  The first line of every entry is build_id, which identifies this exact
  build. Send this log with any support request and we can tell immediately
  which version you are on.
  (includes Windows/Excel details, .NET release, Python/xlwings versions,
   whether the OptionChain process is running as administrator, and Excel
   workbook connection outcomes)
  Unexpected Excel/COM exceptions and their tracebacks are written to:
  %LOCALAPPDATA%\DeltaOptionChain\app.log
  While running, aggregate refresh, chain-build, Excel-write, formatting,
  and missed-refresh timings are appended about every 30 seconds. No prices
  or strategy/leg contents are recorded. The log is rotated at 1 MB.
  The read-only diagnose_excel_addins.ps1 script writes a report on the
  Desktop. Run from PowerShell with:
  powershell -NoProfile -ExecutionPolicy Bypass -File .\diagnose_excel_addins.ps1
  It checks common Excel OPEN registry entries, XLSTART folders and XLL
  signatures/hashes; it does not change add-ins or registry settings.

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

  A2 empty  ->  shows "--" (nothing renders)
  A2 = 50   ->  50 strikes (default, 25 above + 25 below ATM)
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
 TROUBLESHOOT
------------------------------------
"Windows protected your PC" message?
  Click "More info" -> "Run anyway"

 "Failed to load Python DLL" or "_internal" missing?
   The folder is incomplete, so the app cannot start. Delete the extracted
   folder completely, extract the WHOLE ZIP again into a plain local folder
   such as C:\OptionChain (not OneDrive, not a syncing Desktop), and make
   sure OptionChain\_internal\python310.dll exists. Copying OptionChain.exe
   on its own is never enough - the _internal folder must travel with it.

 Excel add-in popup or NorenLink XLL warning?
   Close Excel, run the read-only PowerShell command above, then send the
   Desktop report and exact popup text to support. Delta does not include
   or register NorenLink and will not disable add-ins from other software.
   If the popup is your broker's add-in, dismiss it (No / OK) and reopen
   Excel once; after that Excel stays running and Delta reuses it.

 Data not updating?
   The tool waits and reconnects on its own. Close Excel, reopen the
   chain workbooks, and it refills within a few seconds. Only double-click
   OptionChain.exe again if nothing happens after that.

Excel shows an add-in or XLL warning?
  This Delta package uses xlwings/Excel COM and contains no NorenLink XLL.
  Check which add-in or startup file is registered in Excel on that PC.
  The tool does not remove or disable add-ins installed by other software.
  Run diagnose_excel_addins.ps1 and send the exact popup text, Desktop report,
  startup_diag.log and app.log
  when requesting support.

Excel becomes slow while using a strategy?
  Note which workbook is slow, what action you took (for example, changing a
  leg or expiry), and whether Excel freezes or only live quotes fall behind.
  Send those details with startup_diag.log and app.log. The log separates Delta's chain
  preparation and Excel writes/formatting; it does not include strategy data.

------------------------------------
 CONTACT
------------------------------------
For algo trading bots, screeners, scanners,
option chain data fetching & custom dashboards:

  Instagram: https://www.instagram.com/stoplossdiaries/

------------------------------------
DONE
------------------------------------
