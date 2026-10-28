# DELIVERY CHECKLIST

Both packages are built and verified from the delivered artefacts. Everything
below has been run against the actual `.exe` and `.zip`, not against the
source tree.

---

## 1. The two root causes behind "it doesn't work"

Both were found today and both are fixed. Neither is a guess - each was
reproduced, measured and re-measured after the fix.

### 1.1 The Delta ZIP was a month old

`package_release.ps1` zipped `dist\OptionChain` - a folder that nothing
writes to. The build produces `dist\delta_option_chain`. The folder that was
being shipped was dated **14 October**, eight days before the work started.

**Every package you have ever sent the customer was old code.** Reinstalling
it could not possibly have helped, because the new code was never in it.

Fixed: the release script now zips the folder the build actually produces,
and it **fails the build** if the executable is older than the source:

```
The built executable is older than the source (...). Run build_final.bat
again before packaging.
```

The stale `OptionChain.exe` (the old wrongly-named build) is also removed
from the client folder, so there is no longer a second executable in the
package for the customer to run by mistake.

### 1.2 Excel could not open the workbooks

Two separate faults, both fatal:

**Excel stops opening a workbook somewhere past 128 characters and never
returns.** A package unzipped under a user's Downloads or Desktop folder is
over that easily. The symptom was the worst kind: the workbook never
opened, the tool said *"Excel is not running"* while Excel was sitting right
there, and the sheets stayed empty. That is exactly what "it doesn't work"
looks like.

Measured on this machine:

| Path | Length | Result |
|---|---|---|
| `C:\OptionChainTest\spx_option_chain.xlsx` | 40 chars | opened in 0.2 s |
| `...\deltax-option-chain\US Option Chain\source\spx_option_chain.xlsx` | 132 chars | **never returned** |
| the same file by its 8.3 short name | 80 chars | opened in 0.2 s |

Fixed: the workbooks are now opened through the Excel API using the 8.3
short name when the path is long, so it works at any path.

**Going through the Windows shell started a separate Excel for every
workbook.** Three workbooks meant three Excel processes, two of them quit,
and the one that survived was never registered where the tool could reach
it. Same symptom.

Fixed: one Excel, opened through the API the tool writes with. The empty
`Book1` Excel always starts with is closed again, and only if this process
started Excel and only if the book is empty.

---

## 2. What is in the packages now

Send **only the ZIP**. There is no `.sha256` sidecar and none is needed - the
build id is stamped inside the package (`build_info.txt`), so support can
still tell which build a customer is running from `startup_diag.log`.

Rebuild them any time with:

```
powershell -NoProfile -ExecutionPolicy Bypass -File package_delivery.ps1
```

It refuses to publish anything that is missing a file the product needs, or
that carries the owner's activation file or the Databento key.

### US Option Chain - `DELIVERY\us-option-chain.zip` (71.6 MB)

| File | |
|---|---|
| `us_option_chain.exe` | 13.4 MB |
| `CHECK-INSTALL.bat` | **new** - the customer's self-test |
| `HOW-TO-USE.txt` | **updated** - self-test and short-folder instructions |
| `spx_option_chain.xlsx` / `djx_option_chain.xlsx` / `us_stock_option_chain.xlsx` | |
| `config.json`, `optionchain.product`, `build_info.txt` | |

No `license.json` (the customer's own file), no `instance.lock`, no
`instance_guard.log`, no Databento key.

### Delta Option Chain - `DELIVERY\delta-option-chain.zip` (46.4 MB)

Was 76 MB - `pyarrow` was 78 MB of dead weight the tool never uses; removed
and the build re-verified.

| File | |
|---|---|
| `delta_option_chain.exe` | 13.1 MB |
| `CHECK-INSTALL.bat` | **new** |
| `HOW-TO-USE.txt` | **new** - Delta had no setup instructions at all |
| `btc_chain.xlsx` / `eth_chain.xlsx` / `xaut_chain.xlsx` | |
| `config.json`, `optionchain.product`, `build_info.txt`, `README.txt`, `instruction.txt`, `diagnose_excel_addins.ps1` | |

---

## 3. Verify on the customer's machine - 60 seconds

**Send them this first, before the tool:**

1. Unzip to a **short folder**, e.g. `C:\OptionChain`. Not in Downloads, not
   on the Desktop, not in a folder with a long name.
2. Double click **CHECK-INSTALL.bat**.
3. It writes `doctor-report.txt` next to the EXE and opens it.

The report checks, in plain English:

| Check | What a failure means |
|---|---|
| folder path is short enough for Excel | Move the package to `C:\OptionChain` |
| python, xlwings, requests, pandas, websockets, databento, openpyxl | Something is not installed |
| activation key saved | Expected to be missing until they enter the key |
| Databento key saved / **accepted** | The key is missing, or Databento rejected it |
| Excel is running | The tool starts it, but if it keeps happening, start Excel by hand once |
| the three workbooks | Present and readable |
| the log files | Where support will look |

**If that report is all PASS, the tool will work.** If it is not, the report
tells you exactly what to fix. No more guessing, no more reinstalling.

---

## 4. What was verified, on the delivered files

| | US | Delta |
|---|---|---|
| Self-test from the delivered EXE | PASS | PASS |
| Tool opens all three workbooks by itself | yes | yes |
| Sheets populate with live data | yes | yes |
| Spot | SPX 7757.17, DJX 528.56, AAPL 343.99 | BTC 81853.10, ETH 2406.77, XAUT 4075.73 |
| PCR | 0.92 / 1.14 / 0.99 | 0.71 / 0.73 / 0.62 |
| Expiry cell written | 20/11/2026, 23/10/2026, 23/10/2026 | today's expiry present |
| Live feed | Databento connected, 5382 symbols subscribed, 353 quotes streaming | live |
| Status line | **LIVE** | live quotes |

Both were run from a short folder exactly as the customer will.

### The expired Databento key

Tested on the delivered EXE with a deliberately broken key in
`%LOCALAPPDATA%\USOptionChain\databento.key`:

```
databento_key_rejected masked=db-T...0000 detail=Databento rejected this
key. Paste a valid live key.
```

The dead key was **deleted from disk** and the tool opened the
**"Databento API Key"** dialog, waiting for a new one. That is the behaviour
you asked about, and it is on the US product.

**But it only deleted the key because Databento actually refused it.** A
connection failure used to be treated as a refusal, so a customer whose
internet dropped for a moment had their working key thrown away and had to
dig out a new one. Now a key is only ever deleted when Databento answers
401 or 403; if Databento cannot be reached the key is kept and the app says
so.

### US and Delta running together

Tested with both delivered builds running at the same time in one Excel:

- One Excel instance, six workbooks open - three US, three Delta.
- Both processes stayed alive. Neither killed the other: the guards match on
  the executable name (`us_option_chain.exe` vs `delta_option_chain.exe`) and
  then on the folder, and each lock file lives in its own product folder.
- Each tool writes only its own three workbooks.
- The calculation mode is shared, because it is one Excel. The tool that owns
  it holds it and the other steps aside - the Delta log showed
  `calc_watchdog_skip_holder note=another chain tool holds manual
  calculation; watchdog left it alone`, and the US tool was streaming 353 live
  quotes at the same time.

**There is no conflict between the two products.**

---

## 5. Also fixed today (both products)

**US**
- A rejected Databento key no longer retries forever. It is recognised as
  permanent, the sheet header says **DATABENTO KEY REJECTED**, and the
  customer gets one message box instead of a blank sheet and a log full of
  401s.
- Three bugs where a name was read before it was assigned - an
  `UnboundLocalError` waiting to happen. One of them meant the first expiry
  roll always died before the workbook's date cell was updated.
- The "slow down when Excel will not go manual" path was dead code behind a
  flag nothing ever set, and would have crashed if it had run. Both fixed.

**Delta**
- The expiry was dropped four hours before it actually expired (BTC and ETH
  settle at 12:00 UTC; the code dropped it at 08:00 UTC). The settlement
  time is now read off the products instead of being a hard-coded hour.
- A failed probe made the expiry ladder flip-flop under the trader. It now
  keeps what is on screen and retries.
- The feed heartbeat was only wired to the expiry roller, so a healthy feed
  said `LOADING: (last good never)` forever.
- Two writers could run at once. Now the second one refuses to start.
- Split into `delta_state`, `delta_chain`, `delta_spot`, `delta_excel`,
  `delta_instance`, `delta_licensing`, `delta_diagnostics`,
  `delta_performance`. `main.py` went from 3334 to 2900 lines.
- **Faster than US**: render cycle 152 ms against 460-600 ms, cycle period
  0.69 s against 0.97-1.19 s, zero refresh overruns, and the 340-712 ms
  formatting spike is gone.

**Both**
- The test suite no longer writes into the customer's `feed.log` - a full
  test run adds zero lines. That file is what you ask the customer to send.
- New static checks that catch the bug classes above.

---

## 6. Tests

```
cd "US Option Chain"    && python -m unittest discover -s source/tests -p "test_*.py" -t source   -> 169 tests, OK
cd "Delta Option Chain" && python -m unittest discover -s source/tests -p "test_*.py" -t source   ->  74 tests, OK
```

---

## 7. When you install on the customer's machine

1. **Delete the old package first** - including the old `OptionChain.exe`
   if it is there. Extract the new one into a **short** folder.
2. If the old install is in a long path, that alone may have been the
   problem.
3. Run `CHECK-INSTALL.bat` before starting the tool. Keep the report.
4. The activation key goes in `license.json` next to the EXE, on first run.
5. The Databento key goes in `%LOCALAPPDATA%\USOptionChain\databento.key`.
   A frozen build ignores `.env` entirely - the key must be in that file or
   entered in the tool's prompt.

---

## 8. Still open (not needed for delivery)

- Delta's chain, spot and Excel logic is still in `main.py` (~2200 lines).
  The split is half done; the modules that exist are done properly.
- Delta Exchange's own put quotes look wrong: the put curve rises with
  strike on every expiry, and a 79200 put is quoted at 0.4 with delta
  -0.002 while BTC is at 81815 - it must be worth at least 2595. Their own
  product metadata and their pricing engine disagree with each other. We
  reproduce their numbers faithfully, so the customer sees the exchange's
  prices, not ours. A sanity check that flags a put priced below its
  intrinsic value would turn a silently wrong number into a visible warning.
- Neither product saves the workbook. The chain lives in Excel's memory, so
  closing Excel without saving loses it and the file on disk still shows the
  template's sample strikes. Deciding whether the tool should save is a
  product decision.
- The Delta put-quote issue above is the one I would look at first if the
  customer says the prices look wrong.
