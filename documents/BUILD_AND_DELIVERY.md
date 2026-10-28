# Build & Delivery Flow

## Overview

Two Python + xlwings COM tools, each shipped as a PyInstaller **onedir** executable:

| Product | Source dir | Exe | Data files |
|---|---|---|---|
| US Option Chain | `US Option Chain/source/` | `us_option_chain.exe` | `spx_option_chain.xlsx`, `djx_option_chain.xlsx`, `us_stock_option_chain.xlsx` |
| Delta Option Chain | `Delta Option Chain/source/` | `delta_option_chain.exe` | `btc_chain.xlsx`, `eth_chain.xlsx`, `xaut_chain.xlsx` |

Both read `config.json` and `optionchain.product` from their own folder. Neither ships `license.json`, `databento.key`, `.env`, or any lock/log file.

---

## 0. Pre-flight

- Working directory: repo root `deltax-option-chain/`, branch `simplify-lag-strategy` (or the feature branch you are shipping).
- Python 3.10, xlwings, PyInstaller installed.
- Run the full test suites first and confirm green:
  - `cd "US Option Chain/source" && python -m unittest discover -s . -p "test_*.py"`  (expect ~276 OK)
  - `cd "Delta Option Chain/source" && python -m unittest discover -s . -p "test_*.py"` (expect ~180 OK)
- Import gates:
  - `cd "US Option Chain/source" && python -c "import main, us_excel, us_state; print('US OK')"`
  - `cd "Delta Option Chain/source" && python -c "import main, delta_state, delta_chain, delta_spot; print('Delta OK')"`

---

## 1. US Option Chain build

```
cd "US Option Chain/source"
cmd /c build_exe.bat
```

- `build_exe.bat` runs PyInstaller with `us_option_chain.spec`, then:
  - strips unused pyarrow submodules (`--exclude-module pyarrow.flight` etc.)
  - surgically trims hook-shipped pyarrow `.pyd`/`.dll`/`.pyx` (~67 MB -> ~43 MB)
  - regenerates `us_option_chain.spec`
- Output: `dist\us_option_chain\` (onedir).
- **No zip is produced** (by design; zip-free per repo rule).
- The build stamps `client package\build_info.txt` with `src-<12 hex>` computed over every `.py/.json/.spec` in the source dir (excluding `_`, `test_`, `license.json`).

**Verify the stamp:**
```
python -c "import hashlib,pathlib; r=pathlib.Path('.'); EX=('license.json',); fs=sorted(p for p in r.iterdir() if p.is_file() and p.suffix.lower() in {'.py','.json','.spec'} and not p.name.startswith(('_','test_')) and p.name not in EX); h=hashlib.sha256(); [h.update(p.name.encode()+b'\0'+p.read_bytes()) for p in fs]; print('src-'+h.hexdigest()[:12])"
```
Compare to `client package\build_info.txt`. They must match. If not, rebuild.

**Copy to client package:**
- The build copies the onedir output into `US Option Chain\client package\`.
- Confirm `client package\` contains: `us_option_chain.exe`, `spx_option_chain.xlsx`, `djx_option_chain.xlsx`, `us_stock_option_chain.xlsx`, `config.json`, `HOW-TO-USE.txt`, `CHECK-INSTALL.bat`, `optionchain.product` (value `us`), `build_info.txt`.

---

## 2. Delta Option Chain build

```
cd "Delta Option Chain/source"
cmd /c build_final.bat
```

- `build_final.bat` runs PyInstaller with `OptionChain.spec`, then calls `package_release.ps1`.
- `package_release.ps1` validates required files, checks the exe is newer than the newest source `.py` (stale-build guard), checks the product marker is `delta`, and zips `dist\delta_option_chain\` into `dist\DeltaOptionChain.zip` + `dist\DeltaOptionChain.zip.sha256`.

**After the build (mandatory cleanup — repo rule: no delivery zips unless asked):**
```
Remove-Item -LiteralPath "dist\DeltaOptionChain.zip" -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath "dist\DeltaOptionChain.zip.sha256" -Force -ErrorAction SilentlyContinue
```

**Verify the stamp:**
- `client package\build_info.txt` (Delta ships via `source\dist\delta_option_chain\` for the combined delivery) contains `src-<12 hex>`.
- Recompute the hash the same way as US (over `.py/.json/.spec`, excluding `_`, `test_`, `license.json`) and compare.

**Confirm `dist\delta_option_chain\` contains:**
`delta_option_chain.exe`, `btc_chain.xlsx`, `eth_chain.xlsx`, `xaut_chain.xlsx`, `config.json`, `README.txt`, `HOW-TO-USE.txt`, `CHECK-INSTALL.bat`, `diagnose_excel_addins.ps1`, `optionchain.product` (value `delta`), `build_info.txt`, plus `_internal\` with the pywin32 DLLs.

---

## 3. Combined client delivery (root-level zips)

Only run this when you actually want to ship archives to a customer.

```
cd <repo root>
powershell -NoProfile -ExecutionPolicy Bypass -File package_delivery.ps1
```

- Produces `DELIVERY\us-option-chain.zip` and `DELIVERY\delta-option-chain.zip`.
- Files sit at the **zip root** (no wrapper folder) so extraction lands directly in `C:\OptionChain\` — Excel breaks past ~128 characters, so no extra folder nesting.
- The script refuses to publish if any required file is missing or if a forbidden file (`license.json`, `databento.key`, `.env`, `.env.example`, `instance.lock`, `instance_guard.log`, `doctor-report.txt`, or `OptionChain.exe`/`.pyd` for Delta) is present.
- It also verifies the `optionchain.product` marker matches the expected product.

---

## 4. Rules to remember

- **Never kill processes you did not start.** Other trading bots live on this machine. Only stop PIDs you launched, and track them.
- **Never commit secrets.** `license.json`, `*.key`, `.env` must never be committed or shipped.
- **No delivery zips unless the user asks.** The Delta build produces `DeltaOptionChain.zip` — delete it (and `.sha256`) after every build.
- **Pure module moves, no logic changes** when refactoring; tests gate every step.
- **Verify `build_info.txt`** matches the recomputed `src-` hash after every rebuild.
- **Excel COM close**: use `book.api.Close(SaveChanges=0)` then `app.quit()` — `xlwings.Book.close(save=...)` is not supported in this version.
- **`rg` is not installed** — use `Select-String` / the grep tool.
- **Tests**: run from the product `source/` dir: `python -m unittest discover -s . -p "test_*.py"`.

---

## 5. Quick checklist (copy/paste)

```
# US
cd "US Option Chain/source"
python -m unittest discover -s . -p "test_*.py"
python -c "import main, us_excel, us_state; print('US import OK')"
cmd /c build_exe.bat
# verify build_info.txt stamp, then copy to client package

# Delta
cd "Delta Option Chain/source"
python -m unittest discover -s . -p "test_*.py"
python -c "import main, delta_state, delta_chain, delta_spot; print('Delta import OK')"
cmd /c build_final.bat
Remove-Item -LiteralPath "dist\DeltaOptionChain.zip" -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath "dist\DeltaOptionChain.zip.sha256" -Force -ErrorAction SilentlyContinue
# verify build_info.txt stamp

# Combined delivery (only when shipping)
cd <repo root>
powershell -NoProfile -ExecutionPolicy Bypass -File package_delivery.ps1
```
