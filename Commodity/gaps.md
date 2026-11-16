# MCX Commodity Option Chain — Delta-parity gap audit: ALL COMPLETE ✅

## Situation

Repo restructured. Products now live at:
- `Commodity/source/` — MCX product (Upstox, CRUDEOIL/GOLD/SILVER)
- `Deltax-option-chain/source/` — Delta reference (BTC/ETH/XAUT)
- `US option chain/source/` — US reference (SPX/DJX/US stocks)

`Commodity/source/` core product was already working (verified by running it): private Excel instance, 3 workbooks open, instrument master fetch, quote fetch, ATM-centered strikes, PCR, expiry dropdown, stale-data fallback, net badge. Tests pass (19). Doctor passes.

All Delta-parity gaps below are now implemented.

## Data source (verified working)

Upstox token validated. Endpoints confirmed:
- Instrument master: `GET https://assets.upstox.com/market-quote/instruments/exchange/MCX.json.gz` → 16,634 instruments
- Quote: `GET https://api.upstox.com/v2/market-quote/quotes?instrument_key=<key>` → ohlc, depth (buy/sell bid/ask), oi
- Instrument key format: `MCX_FO|<numeric_id>` (e.g. `MCX_FO|579520`)
- Strike embedded in `trading_symbol` (e.g. `CRUDEOILM 13400 CE 16 DEC 26`) — master omits `strike`, parse from symbol
- Expiry epoch ms; nearest: CRUDEOIL 17 Nov 26, GOLD 27 Nov 26, SILVER 23 Nov 26
- Market closed Saturday → quotes 0/empty; live data flows weekdays

## Gap audit — ALL COMPLETE ✅

| # | Area | Status |
|---|------|--------|
| 1 | ATM/ITM/OTM coloring | Done — yellow ATM row, green ITM calls (198,224,180), red ITM puts (255,200,200), repaint gated by `_colour_state` signature `(n, k, m, atm_index)`, cleared before repaint |
| 2 | Stale-row clearing / extent tracking | Done — `_last_rows`, `_row_extent_seen` (never shrinks), `_deep_clear_below_block`, clears rows below live count on shrink/first sight |
| 3 | Put intrinsic warning | Done — `_put_intrinsic_rows` logs `put_intrinsic_warning` when put LTP < spot - strike, throttled 300s |
| 4 | Spot/parity system | Done — full Delta parity: `spread_ok`, `mid`, median of N strikes nearest hint, deadband cache, forward_discount/call_put_adj formula |
| 5 | Performance metrics recording | Done — `record()` in chain_build, excel_block_prep, excel_range_skipped, render_cycle, excel_recalc; 30s speed summary to feed.log |
| 6 | Calc-mode watchdog | Done — `_begin_excel_bulk_mode` (suspend calc+screen_update once per cycle, restore with 3-attempt verify, calc_mode_guard), `ensure_calculation_not_stranded()` every cycle, MANUAL_WATCHDOG_SECONDS=4s |
| 7 | Excel connection tracking popup | Done — `excel_fail_since`/`excel_fail_notice_shown`, 30s cooldown, "Lost connection to Microsoft Excel" popup |
| 8 | Book sheet cache | Done — `_book_sheet_cache`, cached book reused, invalidated on COM error |
| 9 | Adaptive refresh / recalc backoff | Done — measures `Application.Calculate()` cost in `_time_recalc`, drives `_recalc_backoff_seconds` and `_refresh_widen` in adaptive sleep with WIDEN_STEP/WIDEN_MAX/WIDEN_DECAY_SECONDS decay |
| 10 | Strategy formula column support | Done — HOW-TO-USE.txt has broken vs corrected VLOOKUP/XLOOKUP patterns + fixed row geometry guidance |
| 11 | Startup retry | Done — `retry_open_pass_after_grace()` wired into display_loop |
| 12 | _retire_unused_blocks | Done — clears blocks beyond n_blocks at connect time |
| 13 | Doctor thoroughness | Done — folder length, packages (xlwings/requests/pandas/dotenv/openpyxl), Upstox reachability, Excel, 3 workbooks, logs — all PASS |
| 14 | Build files | Done — build_exe.bat, build_final.bat, mcx_option_chain.spec, --noupx, pyarrow trim, build_info.txt stamp |
| 15 | maybe_exit_after_close_all wired | Done — `excel_alive` computed each cycle; `maybe_exit_after_close_all(excel_alive=...)` called after `retry_open_pass_after_grace()` so the tool exits when all 3 workbooks are closed |

## Validation criteria

- `python -m pytest Commodity/source/tests/ -q` — all pass (19 tests, 30 subtests)
- `python Commodity/source/doctor.py` — all PASS
- Tool launches without crash; opens 3 workbooks in private Excel instance
- Feed log shows chain building for all 3 symbols; coloring visible (yellow ATM, green/red ITM)
- Stale data shows "DATA STALE" annotation when market closed
- No `.env` with real token committed — only `.env.example`

## Open questions (resolved)

1. No activation key/license gate for MCX (removed per customer). ✅
2. MCX symbols: CRUDEOIL, GOLD, SILVER. ✅
3. Workbook names: crudeoil_chain.xlsx, gold_chain.xlsx, silver_chain.xlsx (templates exist). ✅
4. REST-only polling for MCX (no WebSocket needed). ✅
