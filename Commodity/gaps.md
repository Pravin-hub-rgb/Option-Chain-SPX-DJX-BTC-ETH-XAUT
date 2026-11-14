Now I have a complete picture. Here's my **gap analysis** — everything Delta/US has that my MCX build is missing, organized by category:

---

### ✅ Already fixed (you pointed these out)
- Column header labels on row 4 (Call LTP … Put LTP × 3 blocks) — template regenerated
- Expiry dropdown on D3/O3/Z3 — `setup_expiry_dropdown()` added (hidden AH column + validation)

---

### ❌ Still missing (deep dive)

| # | Area | Delta/US has | MCX missing |
|---|------|-----------|-------------|
| 1 | **ATM/ITM/OTM coloring** | Yellow ATM row, green ITM calls, red ITM puts; repaints only when picture changes (colour_sig) | No coloring at all — all cells default |
| 2 | **Stale-row clearing / extent tracking** | `_last_rows` per block; clears rows below live count; never shrinks; `_deep_clear_below_block` | `pad_block_rows` pads but no extent tracking; leftover rows from previous expiry may linger |
| 3 | **Put intrinsic warning** | `feed_log("put_intrinsic_warning")` when put quoted below intrinsic (spot - strike) | Not present |
| 4 | **Spot/parity system** | `note_ws_recv`, `set_spot` (tracks last actual move), `spot_age`, `spot_is_stale`, full parity median with `spread_ok`/`mid`/deadband/cache | Basic parity exists but lacks spread filtering, mid-price, deadband cache, `last_spot_change`/`last_ws_recv` freshness |
| 5 | **Performance metrics recording** | `record()` called in every render phase (chain_build, excel_block_prep, excel_formatting, render_cycle) | `PerformanceMetrics` class exists but `record()` is never called |
| 6 | **Calc-mode watchdog** | `MANUAL_WATCHDOG_SECONDS`, `_calc_failed_at`, `_calc_manual_failed`, recovery + popup | `calc_mode_guard.py` exists but not wired into render loop/watchdog |
| 7 | **Excel connection tracking per symbol** | `excel_fail_since`, `excel_fail_notice_shown`, 30s cooldown, "Lost connection to Microsoft Excel" popup, `invalidate_com_cache()` | `excel_connected`/`excel_miss_counts` exist but no fail-since/popup/invalidate flow |
| 8 | **Book sheet cache** | `_book_sheet_cache`, `invalidate_com_cache(asset)` re-resolves dead COM handles | No cache; `find_open_book` re-enumerates every cycle |
| 9 | **Adaptive refresh / recalc backoff** | Measures `Application.Calculate()` cost, stretches sleep proportional to recalc time | Fixed `max(0.1, 0.5 - elapsed)` sleep, no recalc awareness |
| 10 | **Strategy formula column support** | Header comments / LTP-OI column guidance for customer XLOOKUP/VLOOKUP | Not present |
| 11 | **Startup retry** | `retry_open_pass_after_grace()` with 45s jittered backoff | Exists but not fully wired |
| 12 | **Doctor thoroughness** | Checks activation key, all packages, Excel instances, workbook readability, log files | Basic version; missing key-gate check (removed per your request) |

---

### What I'd recommend

Before I build more, tell me **which of these matter to you** (or point out anything else I missed). I can batch the important ones rather than trivia you don't care about. The coloring (1) and stale-row clearing (2) are the biggest visual gaps — want those first?