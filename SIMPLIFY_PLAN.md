# Simplification Plan — US / Delta Option Chain

Reversible working plan. Each item is small and independently testable; take them in order.

## Done (this pass)
- [x] Git checkpoint + branch `simplify-lag-strategy`.
- [x] Baseline green: US 246 tests, Delta 152 tests.
- [x] Delta OI numeric end-to-end (`fmt_oi` float, `#,##0` format).
- [x] Adaptive refresh widen cap lowered 4.0 -> 2.0 (US + Delta).

## Next (safe, preserve behaviour)
- [x] Verify `Calculate()` is called exactly once per render cycle (no double calls).
- [x] Verify owed-recalc state survives a busy Excel (render yield tests).
- [x] Remove dead `_oi_fmt_done` / text-format re-arm paths already guarded by tests.
- [x] Simplify `force_usd_block_formats` gate: keep `fmt_key` dedup, drop row-count sensitivity.
- [x] Keep watchdogs; add soak notes before any removal.

## Soak notes
- Watchdogs retained (user32-only, main loop, calc-mode).
- `Calculate()` once per cycle guarded by `RecalcOncePerCycleTests`.
- Owed-recalc survives busy Excel (skipped/aborted cycles set `_recalc_owed = True`).
- `_oi_fmt_done` / text re-arm removed; OI numeric end to end.
- `force_usd_block_formats` keeps `fmt_key` dedup, no row-count sensitivity.
- Adaptive refresh widen cap 2.0x (US + Delta).
- Monitor feed.log: `excel_recalc_second_pass`, `render_cycle_skipped_user_input`,
  `calc_watchdog_restored`, `refresh_widened`.

## Do not touch without validation
- Calculation mode ownership / restore logic.
- Watchdog threads (user32-only and main loop).
- Once-per-cycle `Calculate()` and owed-recalc carry-over.
