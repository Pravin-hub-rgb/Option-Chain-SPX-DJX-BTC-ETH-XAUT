# Simplification Plan — US / Delta Option Chain

Reversible working plan. Each item is small and independently testable; take them in order.

## Done (this pass)
- [x] Git checkpoint + branch `simplify-lag-strategy`.
- [x] Baseline green: US 246 tests, Delta 152 tests.
- [x] Delta OI numeric end-to-end (`fmt_oi` float, `#,##0` format).
- [x] Adaptive refresh widen cap lowered 4.0 -> 2.0 (US + Delta).

## Next (safe, preserve behaviour)
- [ ] Verify `Calculate()` is called exactly once per render cycle (no double calls).
- [ ] Verify owed-recalc state survives a busy Excel (render yield tests).
- [ ] Remove dead `_oi_fmt_done` / text-format re-arm paths already guarded by tests.
- [ ] Simplify `force_usd_block_formats` gate: keep `fmt_key` dedup, drop row-count sensitivity.
- [ ] Keep watchdogs; add soak notes before any removal.

## Do not touch without validation
- Calculation mode ownership / restore logic.
- Watchdog threads (user32-only and main loop).
- Once-per-cycle `Calculate()` and owed-recalc carry-over.
