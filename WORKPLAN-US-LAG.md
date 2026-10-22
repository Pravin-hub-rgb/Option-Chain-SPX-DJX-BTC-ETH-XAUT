# US Option Chain — Lag Fix & Refactor Workplan

Owner: opencode session | Started: 2026-10-22
Repo: remote `origin` = `git@github.com:Pravin-hub-rgb/Option-Chain-SPX-DJX-BTC-ETH-XAUT.git` (branch `main`)

## Standing rules

- Never kill or stop any process we did not start. Other trading bots are live on this machine.
  Our own launched test processes are tracked by exact PID and stopped only by that PID after verifying it is ours.
- No Chinese or garbled text in code, commits, docs, or chat.
- Every module move is a pure move: no logic changes mixed in. Tests are the gate for every step.
- Never commit secrets: license.json, *.key, databento.key, .env*, kg.dat, dxc_kg.dat, KeyGen folders stay ignored.

## Diagnosis summary (evidence-based, 2026-10-22)

| # | Cause | Evidence |
|---|-------|----------|
| 1 | Calc-mode flapping every ~4.5s: `ensure_calculation_not_stranded` (4s watchdog, never checks ownership) vs "hold manual for session" design; next cycle drops ownership | main.py:3365-3399, 3477-3478, called 2x/cycle at 3239+3620 |
| 2 | ~45 fixed Excel COM round-trips per 0.5s cycle even with nothing to write | main.py:2728-2745 excel_info per cycle/role, single-cell reads, 2x watchdog probes, CalculationState probe, tasklist 1s. Measured idle: render_cycle 190ms |
| 3 | Both-tools collision: Delta restores automatic after each window while US grabs/holds/flaps manual | Delta main.py:2172-2191 vs US session-hold; shared calc_mode_guard.py |
| 4 | Live-market write amplification: spot header rewritten every cycle, full-block writes when >8 scattered runs, Calculate() app-wide every cycle | main.py:1128, 2907-2925, 3093-3122, 3167 |
| 5 | Feed stalls: 30s semaphore-starving universe load, reconnect resubscribe under _sub_lock with 0.12s chunks, no-op subscribe logging (~695 lines/run), unbounded _quotes copies | databento_feed.py:726-731, 1204-1214, 1116-1121, 1373-1374 |
| 6 | Monolith risk: US main.py 4,167 LOC (bad-edit history: 2 misfires tonight by other bot) | repo structure |

Baseline numbers (idle machine, this morning): render_cycle avg 191ms / max 311ms; excel_pass 44-50ms per role; recalc skipped (nothing written). Target after fixes: <150ms idle cycle, zero flap.

## Track A — Git safety net + modular split (pure moves, test-gated)

| # | Task | Status |
|---|------|--------|
| 0 | .gitignore extended (DELIVERY/, client package/, RUN-THIS/, GIT-BACKUP/, *.exe, *.xlsx, *.bak, .pytest_cache/) | DONE |
| 0 | git init -b main, remote origin attached, SSH auth tested OK (Pravin-hub-rgb) | DONE |
| 0 | git fetch, reconcile with remote history, baseline commit, push | DONE (baseline 0e17db3, history preserved from 4703491) |
| A1 | Move shared globals into `us_state.py`; import check + 117 tests; commit | DONE (4574b49) |
| A2 | Extract `us_licensing.py` (~500 LOC key/licence/dialogs); verify; commit | DONE (b8a9072) |
| A3 | Extract `us_workbook.py` (~220 LOC template/ensure/open); verify; commit | DONE (5b7c651, includes symbol text helpers) |
| A4 | Extract `us_spot.py` (~500 LOC Yahoo/Tencent/parity/cache); verify; commit | DONE (3b91528) |
| A5 | Extract `us_chain.py` (~1,100 LOC symbol state, expiry, build, pollers); verify; commit | DONE (149e75e) |
| A6 | Extract `us_excel.py` (~1,700 LOC all COM writes, bulk mode, render cycle); verify; commit | DONE (df519d7) |
| A7 | Extract `us_instance.py` (lock/pid/guard + calc-mode lifecycle); verify; commit | DONE (96a69ae) |
| A8 | Slim main.py to wiring (main_async, startup, feed start/stop); print_chain to us_chain, flush to us_performance; full suite; commit | DONE (3be49ed) |

Gate after every step: `python -m unittest` all suites in `US Option Chain/source` + `python -c "import main"` passes.

## Track B — Delta (crypto open tonight) — validate shared protocol, do NOT ship client build

| # | Task | Status |
|---|------|--------|
| B1 | Delta: calc-mode restore respects other holder (skip restore automatic if another guard holder alive) | DONE (commit 4b0f492, 9 new tests green) |
| B2 | Log `excel_calc_mode_churn` counter in Delta | DONE as `calc_switch_summary` observer (commit 4b0f492) |
| B3 | Live local test on tonight's crypto session with measure_strategy/latency_harness; keep client build untouched | DEFERRED at user's discretion (trading done); coexistence covered by the 9 unit tests |
| B4 | Delta parity: port C2+C3 into the Delta monolith (merged 4x31 strip read, connection-change env detail, 2s tasklist cache, 30s calc probe gate, row-run merged span-write); 33 tests green | DONE (3dda395) |

## Track C — US lag fixes (land after split; each fix = one module + one test)

| # | Task | Status |
|---|------|--------|
| C1 | Calc-mode flapping: skip watchdog while `_calc_session.owned`; keep ownership in already-manual branch | DONE (d1807c0) |
| C2 | COM-tax: excel_info computed only on connection-change; 1 merged block-read replaces ~9 single-cell reads/role; 1 watchdog call/cycle; gate CalculationState probe to 30s; tasklist cache 2s | DONE (8400c8f) |
| C3 | Unprotected fallback: merge scattered runs into one span-write (unchanged rows carried from prev values) | DONE (bb6dbbe) |
| C4 | Feed: pollers use wait=False after initial sync; universe timeout 30s->5s; reconnect resubscribe on worker thread (no _sub_lock hold); subscribe log only added>0; prune _quotes to target set; single lock section in _on_record | DONE (3f8f202) |
| C5 | Both-tools coexistence: port B1 protocol into US; no fight over calc mode | DONE (58bd615) |
| C6 | Observability: `calc_mode_flap` counter + 30s summary log line | DONE (4d791a2) |

## Track D — Build, verify, ship (before US open ~19:00 IST)

| # | Task | Status |
|---|------|--------|
| D1 | Full unittest suite green (target >117, incl. new regression tests) | DONE (a027d5c) |
| D2 | Idle-Excel run: render_cycle <150ms, zero flapping, compare vs baseline | DONE (1c3a9c7): package build src-307f2e859ee4, steady speed lines cycle_ms 92-104, flap_total stays 1 after startup, Databento universes + OI fresh, SSE quotes waiting for the open |
| D3 | Rebuild EXE via build_exe.bat; verify build_info stamp | DONE (26a95f6) |
| D4 | Closed-market smoke: launch our PID, verify A1/strikes/expiries, stop by PID only | DONE (1c3a9c7): PID-tracked launch showed "SPX Spot: 7757.17 | 201 strikes", "DJX Spot: 528.56 | 27", "AAPL Spot: 343.99 | 77" with night-stale headline text; A-split NameError breakage (spot_source_for, refresh_tencent_spot_cache, SPOT_POLL_SECONDS, display_loop, math) found and fixed live; static name-binding test locks it (144 tests); our PID + books cleaned after the run |
| D5 | Refresh DELIVERY zip + WORKPLAN/DIAGNOSIS handoff note | PENDING |
| D6 | Client instructions: send %LOCALAPPDATA%\USOptionChain\feed.log speed lines + title-bar recalc after 19:00 open | PENDING |
