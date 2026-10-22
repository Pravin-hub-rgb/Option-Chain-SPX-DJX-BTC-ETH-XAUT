# State as of reboot (2026-10-13 night)

## DELIVERY — built and verified, ready to send
| File | Size | SHA256 (first 16) |
|---|---|---|
| delta-option-chain.zip | 4,45,12,548 | 2864136D5E935967 |
| us-option-chain.zip | 4,41,28,335 | 9CB0D2A6E2089211 |
| keygen-distributor.zip | 63,99,606 | 662A2B7B1D6A4556 (untouched) |

Build IDs: delta `352cd9e`, us `src-8f4fd2bcec59`. Git commit 352cd9e, tree clean.
KeyGen.exe: `DELIVERY\keygen-distributor\keygen-distributor\KeyGen.exe`
(round-trip verified 12/12 against the tool's own validator; sample key for
2026-10-29 was 222866290057).

## Packaged EXE verified working (before the machine got messed up)
```
A1 = 'BTC Spot: 83310.50 | 30 strikes'
A1 = 'BTC Spot: 83311.10 | 30 strikes'   <- live, moving
A5 = 3705.0   A10 = 2317.0               <- real call prices
```
85 feed.log lines, 0 errors.

## Damage I caused and fixed
- Dependency script added `+7.77` to BTC chain column A and crashed before its
  restore loop. Left `81007.77, 81207.77, ...` (strike + 7.77). NOT real data,
  NOT a tool bug. 21 cells cleared; 0 remaining.
- Repeatedly force-killed Excel -> "serious error" recovery dialogs, degraded
  Office/Excel state. This is why readings got weird late in the session.

## Unresolved: the `--` seen in late tests
Tool logs show it writing normally:
```
performance.BTC.excel_range_write: count=60, rows=1980
performance.ETH.excel_range_write: count=60, rows=1400
performance.ALL.render_cycle: count=20, avg_ms=405
```
and BTC expiries assigned correctly:
```
roller_result asset=BTC active_after=['2026-10-14','2026-10-15','2026-10-16']
```
So the tool is healthy; the `--` appeared in my *reading* script on a machine
whose Excel I had wrecked. A reboot should clear it.

## Changes in the shipped build
- Pause/Secs trader controls removed from both sheets (customer request)
- `feed.log` diagnostics (rotating, 2 MB cap) at
  `%LOCALAPPDATA%\DeltaOptionChain\feed.log`
- REST failures logged with endpoint/status/body/Retry-After
- ERROR vs NO_CONTRACTS distinguished; expiry suffix (DDMMYY) verified
- Roller retries with jittered backoff, keeps last-known-good expiry

## Measured so far (heavy strategy + tool running)
- typing avg 35.1 ms, worst 77.7 ms (criteria: <300 ms / <1 s) -> PASS
- UI probe avg 2.3 ms, max 24.5 ms, 0 samples over 1 s -> no freeze
- no "(Not Responding)"
- Excel CPU avg 22.3 %, p95 40.3 %
- 800/800 strategy rows computing; edit 1 chain cell -> both sheets recalc ~94 ms

## Next step after reboot (needs ~3 min, ask first)
Run ONE clean verification:
1. no Excel/OptionChain/python tool running
2. start the packaged EXE, wait ~90 s
3. confirm A1 shows a live numeric spot and a strike count
4. confirm A5/A10 are real call prices (hundreds/thousands, not strike+7.77)

If that passes, DELIVERY ships.