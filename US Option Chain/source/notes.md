# US Option Chain — XLOOKUP Strategy Formula Fix Notes

## Problem

Client writes XLOOKUP-based strategy formulas (e.g. put calendar spreads) against
the BTC chain sheet. Three failure modes were identified:

1. **"NA" default** → `number + "NA"` = `#VALUE!` when a strike exists in one
   block but not the other.
2. **"0" default** → `number + 0` = silently wrong one-legged result.
3. **Row range asymmetry** → `P5:P33` (29 rows) vs `E5:E41` (37 rows); strikes
   in the wider range's extra rows can't be found in the narrower range.

## Root Cause in the Tool

The US tool already writes **numeric** values (thanks to `fmt_oi` returning
`float` since the prior fix), but `force_usd_block_formats` was setting OI
columns to `NumberFormat = "@"` (text format). While xlwings-written floats
stay numeric under `"@"` format, this is fragile: a cell re-entered by Excel
during a recalculation or a bulk-restore from a cached text value can flip the
underlying type to text, which then breaks `XLOOKUP` arithmetic with
`#VALUE!`.

The Delta tool has the same `force_usd_block_formats` code but its `fmt_oi`
still returns **text** (`"$1.23M"`), so Delta OI columns are genuinely text —
any `XLOOKUP` arithmetic on Delta OI would fail today.

## Changes Made (US Side)

### 1. `us_excel.py` — `force_usd_block_formats` (line 606)

**Before:** OI columns got `NumberFormat = "@"` and the docstring said
"OI = text with $ only."

**After:** OI columns get `NumberFormat = "$#,##0"` (dollar display, numeric
underlying). Docstring rewritten to reflect that `fmt_oi` returns float and
all columns are numeric.

```python
# OI: was "@"  →  now "$#,##0"
for off in (3, 5):
    col = get_column_letter(c0 + off)
    sheet.range(f"{col}5:{col}{end}").api.NumberFormat = "$#,##0"
```

Also broadened the one-off cleanup: any string containing `$`, `₹`, `Rs`,
`INR`, commas, `M`, or `K` is now caught and re-run through `fmt_oi`,
catching old `"$1.23M"` leftovers from previous builds.

### 2. `us_chain.py` — `fetch_symbol_payload` (line 295)

**Before:** `raise RuntimeError(...)` — original exception chain lost.

**After:** `raise RuntimeError(...) from e` — preserves the original cause
(e.g. `BentoHttpError` with `http_status`, `requests.HTTPError` with
`response.status_code`) so `symbol_poller` can log the real reason:

```python
raise RuntimeError("; ".join(errors) or "no data source") from e
```

Upstream, the poller at line 442-443 can now access `e.__cause__` to log the
actual HTTP status / Databento error detail in `feed_log`.

### 3. `test_us_lag.py` — `test_startup_open_sets_the_latch` (line 448)

**Before:** `inspect.getsource(us_workbook.open_excel_files)` — the
thread-spawning wrapper (`open_excel_files` just calls
`threading.Thread(target=_open_excel_files)`).

**After:** `inspect.getsource(us_workbook._open_excel_files)` — the actual
function where `_auto_open_done[role] = True` is set (lines 361, 368).

### 4. `test_us_names.py` — `test_no_nested_assignment_shadows` (line 248)

**Before:** `io.open(path, ...).read()` — file handle never closed; Python
emits `ResourceWarning: unclosed file`.

**After:** `with io.open(path, ...) as fh: src = fh.read()` — properly
closed.

## Verification

- `python -m pytest` → **168 passed, 48 subtests passed** (all green)
- `build_exe.bat` → build OK, templates verified (B2 = "± strikes — edit A2")
- `package_delivery.ps1` → `DELIVERY/us-option-chain.zip` (71.6 MB, 2,841 entries)
- Build ID: `src-4028e18df291` (built 22-10-2026 21:19:16)
- No `license.json` or `databento.key` in the ZIP (client enters their own keys)
- Product marker: `us` (correct)

## Column Reference (for client strategy formulas)

| Col | Offset | Block 1 | Block 2 | Block 3 |
|-----|--------|---------|---------|---------|
| Strike | +4 | E | P | AA |
| Call LTP | +0 | A | L | W |
| Call Bid | +1 | B | M | X |
| Call Ask | +2 | C | N | Y |
| Call OI | +3 | D | O | Z |
| Put OI | +5 | F | Q | AB |
| Put Bid | +6 | G | R | AC |
| Put Ask | +7 | H | S | AD |
| Put LTP | +8 | I | T | AE |

**Correct XLOOKUP pattern** (put calendar spread, both legs must exist):

```excel
=IF(AND(NOT(ISNA(XMATCH(E6,BTC!P5:P41,0))),NOT(ISNA(XMATCH(E6,BTC!E5:E41,0)))),
   XLOOKUP(E6,BTC!P5:P41,BTC!R5:R41) - XLOOKUP(E6,BTC!E5:E41,BTC!H5:H41),
   "strike not in both expiries")
```

Key fixes vs. the client's current formula:
- Uses `XMATCH` + `ISNA` guard to check strike existence in both blocks
- Extends both ranges to row 41 (eliminating the 29-vs-37 row asymmetry)
- Uses `-` (subtraction for net credit/debit) instead of `+`
- Returns a human-readable message instead of `#VALUE!` or a misleading `0`

## Delta To-Apply List (deferred)

1. **`delta main.py:243`** — `fmt_oi` returns text (`"$1.23M"` etc.). Must
   return `float` to match US. Then OI columns will be numeric.
2. **`delta main.py:1541`** — `force_usd_block_formats` sets OI columns to
   `NumberFormat = "@"`. Must change to `"$#,##0"`.
3. Same test fixes: `_open_excel_files` vs `open_excel_files`, unclosed file
   handles in `test_delta_names.py`.

---

# 2026-10-23 — Fixed Row Extent + Text-Format Retirement (strategy robustness)

Client follow-up: their strategy books still showed **wrong values and lag**.
Investigation (lab test `lab_test_excel_types` + feed logs + code read) found
three independent mechanisms; all are now fixed in BOTH tools.

## 1. Variable chain row extent breaks literal ranges (the "wrong values")

The client's formulas hard-code rows (`XLOOKUP(E6, BTC!P5:P33, ...) +
XLOOKUP(E6, BTC!E5:E41, ...)`), but the old write path wrote exactly
`len(chain)` rows: every expiry change moved the rows out from under their
ranges, one leg returned the `,0` if_not_found default, and the book showed
a silently wrong one-legged sum.

**Fix:** every chain block is now written to a **fixed row extent**.
- `ROW_EXTENT_DEFAULT = 128` (US) / `256` (Delta) — grown automatically the
  first time a longer chain appears, never shrunk; the growth is logged as
  `chain_extent_changed`.
- Live strikes fill from row 5; rows past the live count are written as
  blanks (`pad_block_rows`), so block geometry is constant and the literal
  ranges keep covering the data.
- `_flush_block_writes` writes the PADDED rows (no raw DataFrame, whose row
  count was the original instability). Payload signature and `_block_values`
  carry the padded rows too, so a shrunken chain is diff-cleared without any
  extra COM clear.
- A one-time deep check per block per run (`_deep_clear_below_block`) reads
  rows below the extent and clears to row 2000 **only if an older build left
  junk there** — first upgrade of an old workbook cleans up the old
  unbounded writes.

## 2. OI text-format re-arm retired (the "numeric guarantee")

Lab result: floats written via COM stay numeric under `"@"` AND `"$#,##0"`,
so `"@"` was not corrupting values on this Office build — but the write path
still re-applied `"@"` on OI columns whenever a block grew a row
(`_oi_fmt_done`), directly fighting `force_usd_block_formats`'s `"$#,##0"`.
That block is REMOVED; OI columns are numeric end to end and the only `"@"`
left is on the date cells (D3/O3/Z3, deliberate text dates). `force_usd`'s
one-off cleanup now also catches plain digit strings ("223") left as text by
old builds.

Delta OI was fixed earlier (fmt_oi float + `#,##0`), so it already pins
numeric — this change is US-only for OI.

## 3. Refresh widened adaptively (the "lag")

US default `refresh_interval_seconds` was 0.1s in `config.py` (the shipped
config.json already ships 0.5; both tools now default 0.5 when no config
file exists). On top, when the post-restore `CalculationState` probe says
Excel still owed a recalculation (`excel_recalc_second_pass`), the render
loop multiplies its sleep by 1.5 per hit (cap 4.0×) and decays back by half
after 120 quiet seconds; every stretch and decay is feed-logged
(`refresh_widened` / `refresh_widen_decayed`). A heavy strategy book gets
idle time instead of being recalc-stormed every 0.1s.

## Log hygiene

`conftest.py` added to both source dirs: pytest runs now always redirect the
product logs to a throwaway temp dir, independent of whether the
alphabetically-first `test_00_isolation.py` is in the run (test noise had
landed in the customer's real `feed.log` on 23-10-2026 during a targeted
run of a single test file).

## Verification

- US: test_us_lag 120 passed; full suite (minus real-Excel
  test_strategy_formula) 191 passed + the 2 pre-existing date-drift
  failures in test_us_expiries.
- Delta: test_delta_lag 66 passed; full suite minus real-Excel
  test_delta_strategy_formula 107 passed.
- Real-Excel strategy suites: US 6 passed, Delta 6 passed.

---

# 2026-10-23 (later) — Two expiry blocks per US sheet

Customer direction: show 2 expiries per sheet, like the Delta tool already
does. Their strategy sheet reads blocks 1+2 only, so the third block was
pure per-cycle Excel tax (writes, recolour region, recalc dirtying).

- `us_workbook.BLOCKS_PER_SHEET = 2` drives everything: slots, roller,
  dropdowns (D3/O3 now; Z3 retired), template labels, hint labels,
  `format_sheet` fills, writer loop `BLOCK_COLS[:BLOCKS_PER_SHEET]`.
- `_retire_unused_blocks` (us_excel) sweeps columns W..AE rows 1..2000
  once per session per sheet, kills the Z3 dropdown validation and blanks
  the cell — logged as `expiry_block_retired`.
- `active_expiries` / `user_pinned` are now 2-slot lists; the initial pick
  takes `dates[:BLOCKS_PER_SHEET]`. Tests updated: US test_us_lag 125
  passed; full 196 passed (+2 pre-existing date-drift failures); real-Excel
  formula suite still 6 passed.
- Delta made two-block everywhere right after: `n_blocks` now returns 2 for
  BTC/ETH too (XAUT already was 2), so the existing `bi >= nb` machinery
  clears and hides the third expiry columns (W:AE) once per session. Delta
  full suite 108 passed + real-Excel 6 passed. Its A2-empty "show
  everything" behaviour measured 20-105 rows per expiry in the live logs;
  the briefly-discussed strike cap was dropped and the code now says why.
