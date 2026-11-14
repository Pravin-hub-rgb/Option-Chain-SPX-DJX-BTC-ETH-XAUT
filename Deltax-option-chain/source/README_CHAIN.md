# Delta BTC Option Chain — live Excel

Live BTC option chain in normal desktop Excel (.xlsx), driven by the same
WebSocket feed the Delta website uses.

## Files

| File | Purpose |
|---|---|
| `delta_chain_live.py` | The app. WebSocket → in-memory → batched Excel writes. |
| `latency_harness.py` | Measures freshness, writes `latency_log.csv`. |
| `DeltaChain.xlsx` | Output workbook (created on first run, sheet `Chain`). |

## Setup

```powershell
pip install requests websockets xlwings
```

Requires Microsoft Excel on Windows. `xlwings` drives Excel over COM.

## Run

```powershell
# 10 minute run, auto-picks the nearest live expiry
python delta_chain_live.py --minutes 10

# explicit expiry, custom workbook
python delta_chain_live.py --book MyChain.xlsx --expiry 08-10-2026 --minutes 10
```

## Sheet layout

Header row shows freshness at all times:

```
Index (spot_price) | Expiry | Last update | Server time | Age | Freshness |
Last index change | Secs since change | Index feed
   82617.0        | 08-10-2026 | 09:49:55 | -          | 1.86 | LIVE |
   09:49:53       | 1.86        | spot_price .DEXBTUSD
```

`Freshness` turns red `STALE` when age exceeds 3 s. `Secs since change` makes
the exchange's publish cadence visible — the number only moves when the exchange
actually publishes, which is by design, not a stall.

Data starts at row 10, one row per strike ascending, calls left / puts right:

```
Call Delta | Call Bid Qty | Call Bid | Call Mark | Call Ask | Call Ask Qty | Call OI |
Strike |
Put OI | Put Bid Qty | Put Bid | Put Mark | Put Ask | Put Ask Qty | Put Delta
```

## API gotchas — read before changing the fetch code

These are all verified against `api.india.delta.exchange` and each one fails
silently rather than loudly.

**1. `expiry_date`, not `expiry`.**

| query | HTTP | rows | distinct expiries |
|---|---|---|---|
| `expiry_date=08-10-2026` | 200 | 72 | 1 ✅ |
| `expiry=08-10-2026` | 200 | **564** | **8** ❌ |
| `expiry_date=2026-10-08` | 400 | 0 | — |

`expiry` is ignored and returns every expiry with HTTP 200. The format is
strictly `DD-MM-YYYY`; ISO gives HTTP 400.

Because `success` is still `true` when the filter is ignored, asserting
`data["success"]` does not protect you. Both files verify each returned symbol
ends with the expected `DDMMYY` suffix and discard the batch otherwise.

**2. `underlying_asset_symbols` is plural.** The singular form is ignored and
returns every asset's contracts — that is how ETH/XAUT symbols end up in BTC's
expiry buckets.

**3. The header price is the index.**

Fastest and correct: subscribe to the `spot_price` channel on `.DEXBTUSD`
(see the section below). Falls back to `spot_price` carried inside option
tickers on `v2/ticker`.

| field on `C-BTC-81200-081026` | value | what it is |
|---|---|---|
| `spot_price` | 82930.7 | underlying index ← correct, but slow (4.9 s) |
| `close` | 2051.0 | that call's own premium |
| `mark_price` | 1978.2 | that call's mark |

The trap is the *perpetual* ticker `BTCUSD`. It is `contract_type:
perpetual_futures`, and its `close` is the perp's own last trade, ~30–45 USD
from the index. Reading the header from there produces a plausible-looking but
wrong number. `spot_poller()` in the main app only uses the perp as a fallback.

Index symbols come from `GET /v2/indices`, not from ticker names:

| underlying | index symbol |
|---|---|
| BTC | `.DEXBTUSD` |
| ETH | `.DEETHUSD` |
| XAUT | `.DEXAUTUSD` ← not `.DEXXAUTUSD` |

**4. Server `timestamp` is microseconds.** Divide by 1e6. Nanoseconds gives
1970. The `spot_price` channel carries no timestamp — use receive time.

**5. Freshness must use WebSocket receive time, not per-option timestamp.**
Illiquid strikes carry stale server timestamps while the index is moving, which
reads as false staleness.

## Two sockets, on purpose

| purpose | host | channel |
|---|---|---|
| header index price | `wss://public-socket.india.delta.exchange` | `spot_price` |
| option chain rows | `wss://socket.india.delta.exchange` | `v2/ticker` |

The public host is the documented one (changelog 17.04.26 moved `spot_price`,
`mark_price`, `funding_rate` etc. there), so the header number — the thing
customers actually question — comes from the official public endpoint.

The public host does **not** carry `v2/ticker`. It was renamed `ticker` and uses
a compact array schema:

| legacy `v2/ticker` | public `ticker` |
|---|---|
| `mark_price` | `d[].m` |
| `oi_value_usd` | `d[].oi[1]` |
| `quotes.best_bid` | `d[].q[0]` |
| `quotes.ask_size` | `d[].q[1]` |
| `quotes.best_ask` | `d[].q[2]` |
| `quotes.bid_size` | `d[].q[3]` |
| `greeks.delta` | `d[].g[0]` |
| `spot_price` | `sp` (frame top level) |

Two cautions if migrating the chain rows later:

- **The `q` array order is not documented.** It was inferred by matching
  magnitudes against the legacy payload, and `ask_size` sits at index 1 while
  `bid_size` is at 3. Getting it backwards silently corrupts every row.
- **OI disagrees between hosts.** On the same contract minutes apart, legacy
  `oi_value_usd` read 123,140 while public `d[].oi[1]` read 100,351 — a 23 %
  gap. Which one matches the website is unresolved, so the legacy channel stays
  until that is settled. Legacy public channels are slated for removal
  31 Jul 2026, so this migration is owed, just not blindly.

Index frames on the public socket use compact keys `{"p","sy","ts","type"}`;
`index_listener()` accepts both that and the legacy `{"price","symbol"}`.

Verified with both sockets running concurrently for 60 s: all three index
symbols streaming (`.DEXBTUSD` 0.249 s cadence, `.DEETHUSD` 0.250 s,
`.DEXAUTUSD` 0.505 s) and both option symbols streaming on the legacy host.

## The index has a dedicated WebSocket channel

This is the important one, and it changes the latency story.

Delta has a `spot_price` channel that pushes the **index directly**, separate from
`v2/ticker`:

```
wss://socket.india.delta.exchange
{"type":"subscribe","payload":{"channels":[
  {"name":"spot_price","symbols":[".DEXBTUSD"]}]}}
→ {"price":82673.3,"symbol":".DEXBTUSD","type":"spot_price"}
```

Measured over 202 frames / 60 s:

| feed | index change cadence |
|---|---|
| `spot_price` `.DEXBTUSD` | **median 0.252 s, p95 0.560 s** |
| `v2/ticker` option `spot_price` | median ~4.9 s |
| REST snapshot | 2–11 s already stale on arrival |

**~19x more responsive.** The `v2/ticker` path only refreshes the index when an
option contract happens to print, which is why it looked like the exchange
publishes every 5 s. It doesn't — that was the sampling artefact.

So subscribe to **both**: `spot_price` for the header, `v2/ticker` for the chain
rows. The app falls back to option-ticker `spot_price` if the index channel dies.

Index symbols are **not** derivable from the ticker name — XAUT's is
`.DEXAUTUSD`, not `.DEXXAUTUSD`. Read them from `GET /v2/indices`.

### "Age" is not latency

We were measuring `now - payload.timestamp` and calling it latency. It is not.
Delta stamps `timestamp` when the price is *computed*, then batches sends. Age
therefore has a hard floor of one compute-to-send cycle, which is why it never
dropped below ~0.5 s no matter how few symbols we watched.

The meaningful metric is **time between consecutive index value changes**, which
is what the header now reports:

```
Last index change | Secs since change | Index feed
      09:49:53    |       1.86        | spot_price .DEXBTUSD
```

### Timestamp units

REST `timestamp` is **microseconds**. The `spot_price` channel carries no
timestamp at all — freshness there is receive time only.

## Measured behaviour

`latency_harness.py --minutes 3 --compare`, 2014 samples:

| mode | n | avg age | p95 | max | under 2 s |
|---|---|---|---|---|---|
| REST (`expiry_date`) | 22 | 6.72 s | 10.73 s | 10.96 s | **4.5 %** |
| WebSocket | 1872 | 1.43 s | 2.46 s | 2.59 s | **84.9 %** |

REST round trip averages 92 ms but the *payload* is already 2–11 s old when it
arrives, which is why REST polling cannot match the website. Over WebSocket,
the index now moves every **~0.25 s** via the dedicated channel.

Cross-check against a second, independent WebSocket client running concurrently,
720 paired samples within 1 s: **max absolute difference 0.0000**. No fixed
offset, so the field mapping is correct.

Excel block writes averaged **35–40 ms** (max 88 ms), comfortably inside the
500–1000 ms throttle. Writes are batched into a single
`sheet.range(...).value = rows` call and skipped entirely when the payload is
unchanged.

## Troubleshooting

**Sheet stays empty.** Read the file with `openpyxl` while Excel has it open
and you get the stale on-disk copy, not Excel's live buffer — the data looks
missing when it isn't. Check via COM:

```python
import xlwings as xw
print(xw.Book("DeltaChain.xlsx").sheets[0].range("B2").value)
```

Also allow ~60–75 s before judging: license check, three `init_asset` calls,
WebSocket connect and an `expiry_roller` probe all run first.

**Header shows a serial number instead of a time.** Excel coerced the string
into a datetime. Header cells are written with a leading `'` to force text.

**Writes vanish after start.** Do not use `xw.App().books[0]` — a freshly
created `xw.Book()` lands in the same Excel instance as a new unsaved workbook,
so `books[0]` can be that instead of your file. Use `xw.Book(path)`.