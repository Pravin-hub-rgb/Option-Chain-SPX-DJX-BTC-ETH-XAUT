# MCX Upstox probe

Lightweight probe to verify that Upstox can deliver MCX option-chain data
(bid/ask/last + open interest) before committing to a full Excel product.

## Setup

1. Copy `.env.example` to `.env`:
   
   ```powershell
   copy .env.example .env
   ```

2. Open `.env` and paste your Upstox access token on the `UPSTOX_API_TOKEN=` line.

3. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

## Run

```powershell
python probe_upstox.py
```

## Confirmed findings (with your token)

| Check | Result |
|---|---|
| Token valid | ✅ `/user/profile` → 200 (user: Pravin Kumar, MCX exchange enabled) |
| MCX instrument master | ✅ `GET https://assets.upstox.com/market-quote/instruments/exchange/MCX.json.gz` → 16,810 instruments |
| Instrument key format | `MCX_FO|<numeric_id>` e.g. `MCX_FO|579520` |
| Trading symbol | `SILVER 226000 PE 26 MAY 27` (strike + type + expiry encoded in name) |
| Quote endpoint | ✅ `GET /v2/market-quote/quotes?instrument_key=<key>` → 200 |
| Quote fields | `ohlc` (open/high/low/close), `depth` (buy/sell arrays with price/quantity/orders), `oi` (open interest) |
| Near-term expiries | CRUDEOIL 17 NOV 26, GOLD 27 NOV 26, SILVER 23 NOV 26 |
| Live prices | ⚠️ 0/empty — MCX market is closed (Saturday) |

## What this means

- Your Upstox token works and MCX data is reachable.
- Option-chain data comes via `/market-quote/quotes` (no separate `option-chain` endpoint needed for MCX).
- You discover tradable keys by downloading the MCX instrument master (gzip JSON) and filtering for `segment=MCX_FO`, `instrument_type` in (CE, PE).
- Bid/ask live in `depth.buy[].price` / `depth.sell[].price`; OI lives in the top-level `oi` field.

## Notes

- `.env` is gitignored — never commit your real token.
- `.env.example` is a placeholder only, safe to commit.
- Market data is empty on weekends; re-test on a weekday for live prices.
