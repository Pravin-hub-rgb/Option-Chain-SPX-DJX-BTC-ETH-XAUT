# US Stock Option Chain — Data Source Decision (2026-09-23)

## Primary: BigClawd (free, no key, near real-time)

- Chain: `GET https://bigclawd.com/api/v1/trading/options/chain/{SYMBOL}`
  - Query: `strike_count` (N around ATM, default 50), `expiration` (YYYY-MM-DD, repeatable)
- Spot quote: `GET .../trading/quote/{SYMBOL}`
- Market status: `GET .../trading/market/status`
- Fields: bid, ask, mark, last, volume, open_interest, IV, greeks (delta/gamma/theta/vega/rho)
- Auth: none for public reads. Rate limit **60 req/min/IP**. Server chain cache TTL ~10s.
- Docs: brokerage market-data feed, claims **real-time** quote/chain (not CBOE 15-min delayed).
- Latency: ~2.5–9.8s per pull (full SPY all-expiries ~768KB–6k rows). Prefer `strike_count` / `expiration` filters when possible.

## Fallback: CBOE delayed quotes (~15 min)

- `GET https://cdn.cboe.com/api/global/delayed_quotes/options/{SYMBOL}.json`
- Full chain + greeks, one request; used only if BigClawd fails (`data_source: auto`).
- **Index symbols need underscore prefix:** `SPX→_SPX`, `DJX→_DJX`, `NDX→_NDX`, `VIX→_VIX` (bare name → 403).
- Index payloads are large (`_SPX` ~13MB) — timeout 60s for indices.

## Index vs ETF (verified 2026-09-23)

| Ticker | Type | Source that works |
|--------|------|-------------------|
| SPY | S&P500 ETF | BigClawd (+ premium LTP/bid/ask/OI) |
| SPX | S&P500 index option | BigClawd (`strike_count=50`; 100 → 503) |
| DIA | Dow ETF | BigClawd |
| DJX | Dow index option | **CBOE only** (`_DJX`; BigClawd 503) |
| NDX, VIX, RUT | index | BigClawd + CBOE `_` prefix |

Helper in `main.py`: `INDEX_SYMBOLS`, `cboe_symbol()`, `_BIGCLAWD_NO_INDEX` (CBOE-first order).

## Rejected for free near-real-time trading

| Source | Why |
|--------|-----|
| CBOE alone | ~15 min delay |
| Yahoo / yfinance | Often bid/ask/OI = 0; ToS personal-use |
| Nasdaq API | Works with assetclass+date window but pre-market static / fragile |
| marketdata.app free | Delayed / 100 credits/day; full chain burns credits |
| OptionWatch | Real-time-ish bid/ask snapshot but **no OI** on public REST |
| Robinhood / Webull / MarketWatch | 403 / browser session |
| Finnhub / Polygon / Tradier / IBKR OPRA | Paid key or account for real-time |

## Config

`config.json` → `"data_source": "auto" | "bigclawd" | "cboe"`

`auto` order: BigClawd → CBOE. Sticky per-symbol after first success.

---

# Earlier research notes (raw)

Seedha jawab: koi bhi source completely free + legally resell-karne-layak options data nahi deta — sabme koi na koi catch hai. Yahan breakdown hai:
Yahoo Finance (yfinance jaisi libraries)
Broker account nahi chahiye, India se bhi bilkul chal jayega
Catch: Yahoo ka terms of use isko sirf "personal use" ke liye allow karta hai. Isse ek paid subscription product bana ke bechna technically unke ToS ke against hai. Plus yeh unofficial scraping hai (official API nahi), toh kabhi kabhi break/rate-limit ho jaata hai.
MooMoo
Unka OpenAPI free hai aur quotes ke liye funded trading account bhi zaroori nahi — bas ek moomoo user ID chahiye
Lekin moomoo India me officially available hi nahi hai. 2026 tak sirf US, Canada, Singapore, Australia, Japan, Malaysia aur New Zealand me signup hota hai. India se account banana without some workaround (jo unke terms ke against hoga) possible nahi hai.
Barchart
Inka real production-level OnDemand API $500/month se start hota hai, aur real-time options data (OPRA feed) ke liye alag se exchange fees bhi lagte hai. "Free" wala hissa sirf trial/limited hai — ek paid product ke liye kaam nahi karega.
Better fit for you: Alpha Vantage, Polygon.io, Finnhub, EOD Historical Data jaise vendors — inme koi broker account ya geo-restriction nahi hota, sirf card se signup karke API key milti hai, India se bhi easily kaam karega. Free tier limited/delayed hota hai, lekin paid tiers relatively saste hai (~$20-50/month range) aur woh commercial use allow karte hai (Yahoo ke unlike).
Agar yeh NSE/Indian options ke liye hai, toh NSE khud apni website pe free option chain data deta hai bina account ke — bohot log ise scrape karke apne tools me use karte hai (though NSE ka ToS bhi bulk automated pulling officially allow nahi karta, informally common hai).
Sach yeh hai: pura free + fully-legit-for-resale combo milna mushkil hai. Aapko ya toh ek paid data vendor (jo commercial use allow kare) choose karna hoga, ya Yahoo/NSE jaisa free-but-gray-area route lena hoga aur risk accept karna hoga.Do cheezein clear karni hain yahan:

**Yahoo har sec update nahi deta**
Yahoo ka apna official help page saaf likhta hai ki unka data redistribute nahi kiya ja sakta — sirf viewing ke liye hai, kisi product/tool me resell karne ke liye nahi. Aur technically yeh ek streaming feed bhi nahi hai — yfinance jaisi libraries baar baar Yahoo ke server ko "poll" karti hain (request bhej ke). Agar aap isko har second hit karoge, Yahoo aapko rate-limit/block kar dega. So dono wajah se — legal aur technical — Yahoo se ek reliable "har sec update" product nahi ban sakta.

**OI ka ek chhota sa nuance**
OI (Open Interest) technically exchange (NSE ho ya US ka OCC) sirf **din me ek baar, end-of-day settle** karta hai — yeh bid/ask/LTP ki tarah tick-by-tick nahi badalta. Jo "live OI" Upstox jaisे brokers dikhate hain woh actually ek **estimate** hota hai (order flow se calculate hota rehta hai), official settled number nahi. Toh "har sec OI update" ka matlab hoga estimated live OI, jo bilkul theek hai — bas samajh lo yeh approximation hai, exchange ka final number nahi.

**Upstox wala model kaam kyun kiya tha**
Upstox ek **licensed Indian broker** hai, isliye woh apne khud ke account holders ko free real-time NSE data API se de sakta hai — legally, kyunki NSE se unka official data agreement hai. Yeh Yahoo scraping se bilkul alag cheez hai.

US stocks ke liye iska closest equivalent **Interactive Brokers (IBKR)** hai — India se account khulta hai, aur unka TWS/Client Portal API account holders ko real-time data deta hai. Bas options (OPRA feed) ke liye ek chhota sa monthly data fee lagta hai (~$1.50-10/month range), jo aksar trading commission se waive ho jata hai. Yeh "poore free" nahi hai but legit, per-second-level update dega, aur India se bina kisi restriction ke sign up ho jayega — jo moomoo ke saath problem thi.

So this was from claude you do the research here find some thing reddit etc etc kuch whole internet us stocks aur snp500 ki option chain wagera kaise mile
Worked for 31s

Seedha jawab: Pure free + legally resell-karne-layak real-time US stocks/S&P 500 (SPY/SPX etc.) options chain data nahi milta. OPRA licensing + exchange rules ke wajah se real-time consolidated data free commercial redistribution allow nahi karta. Free sources personal/research use ke liye hain, commercial product (resell) ke liye catch hai. Yahan 2026-updated research (vendor docs, Reddit, comparisons) se practical breakdown:
Free / Near-Free Options (Research / Limited Use)
Source	Data Type	Freshness	OI / Greeks	Limits / Catch	Commercial Resale?	India Access
Yahoo Finance / yfinance	Full chains (calls/puts, volume, OI, some Greeks)	Delayed / poll-based	Yes (OI + basic)	Rate-limit / break risk; ToS = personal use only	No (explicitly against)	Yes, no account needed
CBOE delayed pages / public endpoints	Chains for SPX, SPY, equities, VIX	Delayed	Yes (OI, volume, Greeks)	Automated scraping explicitly prohibited	No	Yes
DoltHub (post-no-preference/options)	Historical EOD chains, 2000+ US symbols ~2019–present	EOD	Yes (Greeks/IV)	Bulk needs dolt CLI clone; public queries limited	Gray (community mirror)	Yes
Massive (ex-Polygon) Options Basic	Contracts reference + EOD	EOD	Limited on free	5 calls/min, 2 yrs history	Individual only; redistribution needs Business plan	Yes
Kaggle / OptionsDX slices	Historical SPY etc.	EOD / some intraday	Yes	One ticker or limited years; stops ~2023 on free	Personal	Yes
Public.com API	Chains + quotes	Real-time (account holders)	Yes	Needs Public account; trading-focused	Account-tied, check ToS	Check residency
Tradier API	Full chains + Greeks (ORATS)	Real-time (live account) / delayed sandbox	Yes	Broker account needed; funding often required after trial	Account holders only	US-centric, India eligibility check karo

Reddit consensus (r/options, r/algotrading, r/Spyoptionsdaytraders etc.): Free real-time OPRA-level chain almost impossible. Log historical ke liye DoltHub recommend karte hain. Live ke liye log ya to broker API (Tradier/Public/IBKR) use karte hain ya paid data vendor. Koi “completely free + reliable + commercial” solution nahi dikha.
Paid but Reasonable (Commercial-Friendly Path)

    Massive (Polygon.io rebrand) Options:
        Starter: ~$29/mo → 15-min delayed, unlimited calls, Greeks/IV/OI, WebSocket.
        Developer: ~$79/mo → more history + trades.
        Advanced: ~$199/mo → real-time.
        Business: ~$1,999/mo → redistribution-friendly, no extra exchange approvals on core.
        Free Basic se start kar sakte ho testing ke liye. India se bilkul chalega (card se signup).
    IBKR (Interactive Brokers) – Closest “Upstox-style” legit model:
        India residents IB India account khol sakte hain (min equity data ke liye ~$100 lower).
        OPRA Top-of-Book (US options L1): ~$1.50/mo (often waived with ~$20 commissions).
        US Securities Snapshot & Futures Bundle: ~$10/mo (waived with ~$30 commissions).
        Client Portal / TWS API se real-time options chains + streaming milta hai account holders ko.
        Yeh licensed broker data hai → legally clean for your own use/tools. Resell derived products ke liye alag check, lekin Yahoo scraping se kaafi better.
    Other solid mid-tier:
        Alpha Vantage: Options on higher paid tiers (~$50–200/mo range). Commercial ke liye unse contact.
        EODHD: US options EOD/historical ~$30/mo range.
        Tradier: Account + API free, real-time chains. Commission plans low ($0–$0.35/contract). India se account open hota hai ya nahi, verify karo.
        Databento / Intrinio / FlashAlpha: Higher-end (analytics + real-time), $100+ se start.

Important Technical Nuances (jo pehle bhi discuss hua)

    OI: Exchange (OCC) official OI sirf end-of-day settled deta hai. Brokers ka “live OI” estimate hota hai (order flow se). Har-sec exact official OI impossible.
    Real-time vs Delayed: True OPRA consolidated real-time ke liye exchange fees + licensing lagti hai. Free tiers almost always EOD ya 15-min delayed.
    Yahoo limitation confirm: Official help + ToS redistribution allow nahi karta. Polling se har-sec reliable product nahi banta (rate-limit/block).

Recommendation for Your Use-Case (Product bana ke bechna, India se)

    Sabse clean low-cost: IBKR account + market data subs (bahut sasta ya free with commissions) → real-time API access.
    Pure data vendor (no broker needed): Massive Options Starter/Developer se start → delayed/commercial testing. Real-time + resale ke liye Business tier.
    Prototype / gray area: Yahoo + CBOE delayed + DoltHub historical. Risk accept karo (ToS break + reliability).
    Avoid: Moomoo (India official nahi), pure free scraping for production product.

Agar specific ticker (sirf SPY/SPX) ya EOD vs real-time priority batao, toh aur narrow recommendations de sakta hoon. Commercial redistribution ke exact license language ke liye vendor se seedha confirm karna best hai — ToS change ho sakte hain.
100 sources

What is the safest rate limit though research on it 5 sec  gap polling?
Worked for 16s

Safest practical rate limit for Yahoo / yfinance polling (especially options chains):

Yahoo publishes no official rate limit. Everything is undocumented, per-IP, and has tightened since 2024–2025. Hitting limits returns YFRateLimitError / HTTP 429 (sometimes temporary IP cool-down of minutes to hours).
Community-observed safe practices (2024–2026 GitHub, blogs, Reddit)
Approach	Interval / Rate	Notes	Risk Level
Very safe / continuous polling	5–10 seconds between requests	Ideal for 1–few symbols (e.g. SPY/SPX options chain every 5–10 s)	Lowest
Common conservative	2–3 seconds	Works for light loops	Low–medium
Aggressive but often used	1 request per second	Needs good User-Agent + backoff	Medium (can still trip after dozens–hundreds of calls)
Risky	<1 s or bursts of many tickers	Quick 429	High
Observed hard thresholds	~60 price-ish req/min or ~10 heavier (info/options) req/min in older reports; much stricter now	Unofficial & variable	—

5-second gap is on the safe side for continuous single-ticker (or very few tickers) options-chain polling. Many people successfully run 5–10 s loops for hours without issues. Going tighter (1–2 s) works for short batches but raises the chance of temporary blocks, especially with options endpoints (which are heavier than simple quotes/history).
Practical recommendations for your use-case

    Minimum safe continuous poll: time.sleep(5) (or better 8–10 s) between successive option_chain / quote calls.
    Always implement:
        Exponential backoff on 429 (start 5–10 s, double up to 60–300 s).
        Real browser-like User-Agent (or use curl_cffi impersonation — recent yfinance versions help with this).
        Caching (options data doesn’t meaningfully change every second; cache 30–60 s for most analytics).
    Prefer batching when possible (yf.download for multiple symbols) over tight per-ticker loops.
    Rotate IPs / proxies only if you need high volume; for one or two symbols a 5 s gap is usually enough.
    Monitor: if you start seeing intermittent 429s, increase the gap or add longer cool-downs.

Code-style example (safe single-symbol poll)

import time
import yfinance as yf
from yfinance.exceptions import YFRateLimitError

ticker = "SPY"
interval = 5   # seconds — safe starting point

while True:
    try:
        opt = yf.Ticker(ticker).option_chain()
        # process opt.calls / opt.puts
        print("Updated")
    except YFRateLimitError:
        print("Rate limited — backing off 60s")
        time.sleep(60)
        continue
    except Exception as e:
        print(e)
        time.sleep(10)
    time.sleep(interval)

Bottom line:

5-second gap is safe and conservative for ongoing options-chain polling of a small number of symbols. 10 seconds is even safer if you want maximum reliability and lowest chance of IP cool-down. Anything under ~2–3 s continuous is pushing it.

Still remember: this remains unofficial scraping — Yahoo can change endpoints or tighten further at any time, and commercial redistribution is against their ToS. For a production product, a proper paid vendor or broker API remains the more reliable long-term path.
40 sources

Snp500 bhi miljayega?? Aur agar koi stocks ka bhi chahiye hoga toh?

Haan, mil jayega.
S&P 500 Options

Yahoo Finance / yfinance se dono easily milte hain:
Symbol	Kya hai	Options chain	Notes
SPY	S&P 500 ETF	Full chain (calls + puts, OI, volume, IV, Greeks)	Sabse reliable aur popular. Almost har tool isko support karta hai.
SPX / ^SPX	Pure S&P 500 index options	Available	Kabhi-kabhi symbol SPX se better kaam karta hai, kabhi ^SPX. Test kar lena.

Example:

import yfinance as yf

# SPY (recommended)
spy = yf.Ticker("SPY")
chain = spy.option_chain()          # nearest expiry
# ya specific expiry:
chain = spy.option_chain("2026-10-17")

# SPX index
spx = yf.Ticker("SPX")              # ya "^SPX"
chain = spx.option_chain()

Individual Stocks

Bilkul same tarike se mil jayenge. Bas ticker change karo:

# Examples
aapl = yf.Ticker("AAPL").option_chain()
tsla = yf.Ticker("TSLA").option_chain()
nvda = yf.Ticker("NVDA").option_chain()
qqq  = yf.Ticker("QQQ").option_chain()   # Nasdaq-100 ETF

Almost saare liquid US stocks (AAPL, MSFT, TSLA, NVDA, AMD, META, AMZN, GOOGL, etc.) + major ETFs (SPY, QQQ, IWM, DIA) pe options chain aati hai.
Practical Tips (5-sec polling ke saath)

    Ek se zyada symbols chahiye ho to sequential polling karo (har request ke beech 5–8 second gap).
    Parallel mat chalao — rate-limit jaldi hit hoga.
    Example pattern:

    symbols = ["SPY", "QQQ", "AAPL", "TSLA"]
    for sym in symbols:
        try:
            chain = yf.Ticker(sym).option_chain()
            # process
        except:
            pass
        time.sleep(5)   # safe gap

    SPY + SPX dono parallel nahi, ek ke baad ek.

Summary:

SPY (S&P 500 ETF) → full support

SPX (index) → mostly support

Koi bhi liquid US stock → same method se mil jayega

Bas 5–10 second gap maintain karna, caching rakhna, aur 429 aaye to backoff.

Mtlb ab 2 cheeze fetch karre hai toh ekk he gap mai karna hai ki uske liye different


Yeh lo simple aur clean **Yahoo Finance (yfinance)** ka code. Isse aap options chain fetch karke dekh sakte ho.

### 1. Basic Options Chain Fetch (Single Stock)

```python
import yfinance as yf
import pandas as pd

# Symbol daalo (SPY, AAPL, TSLA, etc.)
ticker = yf.Ticker("SPY")

# Available expiry dates dekhne ke liye
print("Available Expiries:")
print(ticker.options)

# Specific expiry ka chain nikaalo
expiry = ticker.options[0]   # pehli expiry le rahe hain
chain = ticker.option_chain(expiry)

print("\n----- CALLS -----")
print(chain.calls[['contractSymbol', 'strike', 'lastPrice', 'bid', 'ask', 'volume', 'openInterest', 'impliedVolatility']].head(10))

print("\n----- PUTS -----")
print(chain.puts[['contractSymbol', 'strike', 'lastPrice', 'bid', 'ask', 'volume', 'openInterest', 'impliedVolatility']].head(10))
```

### 2. Multiple Symbols + Safe Rate Limit (Recommended)

```python
import yfinance as yf
import time
import pandas as pd

symbols = ["SPY", "QQQ", "AAPL"]   # jitne chaho add karo
gap = 6   # seconds gap (safe limit)

while True:
    for sym in symbols:
        try:
            ticker = yf.Ticker(sym)
            expiries = ticker.options
            
            if not expiries:
                print(f"{sym} → No options found")
                continue
                
            # Nearest expiry le rahe hain
            expiry = expiries[0]
            chain = ticker.option_chain(expiry)
            
            calls = chain.calls
            puts = chain.puts
            
            print(f"\n=== {sym} | Expiry: {expiry} ===")
            print(f"Calls: {len(calls)} | Puts: {len(puts)}")
            print(f"Total OI (Calls): {calls['openInterest'].sum():,.0f}")
            print(f"Total OI (Puts):  {puts['openInterest'].sum():,.0f}")
            
            # Optional: save karna ho to
            # calls.to_csv(f"{sym}_calls_{expiry}.csv", index=False)
            # puts.to_csv(f"{sym}_puts_{expiry}.csv", index=False)
            
        except Exception as e:
            print(f"Error on {sym}: {e}")
        
        time.sleep(gap)   # important - rate limit ke liye
    
    print("\n--- Round complete, waiting before next cycle ---\n")
    time.sleep(10)   # puri list ke baad thoda extra wait
```

### 3. Sirf Specific Expiry + Full Greeks/Volume/OI

```python
import yfinance as yf

ticker = yf.Ticker("SPY")
expiry = "2026-09-25"   # apni expiry daalo (YYYY-MM-DD)

chain = ticker.option_chain(expiry)

# Full data dekhne ke liye
print(chain.calls.columns.tolist())   # available columns

# Useful columns
cols = ['contractSymbol', 'strike', 'lastPrice', 'bid', 'ask', 
        'change', 'percentChange', 'volume', 'openInterest', 
        'impliedVolatility', 'inTheMoney']

print("\nCALLS:")
print(chain.calls[cols].to_string())

print("\nPUTS:")
print(chain.puts[cols].to_string())
```

### Notes:
- `openInterest` official end-of-day hota hai (live nahi).
- Real-time LTP/bid/ask thoda delayed ho sakta hai.
- Rate limit avoid karne ke liye **5–8 second gap** rakho.
- SPX ke liye symbol `"^SPX"` ya `"SPX"` try karo (kabhi-kabhi `SPX` better kaam karta hai).

Chaho to batao — main isme specific features (PCR, Max Pain, nearest ATM, multiple expiry, etc.) bhi add kar deta hoon.