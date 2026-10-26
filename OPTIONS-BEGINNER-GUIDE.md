# Options From Zero - A Trader's Guide

## 0. How to use this guide

- **Sections 1-7**: pure basics. If you don't know what a strike is, start here.
- **Sections 8-10**: single-leg trades and the idea of a "spread."
- **Sections 11-13**: calendar spreads - this is what your Combo Rate column is for.
- **Sections 14-17**: risk, mistakes, glossary, practice plan.

Skim first, then re-read the parts that confuse you.

---

## 1. Free courses & what to search on YouTube

You don't need to pay. These are free and good:

**YouTube channels (search these names + "options basics"):**
- **tastytrade** - best for plain-English, no-BS explanations.
- **projectfinance** - very beginner-friendly, step-by-step.
- **Option Alpha** - structured free course.
- **Markus Heitkoetter / Rockwell Trading** - good for beginners.
- **Khan Academy** - more academic but solid.

**Specific videos to search (copy-paste these search phrases):**
1. `how to read an option chain for beginners`
2. `what is a strike price option trading`
3. `call vs put option explained simply`
4. `bid ask spread options explained`
5. `open interest vs volume options`
6. `implied volatility explained for beginners`
7. `what is a calendar spread options`
8. `calendar spread options strategy explained`
9. `options theta decay explained`
10. `options vega explained`

**Free structured courses:**
- **CBOE (Chicago Board Options Exchange)** - free "Options Institute" online.
- **Interactive Brokers "Traders' Academy"** - free, very thorough.
- **tastytrade "Learn Center"** - free articles + videos.

**Suggested order of study:** Section 2 -> video #1 above -> video #3 -> video #4 -> video #7 -> then paper trade.

---

## 2. What is an option? (60-second version)

An **option** is a contract that gives you the **right, but not the obligation**, to buy or sell an asset (like BTC, SPX, or a stock) at a fixed price, by a fixed date.

- You **pay a price** (called the **premium**) to get this right.
- If the market moves in your favor, you profit.
- If it doesn't, the most you lose is the premium you paid (for buying options).

Think of it like a **refundable deposit to lock in a price**. You're paying a small fee for the *option* to act later.

**Two parties:**
- **Buyer** (you, usually): pays premium, has the right.
- **Seller** (market maker / another trader): receives premium, has the obligation.

---

## 3. Call vs Put

| | **Call** | **Put** |
|---|---|---|
| Right to | **Buy** the asset at strike price | **Sell** the asset at strike price |
| You profit when | Price goes **UP** | Price goes **DOWN** |
| You use it when | You're bullish | You're bearish |
| Max loss (buyer) | Premium paid | Premium paid |

**Memory trick:**
- **C**all = **C**limb (price climbs up)
- **P**ut = **P**lunges (price plunges down)

A call is a bet the price rises. A put is a bet the price falls.

---

## 4. Strike price

The **strike** is the fixed price written into the contract.

Example: A call with strike 80000 means "I have the right to buy BTC at 80000, no matter what the market price is."

- If BTC is at 82000, your 80000 call is valuable (you can buy cheaper than market).
- If BTC is at 78000, your 80000 call is worthless right now.

**In-the-money (ITM):** strike is favorable vs current price (call: strike < price; put: strike > price).
**Out-of-the-money (OTM):** strike is unfavorable.
**At-the-money (ATM):** strike is approximately current price.

---

## 5. Expiry (expiration)

Every option **dies on a specific date**. After that, it's worthless (if not exercised).

- **Near expiry** = expires soon (days/weeks). Cheaper, decays fast.
- **Far expiry** = expires later (months). More expensive, decays slowly.

This difference in expiry is exactly what a **calendar spread** trades.

---

## 6. Premium

The **premium** is the price you pay (or receive) for the option. It's what the option is worth today.

Premium = **Intrinsic value** + **Time value**
- Intrinsic value = how much the option is "in the money" right now.
- Time value = extra worth because there's still time for the price to move in your favor.

As expiry approaches, time value shrinks to zero. This shrinking is **theta decay** - the most important concept in options.

---

## 7. Reading an option chain

Your BTC sheet IS an option chain. Here's what each column means:

| Column (typical) | Meaning |
|---|---|
| **Strike** | The fixed price of the contract |
| **Call LTP / Bid / Ask** | Last traded / best buy / best sell price for calls |
| **Put LTP / Bid / Ask** | Same for puts |
| **Call OI** | Open Interest - how many call contracts are currently open |
| **Put OI** | Same for puts |
| **IV (implied volatility)** | Market's expectation of how much the price will move (higher = more expensive options) |

**Bid vs Ask:**
- **Bid** = price someone will PAY to buy from you (you sell at bid).
- **Ask** = price someone will SELL to you (you buy at ask).
- The gap between them is the **spread** (cost of trading). Tighter = better.

**Open Interest (OI):**
- Number of contracts currently open (not yet closed/exercised).
- High OI = liquid market, easy to enter/exit.
- Rising OI with rising price = new money entering (trend continuation).
- Falling OI = positions closing (trend weakening).

**Volume:** how many contracts traded today. High volume = active interest.

**Key rule:** Always trade strikes with **high OI and tight bid/ask**. Avoid illiquid strikes - you'll get bad fills.

---

## 8. How a single option makes / loses money

**Long call (you bought a call):**
- Profit if price rises above (strike + premium paid).
- Loss = premium paid (if price stays below strike).
- Unlimited upside, limited downside. Beginner-friendly for learning.

**Long put (you bought a put):**
- Profit if price falls below (strike - premium paid).
- Loss = premium paid.

**The catch:** Even if you're right about direction, **time decay** works against you. The option loses value every day. You need the price to move **fast enough and far enough** to overcome theta.

---

## 9. Basic strategies (start here, not calendars)

Before calendars, understand these:

1. **Long call / long put** - directional bet. Simplest. Learn payoff first.
2. **Covered call** - you own the stock/BTC, sell a call against it. Collects premium. Caps upside. Income strategy.
3. **Cash-secured put** - you set aside cash, sell a put. If price drops, you buy at strike. Collects premium.
4. **Vertical spread** - buy one strike, sell another at same expiry. Limits both risk and reward. Good stepping stone to calendars.

**Practice these first.** Calendars are "level 2."

---

## 10. What is a "spread"?

A spread = **two or more options combined** to create a specific risk/reward profile.

Instead of one big bet, you combine legs so that:
- Risk is defined (you know max loss upfront).
- You trade a *relationship* (between strikes, months, or call vs put) rather than a direction.

Calendar spread = spread across **two expiry months** at the **same strike**.

---

## 11. Calendar spread - deep dive

### The setup
- **Same strike**, **two different expiries**.
- Example: strike 80000, near expiry (this week) + far expiry (next month).

### Two versions

| | **Long calendar** (buy far, sell near) | **Short calendar** (buy near, sell far) |
|---|---|---|
| You pay (debit) or receive (credit)? | Pay debit | Receive credit |
| You profit when | The gap **widens** | The gap **narrows** |
| Best when | Volatility rises / price pins near strike | Volatility falls / price moves away |
| Max risk | Premium paid (debit) | Large (the far leg can run) |

### Why it works (the "why")
- **Theta**: near month decays faster -> if price stays near strike, the near leg you sold loses value faster than the far leg you bought -> you win.
- **Vega**: far month has more vega -> if IV rises, far leg gains more -> you win (long calendar).
- The "sweet spot": price ends **exactly at your strike at near expiry**. Then near leg = 0, far leg still has value -> max profit.

### Payoff shape
Long calendar: profit peaks at the strike at near-expiry, falls off on both sides. Looks like a tent.

### Greeks in plain English
- **Theta (time decay)**: your friend in a long calendar (near decays, you sold it).
- **Vega (volatility)**: your friend in a long calendar (far has more vega).
- **Delta (direction)**: small near expiry; calendar is mostly non-directional.
- **Gamma (acceleration)**: spikes near expiry - this is the danger zone.

---

## 12. Reading YOUR Combo Rate column

Your column: `Combo Rate = Near Call Ask - Far Call Ask`.

This is **exactly the price of a short calendar** (buy near, sell far) at the ask. One number = the whole 2-leg trade's cost.

### How to read it

| Combo value | What it means | Candidate action |
|---|---|---|
| Very negative (e.g. -271) | Far month much more expensive -> market pricing far rich | Short calendar candidate (sell far / buy near) - bet gap narrows |
| Near 0 | Both months similar -> no edge | Skip |
| Positive (e.g. +38) | Near month more expensive -> rare | Long calendar candidate (buy far / sell near) |

### Finding the trade
Don't trade the middle. Look for **strikes where the combo is an outlier** vs its neighbors.

In your data:
- 81000 -> -271 (local extreme negative)
- 82600 -> +38 (positive, rare)

Those are your candidates - the market is pricing the two months unusually differently there.

### Entry / exit
1. **Enter** when combo hits an extreme vs its recent range (watch it for a few days to know the range).
2. **Exit** when combo reverts toward its average, or when near expiry is close (theta danger).
3. **Watch live** - the column ticks every update. Note whether it's widening or contracting.

### Bid/ask caveat (read this)
Your combo uses **asks on both legs** = worst-case (most expensive) quote. Real fills:
- Leg you **buy** -> fill at **ask**.
- Leg you **sell** -> fill at **bid**.

So actual cost = `near_ask - far_bid` (for buy-near/sell-far), which is **better** (less negative) than shown. The displayed combo is a conservative estimate. Always check real bid/ask before sending.

---

## 13. Taking a calendar position: step by step

### Before you trade
- Paper trade for at least 2 weeks first.
- Pick ONE liquid strike (high OI, tight bid/ask, near ATM).
- Decide: long calendar (pay debit) or short calendar (collect credit). Start with **long calendar** - defined risk, easier to learn.

### Entry (long calendar example)
1. Note the combo rate at your strike right now (e.g. -248 at 80000).
2. Place the order as a **spread** (not two separate orders) so both legs fill together:
   - BUY far-expiry call at strike 80000
   - SELL near-expiry call at strike 80000
3. Net cost = debit (you pay). Confirm the debit matches ~combo (adjusted for bid/ask).
4. Set a mental/actual stop: e.g. "exit if combo moves 20% against me or near expiry is 3 days away."

### Management
- Check the combo column daily. Is it widening (good for long) or narrowing (bad)?
- Watch IV: rising IV helps long calendar; falling IV hurts.
- Watch theta: as near expiry approaches, the spread decays faster. Don't hold to last day.

### Exit
- **Profit target**: combo reverts toward its average (e.g. -248 -> -200).
- **Stop**: combo moves further against you (e.g. -248 -> -300) OR near expiry < 5 days.
- **Time stop**: close 3-5 days before near expiry regardless.

### Position sizing
- Risk only 1-2% of account per trade.
- Long calendar max loss = debit paid. Size so that max loss = your risk budget.

---

## 14. Risk & sizing

- **Long calendar**: max loss = debit paid. Defined. Good for beginners.
- **Short calendar**: max loss is large (far leg can run). Avoid until experienced.
- **Assignment risk**: if your short near leg is ITM at expiry, you may be assigned. Close before expiry.
- **Liquidity risk**: wide bid/ask = you lose on entry AND exit. Only trade liquid strikes.
- **Event risk**: earnings, FOMC, CPI can gap the price past your strike. Avoid holding calendars through major events.

---

## 15. Common beginner mistakes

1. **Trading illiquid strikes** - you'll get terrible fills. Check OI + bid/ask first.
2. **Holding to expiry** - theta accelerates and the spread collapses. Exit early.
3. **Ignoring vega** - a volatility crush can kill a long calendar even if direction is right.
4. **Sizing too big** - calendars bleed slowly; big size = big slow loss.
5. **Chasing the most negative combo** without knowing its normal range. Watch for a week first.
6. **Treating combo as a signal, not a measurement** - it's a price, not a prediction.
7. **Not using spread orders** - legging in separately = slippage and partial fills.

---

## 16. Glossary

- **Option**: contract giving right to buy/sell at a strike by an expiry.
- **Call**: right to buy. Profits on price rise.
- **Put**: right to sell. Profits on price fall.
- **Strike**: fixed price in the contract.
- **Expiry**: date the option dies.
- **Premium**: price of the option.
- **Bid**: price you sell at.
- **Ask**: price you buy at.
- **Spread (bid-ask)**: cost of trading.
- **OI (Open Interest)**: open contracts outstanding.
- **Volume**: contracts traded today.
- **IV (Implied Volatility)**: market's expected move size.
- **Theta**: time decay (value lost per day).
- **Vega**: sensitivity to IV changes.
- **Delta**: sensitivity to price changes.
- **Gamma**: rate of change of delta.
- **ITM / OTM / ATM**: in/out/at the money.
- **Calendar spread**: same strike, two expiries.
- **Long calendar**: buy far, sell near (debit).
- **Short calendar**: buy near, sell far (credit).
- **Combo rate**: near ask - far ask (price of short calendar at ask).

---

## 17. 30-day practice plan

**Week 1: Learn**
- Watch the 10 YouTube videos listed in Section 1 (one per day).
- Read Sections 2-7 of this guide.
- Open your broker's paper account. Find the option chain for any liquid underlying.

**Week 2: Single legs**
- Paper trade 5 long calls and 5 long puts. Track entry, exit, P&L, theta.
- Read Sections 8-9.
- Note how theta decay feels in real time.

**Week 3: Spreads**
- Paper trade 3 vertical spreads.
- Read Sections 10-11.
- Start watching your Combo Rate column daily. Note the range of values per strike.

**Week 4: Calendars**
- Paper trade 2 long calendars at liquid strikes.
- Read Sections 12-15.
- Track combo movement vs your entry. Did it revert? How long did it take?

**After 30 days:** Review your paper journal. If profitable and consistent, start with 1 real lot. If not, repeat week 4.

---

*End of guide. Re-read sections as needed. The market rewards patience and punishes rushing.*

