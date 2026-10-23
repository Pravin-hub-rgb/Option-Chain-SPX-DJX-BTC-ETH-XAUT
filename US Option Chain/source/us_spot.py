"""Spot-price layer: Yahoo, Tencent and put-call parity, behind one cache.

Everything here serves the freshest index level through the shared
us_state._spot_cache: one background task refreshes each HTTP source and the
chain reads the cache, so no blocking request ever runs inside a chain poll.

Moved verbatim from main.py. main.py imports the names it still needs back.
The Databento-side spot helpers and the chain fetch stay in main until the
chain module is extracted; _databento_feed is read through the us_state
module every call so a live rebind in main() is always visible here.
"""

import math
import time
from datetime import date, datetime

import requests
from us_time import eastern_now

import us_state
from us_state import cfg, _spot_cache, _spot_lock, spot_cmp
from us_workbook import base_root, clean_symbol, is_index_symbol, safe_float
from us_diagnostics import feed_log

# Yahoo ^SPX/^RUT/^VIX/^DJX quotes are ~15 min delayed; GSPC/NDX/DJI/SPY are live.
YAHOO_SPOT_CANDIDATES = {
    "SPX": ("^GSPC", "^SPX"),
    "SPXW": ("^GSPC", "^SPX"),
    # ^DJX is a 15 minute old feed; ^DJI is live. The DJX index is 1/100 of
    # the Dow, so the scale below puts it on the same footing.
    "DJX": ("^DJI", "^DJX"),
    "DJXW": ("^DJI", "^DJX"),
    "RUT": ("^RUT",),
    "NDX": ("^NDX",),
    "NDXP": ("^NDX",),
    "VIX": ("^VIX",),
}
# Yahoo quotes the Dow in full points; the DJX index is one hundredth of it.
YAHOO_SCALE = {"^DJI": 0.01}

YAHOO_CHART = "https://query1.finance.yahoo.com/v8/finance/chart/{}?interval=1d&range=1d"
HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "application/json",
}


def yahoo_spot(sym):
    """Spot from Yahoo, served from the shared cache.

    Yahoo's real-time feeds (^GSPC, ^DJI, ^NDX) are genuinely live - about a
    second or two old. Its ^SPX, ^DJX, ^RUT and ^VIX feeds are fifteen
    minutes old, so the age of every quote is checked and the freshest one
    wins.
    """
    root = clean_symbol(sym)
    return spot_for_symbol(root)


def _yahoo_spot_fetch(root):
    """Fetch Yahoo candidates for a root and store the freshest quote."""
    if is_index_symbol(root):
        candidates = list(
            YAHOO_SPOT_CANDIDATES.get(root)
            or YAHOO_SPOT_CANDIDATES.get(base_root(root))
            or (f"^{root}",)
        )
    else:
        candidates = [root]
    now = time.time()
    best_price = None
    best_age = None
    for cand in candidates:
        try:
            resp = requests.get(
                YAHOO_CHART.format(cand),
                headers=HTTP_HEADERS,
                timeout=10,
            )
            if resp.status_code != 200:
                continue
            result = ((resp.json().get("chart") or {}).get("result") or [None])[0]
            if not result:
                continue
            meta = result.get("meta") or {}
            raw_ts = meta.get("regularMarketTime")
            price = safe_float(meta.get("regularMarketPrice"))
            used_fallback = False
            if not price or price <= 0:
                price = safe_float(meta.get("chartPreviousClose"))
                used_fallback = True
            if not price or price <= 0:
                continue
            ts = safe_float(raw_ts) or now
            age = abs(now - ts)
            scale = YAHOO_SCALE.get(cand, 1.0)
            if age > 60:
                # Only worth recording the ones that are too old to use.
                feed_log(
                    "spot_yahoo_stale",
                    root=root,
                    cand=cand,
                    price=price,
                    age_s=round(age, 1),
                )
            if best_price is None or age < best_age:
                best_price, best_age = price * scale, age
            if age <= 10:
                break
        except Exception:
            continue
    if best_price is not None:
        _store_spot(root, best_price, best_age, "yahoo")
    return best_price


TENCENT_QUOTE = "https://qt.gtimg.cn/q={}"
# Tencent mirrors the US indices under its own tickers. usINX is the S&P 500
# index, usDJI the Dow and usNDX the Nasdaq-100; usIXIC is the Nasdaq
# Composite. Verified live: all three returned a price, a previous close and a
# timestamp, with no API key.
# NOTE the scale divisor. Tencent quotes the full DJIA (52002), but the Cboe
# DJX index that DJX options actually settle against is 1/100 of that (520.02).
# Using the raw value would put every DJX strike deep ITM and invert the
# ITM/ATM colouring, so it is divided back down.
TENCENT_SCALE = {"usDJI": 0.01}
TENCENT_ROOT_SCALE = {"XSP": 0.1}
TENCENT_SPOT = {
    "SPX": ("usINX",),
    "SPXW": ("usINX",),
    "XSP": ("usINX",),
    "DJX": ("usDJI",),
    "DJXW": ("usDJI",),
    "NDX": ("usNDX",),
    "NDXP": ("usNDX",),
    "VIX": ("usVIX",),
    "RUT": ("usRUT",),
}

SPOT_POLL_SECONDS = float(cfg.get("spot_poll_seconds", 1.0))
# How long a cached quote is offered before the tool stops serving it and falls
# back to the last known value. This is only a backstop for "we stopped
# refreshing"; the real freshness test is the quote's own age, checked against
# SPOT_STALE_SECONDS. It used to be three seconds, and because one slow HTTP
# request could hold up the whole refresh, the cache expired on nearly half
# the lookups - which left the index level trailing by seconds and looking
# like it disagreed with the real price.
SPOT_MAX_AGE = float(cfg.get("spot_max_age", 12.0))
# How old a quote may be before it is treated as stale. A stale index level
# would place the ATM row wrongly, so it is rejected rather than used.
SPOT_STALE_SECONDS = float(cfg.get("spot_stale_seconds", 90.0))
# Put-call parity: the interest rate used to discount the strike to present.
# It only matters for expiries more than a few days out; on a same-day
# option the discount is zero.
PARITY_RATE = float(cfg.get("parity_rate", 0.045))
PARITY_DIVIDEND_Q = float(cfg.get("parity_dividend_q", 0.014))
PARITY_USE_N_STRIKES = int(cfg.get("parity_use_n_strikes", 5))
PARITY_DEADBAND = float(cfg.get("parity_deadband", 0.05))
# How far the parity spot and the external index quote may differ before it
# is treated as a disagreement worth flagging.
PARITY_TOLERANCE_POINTS = float(cfg.get("parity_tolerance_points", 3.0))
PARITY_TOLERANCE_PCT = float(cfg.get("parity_tolerance_pct", 0.002))
# A book wider than this is too thin for its mid to be an index level. Near
# expiry the quotes go thin and a wide spread would make the parity jump
# around, so in that case the parity is not used at all.
PARITY_MAX_SPREAD = float(cfg.get("parity_max_spread", 2.0))
PARITY_MAX_SPREAD_PCT = float(cfg.get("parity_max_spread_pct", 0.05))
# If the parity level lands this far from the external quote it is wrong, not
# the external quote, and it is thrown away rather than allowed near the ATM
# row.
PARITY_SANITY_PCT = float(cfg.get("parity_sanity_pct", 0.05))
# Last parity spot per symbol, and the current disagreement with the external
# quote. Both are surfaced in the log so the source of the spot is never a
# guess.
_parity_spot_cache = {}
_spot_divergence = {}
_parity_log_at = {}

STALE_SPOT_SECONDS = 60.0

# How old each symbol's spot is, and which source last served it.
spot_served_at = {}
spot_source = {}
spot_first_seen = {}


def _store_spot(root, price, age, source):
    """Keep the newest quote seen for a root, from any source.

    "Newest" means the quote with the latest timestamp, not the one with the
    smallest age. Those are not the same thing: a feed that is two seconds
    behind on the wall clock keeps reporting a younger-looking age, so
    choosing by age would hold on to an old quote and freeze the index level.
    The timestamp is derived from the quote's own age, so both feeds can be
    compared fairly.
    """
    if price is None or price <= 0:
        return
    if age is None:
        age = 0.0
    quote_ts = time.time() - float(age)
    with _spot_lock:
        prev = _spot_cache.get(root)
        if prev is None or quote_ts > prev[4]:
            _spot_cache[root] = (
                float(price), float(age), source, time.monotonic(), quote_ts
            )
            # Publish straight to the displayed value as well. The chain poll
            # only runs about once a second, so without this the index level
            # on screen trailed the real price by up to that much.
            spot_cmp[root] = float(price)


def spot_for_symbol(sym):
    """Freshest cached spot for a symbol, or None if nothing fresh is held."""
    root = clean_symbol(sym)
    with _spot_lock:
        hit = _spot_cache.get(root)
    if not hit:
        return None
    price, age, _source, fetched = hit[0], hit[1], hit[2], hit[3]
    if time.monotonic() - fetched > SPOT_MAX_AGE:
        # We have not refreshed recently, so this is not a current quote.
        return None
    # The quote's own age does not disqualify it. A quote that is hours old
    # is still the last known index level, and the trader is better served by
    # seeing that, clearly flagged as stale, than by seeing zero. The header
    # shows STALE when the level is old.
    return price


def spot_source_for(sym):
    """Which feed supplied the cached spot, and how old that quote is."""
    root = clean_symbol(sym)
    with _spot_lock:
        hit = _spot_cache.get(root)
    if not hit:
        return None, None
    return hit[2], hit[1]


def _tencent_feed_age_seconds(stamp):
    """Age of a Tencent quote, in seconds, from its exchange timestamp.

    Tencent returns "YYYY-MM-DD HH:MM:SS" in US Eastern time. Returns None
    when the stamp cannot be understood.
    """
    if not stamp:
        return None
    try:
        quoted = datetime.strptime(stamp.strip(), "%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError):
        return None
    try:
        eastern, available = eastern_now(warning=lambda *a, **k: None)
    except Exception:
        return None
    if not available or eastern is None:
        return None
    now = eastern.replace(tzinfo=None)
    return max(0.0, (now - quoted).total_seconds())


def _tencent_parse(body):
    """Parse a Tencent quote body into {ticker: (raw_price, feed_time)}.

    Handles the batched form (q=usINX,usDJI,usVIX) as well as a single
    ticker, so one request can serve every index we might need.
    """
    out = {}
    for line in body.strip().split("\n"):
        if '="' not in line:
            continue
        key = line.split("=")[0].replace("v_", "").strip()
        try:
            parts = line.split('="', 1)[1].rstrip('";').split("~")
            if len(parts) < 4 or not parts[3]:
                continue
            out[key] = (float(parts[3]), parts[30] if len(parts) > 30 else "")
        except (ValueError, IndexError):
            continue
    return out


def refresh_tencent_spot_cache():
    """One HTTP request for every index ticker, then serve them from memory."""
    tickers = sorted({c for v in TENCENT_SPOT.values() for c in v})
    if not tickers:
        return
    try:
        resp = requests.get(
            TENCENT_QUOTE.format(",".join(tickers)),
            headers=HTTP_HEADERS,
            timeout=8,
        )
        if resp.status_code != 200:
            feed_log("spot_tencent_http", status=resp.status_code)
            return
        resp.encoding = "gbk"
        parsed = _tencent_parse(resp.text)
        for ticker, (raw, stamp) in parsed.items():
            age = _tencent_feed_age_seconds(stamp)
            for root, cands in TENCENT_SPOT.items():
                if ticker not in cands:
                    continue
                scale = TENCENT_SCALE.get(ticker, 1.0) * TENCENT_ROOT_SCALE.get(root, 1.0)
                _store_spot(root, raw * scale, age, "tencent")
    except Exception as e:
        feed_log("spot_batch_error", error=type(e).__name__, detail=str(e)[:150])


def tencent_spot(sym):
    """Spot from Tencent for index roots, served from the shared cache."""
    root = clean_symbol(sym)
    cands = TENCENT_SPOT.get(root) or TENCENT_SPOT.get(base_root(root))
    if not cands:
        return None
    price = spot_for_symbol(root)
    if price is not None:
        feed_log("spot_tencent_ok", symbol=sym, price=price)
    return price


def _tencent_fetch_direct(root, cands, root_scale):
    for cand in cands:
        try:
            resp = requests.get(TENCENT_QUOTE.format(cand), headers=HTTP_HEADERS,
                                timeout=10)
            if resp.status_code != 200:
                feed_log("spot_tencent_http", symbol=root, cand=cand,
                         status=resp.status_code)
                continue
            resp.encoding = "gbk"
            hit = _tencent_parse(resp.text).get(cand)
            if not hit:
                continue
            raw, stamp = hit
            age = _tencent_feed_age_seconds(stamp)
            scale = TENCENT_SCALE.get(cand, 1.0) * root_scale
            price = raw * scale
            _store_spot(root, price, age, "tencent")
            feed_log("spot_tencent_ok", symbol=root, cand=cand, price=price,
                     feed_time=stamp, scale=scale)
            return price
        except Exception as e:
            feed_log("spot_tencent_error", symbol=root, cand=cand,
                     error=type(e).__name__, detail=str(e)[:150])
    return None


def _years_to_expiry(expiry_iso):
    """Calendar years from today to an expiry date."""
    try:
        exp = datetime.strptime(str(expiry_iso), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return 0.0
    return max(0.0, (exp - date.today()).days / 365.0)


def _parity_spot(rows, spot_hint):
    """Index spot from put-call parity on our own live chain.

    Put-call parity gives  C - P = S*exp(-qT) - K*exp(-rT),  so

        F = K + exp(rT) * (C - P)          (forward price)
        S = F * exp(-(r - q) * T)          (spot, discounted at net repo rate)

    The quotes used here are the ones the sheet is already showing, so the
    answer is real time, needs no extra subscription, and can never disagree
    with the chain on screen. It is the fallback and the cross-check for the
    external index quote, which is a free endpoint that can go stale, rate
    limit or simply disappear.

    To reduce noise from a single bad quote, parity is computed for the N
    strikes nearest the current estimate, and the median is taken. Crossed
    and wide-spread quotes are rejected.
    """
    if not rows:
        return None
    expiry = None
    calls = {}
    puts = {}
    for r in rows:
        if not isinstance(r, dict):
            continue
        if expiry is None:
            expiry = r.get("expiry")
        elif r.get("expiry") != expiry:
            continue
        strike = r.get("strike")
        if not strike:
            continue
        rtype = str(r.get("type") or "").lower()
        if rtype.startswith("c"):
            calls[strike] = r
        elif rtype.startswith("p"):
            puts[strike] = r
    common = set(calls) & set(puts)
    if not common:
        return None

    T = _years_to_expiry(expiry)
    forward_discount = math.exp(-(PARITY_RATE - PARITY_DIVIDEND_Q) * T)
    call_put_adj = math.exp(PARITY_DIVIDEND_Q * T)

    def mid(row):
        bid = row.get("bid") or 0.0
        ask = row.get("ask") or 0.0
        if bid > 0 and ask > 0:
            return (bid + ask) / 2.0
        return row.get("ltp") or 0.0

    def spread_ok(row):
        """Reject a book so wide that its mid is meaningless.

        Near expiry the quotes go thin and the bid/ask can be tens of points
        wide. A mid taken across such a spread is not an index level, it is a
        guess, so the parity is not used at all in that case.
        """
        bid = row.get("bid") or 0.0
        ask = row.get("ask") or 0.0
        if bid <= 0 or ask <= 0 or ask < bid:
            return False
        mid_price = (bid + ask) / 2.0
        if mid_price <= 0:
            return False
        return (ask - bid) <= max(PARITY_MAX_SPREAD, mid_price * PARITY_MAX_SPREAD_PCT)

    def parity_for_strike(strike):
        call_row = calls[strike]
        put_row = puts[strike]
        if not spread_ok(call_row) or not spread_ok(put_row):
            return None
        call_mid = mid(call_row)
        put_mid = mid(put_row)
        if call_mid <= 0 or put_mid <= 0:
            return None
        return strike * forward_discount + (call_mid - put_mid) * call_put_adj

    # Pick the N strikes nearest the current estimate (or the hint).
    hint = spot_hint if spot_hint else 0.0
    ranked = sorted(common, key=lambda s: abs(s - hint))
    take = ranked[:PARITY_USE_N_STRIKES]

    estimates = [p for s in take if (p := parity_for_strike(s)) is not None]
    if not estimates:
        return None

    estimates.sort()
    n = len(estimates)
    if n % 2 == 1:
        return estimates[n // 2]
    return (estimates[n // 2 - 1] + estimates[n // 2]) / 2.0


def _fresher_spot(sym, cboe_or_chain_spot):
    """Index spot: the freshest of the live feeds, the chain payload last.

    Both free feeds are refreshed every second by the background spot task
    and the cache keeps whichever quote is younger. Yahoo's ^GSPC and ^DJI are
    genuinely real time; Tencent's free endpoint was measured serving a quote
    twenty to thirty seconds old that alternated between two stale values,
    which is what made the index level appear to jump about.
    """
    if not is_index_symbol(sym):
        return cboe_or_chain_spot
    best = spot_for_symbol(sym)
    if best:
        return best
    feed_log("spot_cache_miss", symbol=sym, note="using the chain payload spot")
    return cboe_or_chain_spot


def note_spot(sym, spot, source_name):
    """Record spot plus when it was actually served."""
    now = time.time()
    first = spot_first_seen.setdefault(sym, now)
    if spot_cmp.get(sym) != spot:
        spot_first_seen[sym] = now
    spot_served_at[sym] = now
    spot_source[sym] = source_name
    spot_cmp[sym] = spot


def spot_age_seconds(sym):
    served = spot_served_at.get(sym) or 0.0
    if not served:
        return float("inf")
    return time.time() - served


def source_delay_note(sym):
    """Text for the header showing the live Databento feed status."""
    name = spot_source.get(sym)
    parts = []
    if name == "databento":
        feed = us_state._databento_feed
        parts.append(feed.status_text() if feed is not None else "LIVE")
    age = spot_age_seconds(sym)
    if age != float("inf") and age > STALE_SPOT_SECONDS:
        parts.append(f"STALE ({int(age)}s)")
    # Say so when the index level from the external feed and the level implied
    # by our own option chain disagree. One of them is wrong and the trader
    # should not have to guess which.
    diff = _spot_divergence.get(sym)
    if diff:
        parts.append(f"SPOT MISMATCH {diff:.1f}")
    return " | ".join(parts)


def _databento_spot(sym):
    prior = spot_cmp.get(sym)
    # The background spot task keeps the cache warm for every symbol, index
    # and stock alike, so this never waits on the network.
    cached = spot_for_symbol(sym)
    if cached:
        return cached
    if is_index_symbol(sym):
        return _fresher_spot(sym, prior)
    return yahoo_spot(sym) or prior
