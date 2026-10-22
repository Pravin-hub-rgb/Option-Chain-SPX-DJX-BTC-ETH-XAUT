"""Symbol-state, expiry, build and poller layer.

Extracted from main.py as part of the module split. State lives in us_state;
spot/quote logic in us_spot; workbook files in us_workbook. A few excel-side
helpers that still live in main are reached by a late import (write_hint_labels,
find_open_book and the block payload caches) - they move out in later split
steps.
"""

import asyncio
import re
import time

import pandas as pd

import us_state
from us_state import (
    DEFAULT_SYMBOL, EXPIRY_REFRESH, ROLE_ORDER, SOURCE_MIN_INTERVAL,
    SYMBOL_CELL, SYMBOL_POLL_SECONDS, _feed_access, _last_symbol_call,
    _pending_date_loads, active_expiries, all_known_expiries, expiry_roots,
    live_data, oi_roots, role_symbol, spot_cmp, user_pinned,
)
from us_diagnostics import append_runtime_warning, build_id, feed_log
from us_spot import _databento_spot
from us_time import eastern_now, filter_expiry_dates
from us_workbook import (
    base_root, clean_symbol, default_k1, is_index_symbol, is_locked_role,
    normalize_symbol, safe_float,
)
from excel_helpers import prefer_expiry_root


async def throttled(fn, *args, **kwargs):
    """Rate-limit a fetch, per symbol.

    This used to be one global gate: every fetch had to wait until two
    seconds after the previous one had *started*, and the lock was held for
    the whole wait. With three symbols that put every symbol about six
    seconds behind, which is what the client was seeing.

    The data now comes from a live WebSocket, so the fetch is a local read
    of the quotes already in memory. There is no remote service to protect,
    so the only thing worth limiting is how often the *same* symbol is
    rebuilt. The actual feed access is serialised by a semaphore so two
    symbols never mutate the feed at once, but no artificial delay is added.
    """
    key = args[0] if args else None
    now = time.monotonic()
    last = _last_symbol_call.get(key, 0.0)
    wait = last + SOURCE_MIN_INTERVAL - now
    if wait > 0:
        await asyncio.sleep(wait)
    us_state._last_source_call = time.monotonic()
    if key is not None:
        _last_symbol_call[key] = time.monotonic()
    for attempt in range(3):
        try:
            async with _feed_access:
                return await asyncio.to_thread(fn, *args, **kwargs)
        except Exception as e:
            if attempt == 2:
                raise
            backoff = 15 * (attempt + 1)
            print(f"Fetch error, retry in {backoff}s: {e}", flush=True)
            await asyncio.sleep(backoff)
            if key is not None:
                _last_symbol_call[key] = time.monotonic()




def _note_expiry_roots(sym, rows):
    """Record which roots exist per expiry, and whether any real OI showed up."""
    base = base_root(sym)
    seen = expiry_roots.setdefault(base, {})
    oi = oi_roots.setdefault(base, {})
    live = set()
    for r in rows:
        exp = r.get("expiry")
        root = r.get("root") or base
        if not exp:
            continue
        live.add(exp)
        roots = seen.setdefault(exp, set())
        roots.add(root)
        if safe_float(r.get("oi")) > 0:
            oi[exp] = oi.get(exp, 0) + 1
    # drop expiries no longer in the payload (e.g. after a source switch)
    for exp in list(seen):
        if exp not in live:
            del seen[exp]
            oi.pop(exp, None)


def standard_expiries(sym, all_expiries):
    """Rank expiries: standard contract first, then ones that actually have OI.

    An index like DJX lists weekly chains (DJXW only) every few days plus
    monthly/quarterly DJX. Sorting purely by date makes the default slots show
    the mini weekly chain, which is what users reported seeing.

    Within the standard group, prefer expiries carrying real open interest:
    CBOE sometimes lists a standard root for a near expiry but publishes OI=0
    for every strike, which leaves the default slot full of zeros. Those are
    still selectable, just not in the first slots. Weekly expiries follow.
    """
    base = base_root(sym)
    roots = expiry_roots.get(base) or {}
    std = [d for d in all_expiries if base in (roots.get(d) or ())]
    weekly = [d for d in all_expiries if d not in set(std)]

    def has_oi(exp):
        return bool(oi_roots.get(base, {}).get(exp))

    std.sort(key=lambda d: (0 if has_oi(d) else 1,))
    return std + weekly


# Configured once in main(), after the owner/customer key has been resolved.


def active_symbols():
    """Unique symbols currently shown across all files (shared cache)."""
    out = []
    for role in ROLE_ORDER:
        s = role_symbol[role]
        if s and s not in out:
            out.append(s)
    return out


def ensure_sym_state(sym):
    if sym not in live_data:
        live_data[sym] = {}
        spot_cmp[sym] = None
        active_expiries[sym] = [None, None, None]
        user_pinned[sym] = [False, False, False]
        all_known_expiries[sym] = []

MONTHS = {"JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06",
          "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12"}

CHAIN_COLS = ["Call LTP", "Call Bid", "Call Ask", "Call OI",
              "strike", "Put OI", "Put Bid", "Put Ask", "Put LTP"]

def safe_float(val, default=0.0):
    if val is None:
        return default
    try:
        f = float(val)
    except (ValueError, TypeError):
        return default
    if math.isnan(f) or math.isinf(f):
        return default
    return f


def fmt_oi(val):
    # Force USD only â€” strip any â‚¹/INR and coerce to float
    if isinstance(val, str):
        s = val.replace("â‚¹", "").replace("Rs", "").replace("INR", "").replace("$", "").replace(",", "").strip()
        try:
            val = float(s)
        except ValueError:
            return "$0"
    try:
        val = float(val)
    except (TypeError, ValueError):
        return "$0"
    if val >= 1_000_000:
        return f"${val / 1_000_000:.2f}M"
    if val >= 1_000:
        return f"${val / 1_000:.2f}K"
    return f"${val:.0f}"


def read_symbol_cell(sheet, role):
    locked = default_k1(role) if is_locked_role(role) else None
    if locked is not None:
        # locked â€” force fixed ticker (user edits overwritten)
        if sheet.range(SYMBOL_CELL).value != locked:
            sheet.range(SYMBOL_CELL).value = locked
        return locked
    # Unlocked (stock): never write back into K1 while the user is typing.
    # Rewriting the cell every cycle fights the user's keystrokes and makes
    # typed text disappear. Just read + normalize internally; the cell keeps
    # exactly what the user typed.
    raw = sheet.range(SYMBOL_CELL).value
    return normalize_symbol(raw) or role_symbol["stock"] or DEFAULT_SYMBOL


def iso_to_user(iso):
    # DD/MM/YYYY
    if not iso or len(iso) < 10:
        return ""
    yyyy, mm, dd = iso[:10].split("-")
    return f"{dd}/{mm}/{yyyy}"


def parse_user_date(s):
    s = str(s).strip()
    # primary: DD/MM/YYYY
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        dd, mm, yyyy = m.groups()
        mi, di = int(mm), int(dd)
        if 1 <= mi <= 12 and 1 <= di <= 31:
            return f"{yyyy}-{mi:02d}-{di:02d}"
        return None
    # secondary: DD-MMM-YY / DDMMMYY (old)
    s2 = s.upper()
    m = re.match(r"^(\d{1,2})[- ]?([A-Z]{3})[- ]?(\d{2,4})$", s2)
    if m:
        dd, mmm, yy = m.groups()
        mm = MONTHS.get(mmm)
        if not mm:
            return None
        if len(yy) == 2:
            yy = "20" + yy
        return f"{yy}-{mm}-{int(dd):02d}"
    # ISO YYYY-MM-DD
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return s
    return None


def _price(ltp, bid, ask):
    ltp = safe_float(ltp)
    bid = safe_float(bid)
    ask = safe_float(ask)
    if ltp > 0:
        return ltp
    if bid > 0 and ask > 0:
        return (bid + ask) / 2
    # CBOE no-quote often sends bid=0, ask=20 sentinel â€” never mark lone ask as LTP
    if bid > 0:
        return bid
    return 0.0


def _is_no_quote(bid, ask):
    """CBOE illiquid sentinel: bid=0, ask=20 (exact)."""
    return safe_float(bid) == 0.0 and safe_float(ask) == 20.0


def _quote_quality(bid, ask, oi, ltp):
    """Higher is better â€” used when same strike has DJX vs DJXW (or SPX vs SPXW)."""
    b, a = safe_float(bid), safe_float(ask)
    real_book = 1 if (b > 0 and a > 0 and a >= b and not _is_no_quote(b, a)) else 0
    return (real_book, 1 if safe_float(ltp) > 0 else 0, safe_float(oi))


def _fetch_databento(sym):
    feed = us_state._databento_feed
    if feed is None:
        raise RuntimeError("Databento key is not configured")
    result = feed.snapshot(
        sym,
        _databento_spot(sym),
        active_expiries.get(sym),
        wait=True,
    )
    if result is None:
        errors = feed.errors()
        raise RuntimeError(errors[-1] if errors else "Databento chain not ready")
    # No quote record in the last few seconds is normal outside trading
    # hours, and it is not a reason to throw the chain away. The contract
    # list, the strikes and the open interest all come from Databento and
    # remain correct; only the live bid/ask/LTP are missing, and the header
    # says so. Failing here used to leave the sheet completely blank.
    if not feed.healthy() and not result[1]:
        errors = feed.errors()
        raise RuntimeError(errors[-1] if errors else "Databento live quotes not ready")
    _note_expiry_roots(sym, result[1])
    return result


def fetch_symbol_payload(sym):
    """Return (spot, rows, expiries, source_name). Databento is mandatory."""
    root = clean_symbol(sym)
    errors = []
    try:
        result = _fetch_databento(sym)
    except Exception as e:
        errors.append(f"databento: {e}")
        feed_log("source_failed", symbol=sym, root=root, source="databento",
                 error=type(e).__name__, detail=str(e)[:180])
        raise RuntimeError("; ".join(errors) or "no data source")
    spot_source_name, spot_age = spot_source_for(sym)
    feed_log("source_ok", symbol=sym, root=root, source="databento",
             spot=(result[0] if result else None),
             spot_feed=spot_source_name,
             spot_age_s=(int(spot_age) if spot_age is not None else None),
             rows=(len(result[1]) if result and len(result) > 1 else None))
    return result


async def load_symbol(sym):
    ensure_sym_state(sym)
    spot, rows, expiries, _src = await throttled(fetch_symbol_payload, sym)

    # Put-call parity from our own chain, as a fallback and a cross-check for
    # the external index quote.
    #
    # The external quote is a free endpoint and can go stale, rate limit or
    # disappear. The parity spot is computed from the very quotes the sheet
    # is showing, so it is always real time and always consistent with the
    # chain. When the external quote is missing or old, the parity spot keeps
    # the index level moving instead of freezing it.
    if is_index_symbol(sym) and rows:
        parity = _parity_spot(rows, spot)
        if parity and parity > 0:
            _parity_spot_cache[sym] = (parity, time.monotonic())
            ext_age = spot_source_for(sym)[1]
            external_is_stale = (
                spot is None
                or (ext_age is not None and ext_age > SPOT_STALE_SECONDS)
            )
            if spot is not None:
                # Hard sanity limit. If the parity level is nowhere near the
                # external one the parity is wrong, not the external quote,
                # and it must not be allowed near the ATM row.
                if abs(parity - spot) > abs(spot) * PARITY_SANITY_PCT:
                    feed_log("spot_parity_rejected", symbol=sym,
                             external=round(spot, 4), parity=round(parity, 4))
                    parity = None
            if parity and external_is_stale:
                feed_log("spot_from_parity", symbol=sym, parity=round(parity, 4),
                         external=spot, external_age_s=ext_age)
                spot = parity
            elif parity is not None and spot is not None:
                diff = abs(parity - spot)
                tol = max(PARITY_TOLERANCE_POINTS, abs(spot) * PARITY_TOLERANCE_PCT)
                if diff > tol:
                    _spot_divergence[sym] = diff
                    feed_log("spot_divergence", symbol=sym,
                             external=round(spot, 4), parity=round(parity, 4),
                             diff=round(diff, 4))
                else:
                    _spot_divergence.pop(sym, None)
                # Record the cross-check now and then, so it is always visible
                # which source is moving and how far apart they are, rather
                # than having to guess.
                last = _parity_log_at.get(sym, 0.0)
                if time.monotonic() - last > 30:
                    _parity_log_at[sym] = time.monotonic()
                    feed_log("spot_parity_check", symbol=sym,
                             external=round(spot, 4), parity=round(parity, 4),
                             diff=round(diff, 4),
                             external_age_s=ext_age)

    if spot:
        note_spot(sym, spot, _src)

    if expiries:
        all_known_expiries[sym] = expiries
    if active_expiries[sym] == [None, None, None]:
        # Only re-pick when we actually have expiry data. If this fetch came
        # back empty the previously assigned expiries stay in place and the
        # trader keeps seeing the last good chain, flagged stale, instead of
        # watching the block blank itself out.
        dates = available_expiries(all_known_expiries[sym], sym)
        if dates:
            active_expiries[sym] = (dates[:3] + [None, None, None])[:3]
        else:
            feed_log("expiry_none_available", symbol=sym,
                     note="no expiry list; keeping previous assignment")

    # full replace â€” drops stale foreign-root / expired keys after source switch
    fresh = {}
    for row in rows:
        key = row["contract"] or f"{sym}-{row['type']}-{row['strike']}-{row['expiry']}"
        fresh[key] = row
    live_data[sym] = fresh

    return {"options": rows, "n": len(rows)}


def available_expiries(all_expiries, sym=None, by_date=False):
    """Expiries to offer.

    The default slots prefer the standard (non-weekly) chain, because an index
    like SPX also lists a daily mini chain and the mini chain is thin, with
    poor open interest and wide quotes. Those expiries are still offered, just
    not in the first slots.

    The drop-down the customer picks from is ordered by date, nearest first.
    That is what a trader expects to see, and it is what their own broker
    shows: the chain they want is almost always the next expiry out, and with
    a standard-first list the near dailies were buried at the bottom.
    """
    now, eastern_available = eastern_now(warning=_log_timezone_fallback)
    dates = filter_expiry_dates(all_expiries, now, eastern_available)
    if sym is None:
        return dates
    if by_date:
        return sorted(dates)
    return standard_expiries(sym, dates)


def _log_timezone_fallback(exc):
    if us_state._timezone_warning_logged:
        return
    us_state._timezone_warning_logged = True
    try:
        append_runtime_warning(
            "US Eastern timezone unavailable; today's expiry remains selectable",
            f"{type(exc).__name__}: {exc}",
        )
    except OSError as log_error:
        print(f"Could not log timezone warning: {log_error}", flush=True)


async def init_symbol(sym):
    data = await load_symbol(sym)
    print(f"[{sym}] loaded {data.get('n', 0)} contracts | spot {spot_cmp[sym]}", flush=True)
    print(f"[{sym}] expiries: {', '.join(active_expiries[sym])}", flush=True)


async def symbol_poller(sym):
    """Rebuild one symbol's chain on its own period.

    Each symbol used to be rebuilt only once per shared round, so with three
    symbols every one of them was about six seconds behind. Each symbol now
    has its own task, and the work is short enough that they do not queue
    behind each other.
    """
    while True:
        started = time.monotonic()
        try:
            await load_symbol(sym)
            us_state.net_note(True)
        except asyncio.CancelledError:
            raise
        except Exception as e:
            print(f"Poller error ({sym}): {type(e).__name__}: {e}", flush=True)
            us_state.net_note(False)
        elapsed = time.monotonic() - started
        # A symbol the user just changed is rebuilt straight away.
        delay = 0.0 if us_state._refetch_now else SYMBOL_POLL_SECONDS
        if delay == 0.0:
            us_state._refetch_now = False
        # Sleep AFTER the work, for at least as long as the work took.
        await asyncio.sleep(max(0.0, delay - elapsed))

async def data_poller():
    """Keep one poller per symbol, starting and stopping as symbols change.

    A poller that dies is restarted. A task left in the map after it has
    finished is never replaced, so a single failure would silently freeze one
    symbol for the rest of the session.
    """
    tasks = {}
    try:
        while True:
            wanted = active_symbols()
            for sym in wanted:
                task = tasks.get(sym)
                if task is None or task.done():
                    if task is not None:
                        # Recover whatever it died on instead of losing it.
                        try:
                            exc = task.exception()
                        except asyncio.CancelledError:
                            exc = "CancelledError"
                        except Exception as e:
                            exc = f"{type(e).__name__}: {e}"
                        print(f"Poller for {sym} stopped ({exc}), restarting", flush=True)
                    tasks[sym] = asyncio.create_task(symbol_poller(sym))
            for sym in list(tasks):
                if sym not in wanted:
                    tasks.pop(sym).cancel()
            await asyncio.sleep(0.5)
    finally:
        for task in tasks.values():
            task.cancel()


async def spot_poller():
    """Keep the shared spot cache warm.

    Every source is refreshed once a second and each one runs on its own, so a
    single slow HTTP request cannot hold up the others. Waiting for all of
    them together meant one slow feed made the cache expire, and the index
    level on screen fell seconds behind the real price.
    """
    in_flight = set()
    while True:
        jobs = []
        for sym in active_symbols():
            root = clean_symbol(sym)
            if not root:
                continue
            key = ("yahoo", root)
            if key in in_flight:
                continue
            jobs.append((key, asyncio.to_thread(_yahoo_spot_fetch, root)))
        key = ("tencent", None)
        if key not in in_flight:
            jobs.append((key, asyncio.to_thread(refresh_tencent_spot_cache)))
        for k, coro in jobs:
            in_flight.add(k)
            task = asyncio.create_task(coro)
            task.add_done_callback(lambda _t, kk=k: in_flight.discard(kk))
        await asyncio.sleep(SPOT_POLL_SECONDS)


def _refresh_stock_spot(sym):
    """Refresh the cached Yahoo spot for a stock symbol."""
    root = clean_symbol(sym)
    return _yahoo_spot_fetch(root)


async def expiry_roller():
    while True:
        await asyncio.sleep(EXPIRY_REFRESH)
        for sym in active_symbols():
            try:
                ensure_sym_state(sym)
                dates = available_expiries(all_known_expiries[sym], sym)
                if not dates:
                    continue
                pinned = [active_expiries[sym][i] for i in range(3) if user_pinned[sym][i]]
                desired = [d for d in dates if d not in pinned][:3]
                di = 0
                changed = []
                for i in range(3):
                    if user_pinned[sym][i]:
                        continue
                    tgt = desired[di] if di < len(desired) else None
                    di += 1
                    cur = active_expiries[sym][i]
                    if tgt == cur:
                        continue
                    active_expiries[sym][i] = tgt
                    changed.append((i, tgt))
                    print(f"[{sym}] slot {i + 1}: {cur or '--'} -> {tgt or '--'}", flush=True)
                if changed:
                    # Rolling an expiry swaps the strike ladder, so the cached
                    # payload/shape must go or the new expiry's first write and
                    # recolour would be skipped as "unchanged".
                    _main._block_payload.clear()
                    _main._colour_state.clear()
                    # write back to every open book currently showing this symbol
                    try:
                        import xlwings as xw
                        for role in ROLE_ORDER:
                            if role_symbol[role] != sym:
                                continue
                            book = _main.find_open_book(xw, excel_path(role))
                            if not book:
                                continue
                            sheet = book.sheets[0]
                            for i, tgt in changed:
                                cell = sheet.range(["D3", "O3", "Z3"][i])
                                try:
                                    cell.api.NumberFormat = "@"
                                except Exception:
                                    pass
                                cell.value = iso_to_user(tgt) if tgt else ""
                                cell.color = (255, 255, 153)
                                cell.font.bold = True
                            _main.write_hint_labels(sheet, role)
                    except Exception as e:
                        print(f"Expiry write-back error ({sym}): {e}", flush=True)
            except Exception as e:
                print(f"Expiry roller error ({sym}): {e}", flush=True)
    import us_excel as _main



async def process_pending_dates():
    import us_excel as _main
    global _pending_date_loads
    while True:
        if _pending_date_loads:
            sym, i, parsed = _pending_date_loads.pop(0)
            try:
                if parsed not in all_known_expiries[sym]:
                    print(f"[{sym}] no contracts for {parsed}", flush=True)
                    continue
                active_expiries[sym][i] = parsed
                user_pinned[sym][i] = True
                # A different expiry means a different strike ladder, so drop the
                # cached payload/shape or the first write and the first recolour
                # of the new expiry would be suppressed.
                _main._block_payload.clear()
                _main._colour_state.clear()
                print(f"[{sym}] pinned block {i + 1} -> {parsed}", flush=True)
            except Exception as e:
                print(f"Date pin error ({sym} {parsed}): {e}", flush=True)
        else:
            await asyncio.sleep(0.5)


def build_option_chain(sym, expiry, range_val=50):
    sp = spot_cmp[sym]
    if sp is None or not expiry:
        return None, 0.0
    want = base_root(sym)
    filtered = prefer_expiry_root(live_data[sym].values(), expiry, want)
    if not filtered:
        return None, 0.0
    all_calls = [e for e in filtered if e["type"] == "call"]
    all_puts = [e for e in filtered if e["type"] == "put"]
    if not all_calls or not all_puts:
        return None, 0.0

    strikes = sorted(set(e["strike"] for e in filtered))
    atm_strike = min(strikes, key=lambda s: abs(s - sp))
    atm_idx = strikes.index(atm_strike)
    # Some index chains span a very wide strike range with a dense listing at
    # the money (DJX: ~42 strikes from 325..625 around a ~513 spot). Rendering
    # the whole span buries the liquid ATM region under far-OTM strikes whose OI
    # is legitimately 0 on every public feed, so the user only sees zeros.
    # Key off the span relative to spot, not the strike count: RUT lists 50
    # strikes across only ~9% of spot and its OI is fine, DJX spans ~60%.
    window = range_val
    if is_index_symbol(sym) and sp > 0:
        span_pct = (strikes[-1] - strikes[0]) / float(sp)
        if span_pct > 0.30 and len(strikes) < window * 2 + 1:
            window = max(5, min(15, (len(strikes) - 1) // 2))
    start = max(0, atm_idx - window)
    end = min(len(strikes), start + window * 2 + 1)
    if end - start < window * 2 + 1:
        start = max(0, end - window * 2 - 1)
    target = set(strikes[start:end])

    def _best_by_strike(rows):
        """Pick one contract per strike: standard root first, then quote quality.

        Same strike can exist as both DJX and DJXW (or SPX/SPXW). Quality alone
        would sometimes pick the weekly because its book looks tighter, so root
        match dominates and quality only orders rows of the same root.
        """
        picked = {}
        for e in rows:
            s = e["strike"]
            if s not in target:
                continue
            prev = picked.get(s)
            if prev is None:
                picked[s] = e
                continue
            is_std_e = 1 if e.get("root") == want else 0
            is_std_p = 1 if prev.get("root") == want else 0
            if is_std_e != is_std_p:
                if is_std_e > is_std_p:
                    picked[s] = e
            elif _quote_quality(e["bid"], e["ask"], e["oi"], e["ltp"]) > _quote_quality(
                prev["bid"], prev["ask"], prev["oi"], prev["ltp"]
            ):
                picked[s] = e
        return picked

    best_calls = _best_by_strike(all_calls)
    best_puts = _best_by_strike(all_puts)
    call_rows = [{
        "strike": e["strike"],
        "Call LTP": e["ltp"],
        "Call Bid": e["bid"],
        "Call Ask": e["ask"],
        "Call OI": e["oi"],
    } for e in best_calls.values()]
    put_rows = [{
        "strike": e["strike"],
        "Put OI": e["oi"],
        "Put Bid": e["bid"],
        "Put Ask": e["ask"],
        "Put LTP": e["ltp"],
    } for e in best_puts.values()]

    df_calls = pd.DataFrame(call_rows)
    df_puts = pd.DataFrame(put_rows)
    if df_calls.empty or df_puts.empty:
        return None, 0.0

    merged = pd.merge(df_calls, df_puts, on="strike", how="outer")
    if merged.empty:
        return None, 0.0
    merged = merged.sort_values("strike").reset_index(drop=True).fillna(0.0)
    merged["Call OI"] = merged["Call OI"].apply(fmt_oi)
    merged["Put OI"] = merged["Put OI"].apply(fmt_oi)

    def oi_total(series):
        total = 0.0
        for v in series:
            s = str(v).replace("$", "")
            mult = 1.0
            if s.endswith("M"):
                mult, s = 1_000_000, s[:-1]
            elif s.endswith("K"):
                mult, s = 1_000, s[:-1]
            try:
                total += float(s) * mult
            except ValueError:
                pass
        return total

    tc, tp = oi_total(merged["Call OI"]), oi_total(merged["Put OI"])
    pcr = tp / tc if tc > 0 else 0.0
    available = [c for c in CHAIN_COLS if c in merged.columns]
    return merged[available], pcr


def _short_build():
    """Just the sha, so the status line stays short enough for a title bar."""
    if us_state._SHORT_BUILD is None:
        raw = str(build_id())
        for token in raw.split():
            if token.startswith("build_sha="):
                us_state._SHORT_BUILD = token.split("=", 1)[1]
                break
        else:
            _SHORT_BUILD = "unstamped"
    return us_state._SHORT_BUILD


def _header_at(header_values, col):
    """Value of the header cell in column `col` from the cached header row."""
    if not header_values:
        return None
    index = col - 1
    if 0 <= index < len(header_values):
        return header_values[index]
    return None


def _set_header_at(header_values, col, value):
    """Keep the cached header row in step with what we just wrote."""
    if not header_values:
        return
    index = col - 1
    if 0 <= index < len(header_values):
        header_values[index] = value


def print_chain(sym):
    sp = spot_cmp[sym]
    for bi, expiry in enumerate(active_expiries[sym]):
        if not expiry:
            continue
        chain, pcr = build_option_chain(sym, expiry, 999)
        if chain is None or chain.empty:
            continue
        sentiment = "BULLISH" if pcr > 1.0 else "BEARISH"
        print(f"{sym} Block {bi + 1} | {expiry} | Spot: {sp:.2f} | PCR: {pcr:.2f} ({sentiment}) | {len(chain)} strikes")
        header = "  ".join(f"{c:>10}" for c in CHAIN_COLS)
        print(header)
        print("-" * len(header))
        for _, row in chain.iterrows():
            parts = []
            for c in CHAIN_COLS:
                v = row.get(c, 0)
                if c in ("Call OI", "Put OI", "strike"):
                    parts.append(str(v).rjust(10))
                else:
                    parts.append(f"{float(v):>10.2f}")
            print("  ".join(parts))
        print()
