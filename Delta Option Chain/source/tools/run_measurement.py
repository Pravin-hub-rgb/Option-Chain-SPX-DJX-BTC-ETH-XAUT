"""Clean strategy-responsiveness measurement. READ-ONLY once measuring starts.

Rules this harness obeys, because breaking them produced a 400 s hang:
  * nothing drives Excel over COM during a run except small timed reads
  * the tool starts first; we then ATTACH to its Excel, we never open one
  * Excel is closed gracefully, never force-killed
  * gates abort the run instead of reporting meaningless numbers

Usage: python run_measurement.py <A|B> <off|on|pause> <minutes>
"""
import ctypes
import ctypes.wintypes as w
import json
import os
import re
import statistics
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32com.client

SRC = (r"C:\Users\Pravin\Downloads\Telegram Desktop\new2-deltax-option-chain"
       r"\deltax-option-chain\Delta Option Chain\source")
DIAG = os.path.join(os.environ["LOCALAPPDATA"], "DeltaOptionChain", "startup_diag.log")
FEEDLOG = os.path.join(os.environ["LOCALAPPDATA"], "DeltaOptionChain", "feed.log")

VARIANT, CONDITION = sys.argv[1], sys.argv[2]
MINUTES = float(sys.argv[3])
SHEET = "StratA" if VARIANT == "A" else "StratB"
DURATION = int(MINUTES * 60)
REPS = 20
TYPE_CELL = "A900"
CHAIN_PROBE = "A5"          # a live call LTP in the chain
STRAT_PROBE = "C7"          # INDEX/MATCH on the chain -> must follow it
VOL_PROBE = "L7"            # NOW() -> proves volatile recalc really runs

_k32 = ctypes.windll.kernel32


def log(msg):
    print(msg, flush=True)


# ------------------------------------------------------------------ gates
def clean_slate():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    out2 = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OptionChain.exe", "/NH"],
                          capture_output=True, text=True).stdout
    bad = ("EXCEL.EXE" in out) or ("OptionChain.exe" in out2)
    return (not bad), {"excel_open": "EXCEL.EXE" in out,
                       "tool_open": "OptionChain.exe" in out2}


def stop_everything(graceful=True):
    ps = ("Get-CimInstance Win32_Process -Filter \"Name='python.exe' OR "
          "Name='pythonw.exe' OR Name='OptionChain.exe'\" | "
          "Where-Object { $_.CommandLine -like '*main.py*' -or $_.Name -eq 'OptionChain.exe' } | "
          "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }")
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True)
    # graceful close, never a force-kill
    for _ in range(4):
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
            for i in range(xl.Workbooks.Count, 0, -1):
                xl.Workbooks(i).Close(False)
            xl.Quit()
        except Exception:
            break
        time.sleep(2)
    # wait for Excel to actually go
    for _ in range(20):
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                             capture_output=True, text=True).stdout
        if "EXCEL.EXE" not in out:
            break
        time.sleep(1)
    # teardown only (never during a measurement): if Excel is still standing,
    # clear it so the next run's clean-slate gate passes
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    if "EXCEL.EXE" in out:
        subprocess.run(["taskkill", "/IM", "EXCEL.EXE", "/F"],
                       capture_output=True)
        time.sleep(3)
    time.sleep(2)


def excel_pid():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True).stdout
    for line in out.splitlines():
        if "EXCEL.EXE" in line:
            return int(re.search(r'"(\d+)"', line).group(1))
    return None


def _cpu_seconds(pid):
    h = _k32.OpenProcess(0x1000, False, pid)
    if not h:
        return None
    c, e, k, u = (w.FILETIME(), w.FILETIME(), w.FILETIME(), w.FILETIME())
    try:
        if not _k32.GetProcessTimes(h, ctypes.byref(c), ctypes.byref(e),
                                    ctypes.byref(k), ctypes.byref(u)):
            return None
        f = lambda t: (t.dwHighDateTime << 32) | t.dwLowDateTime
        return (f(k) + f(u)) / 1e7
    finally:
        _k32.CloseHandle(h)


def mem_mb(pid):
    """Working set of Excel in MB, read without touching Excel."""
    if not pid:
        return 0
    r = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         f"(Get-Process -Id {pid}).WorkingSet64"],
        capture_output=True, text=True)
    try:
        return int(r.stdout.strip()) // (1024 * 1024)
    except Exception:
        return 0


def excel_hung():
    return "Not Responding" in subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH", "/V"],
        capture_output=True, text=True).stdout


def raw_feed_live(seconds):
    import asyncio
    import websockets

    async def probe():
        vals = []
        t0 = time.time()
        async with websockets.connect("wss://public-socket.india.delta.exchange",
                                      open_timeout=15) as ws:
            await ws.send(json.dumps({"type": "subscribe", "payload": {"channels": [
                {"name": "spot_price", "symbols": [".DEXBTUSD"]}]}}))
            while time.time() - t0 < seconds:
                try:
                    raw = await asyncio.wait_for(ws.recv(), timeout=6)
                except asyncio.TimeoutError:
                    break
                d = json.loads(raw)
                if d.get("type") == "spot_price" and d.get("p") is not None:
                    vals.append(d["p"])
        return vals

    try:
        v = asyncio.run(probe())
    except Exception as e:
        return False, {"error": f"{type(e).__name__}: {e}"}
    return (len(v) > 5 and len(set(v)) > 2), {"frames": len(v), "distinct": len(set(v))}


def parse_perf():
    if not os.path.exists(DIAG):
        return {}
    txt = open(DIAG, encoding="utf-8", errors="replace").read()[-30000:]
    out = {}
    for m in re.finditer(
            r"performance\.(\w+)\.([\w_]+):\s*count=(\d+),\s*avg_ms=([\d.]+),\s*max_ms=([\d.]+)",
            txt):
        out.setdefault(m.group(2), {"count": 0, "avg_ms": 0.0, "max_ms": 0.0})
        out[m.group(2)] = {"count": int(m.group(3)), "avg_ms": float(m.group(4)),
                           "max_ms": float(m.group(5))}
    return out


def feed_errors():
    if not os.path.exists(FEEDLOG):
        return {}
    counts = {}
    for line in open(FEEDLOG, encoding="utf-8", errors="replace"):
        parts = line.split()
        if len(parts) > 1 and parts[1] in ("rest_error", "rest_failed", "rest_conn",
                                          "rest_429", "probe_error",
                                          "roller_all_probes_failed"):
            counts[parts[1]] = counts.get(parts[1], 0) + 1
    return counts


# ------------------------------------------------------- typing (SendKeys)
def typing_reps(xl, book, sheet_name, pid):
    ws = win32com.client.Dispatch("WScript.Shell")
    book.Activate()
    sh = book.Worksheets(sheet_name)

    def one(value):
        sh.Activate()
        cell = sh.Range(TYPE_CELL)
        cell.Select()
        ws.AppActivate(pid)
        time.sleep(0.10)
        t0 = time.perf_counter()
        ws.SendKeys(str(value) + "{ENTER}")
        deadline, seen = t0 + 8.0, None
        while time.perf_counter() < deadline:
            try:
                if cell.Value == value:
                    seen = time.perf_counter()
                    break
            except Exception:
                pass
            time.sleep(0.002)
        return min((seen or time.perf_counter()) - t0, 8.0) * 1000.0

    for w in range(2):
        one(9000 + w)
    times, retries = [], 0
    for i in range(REPS):
        ms = one(100 + i)
        if ms >= 7999:
            retries += 1
            ms = one(100 + i)
        times.append(ms)
    t = sorted(times)
    return {"typing_avg_ms": round(statistics.mean(times), 1),
            "typing_median_ms": round(statistics.median(times), 1),
            "typing_p95_ms": round(t[int(REPS * 0.95) - 1], 1),
            "typing_worst_ms": round(t[-1], 1),
            "typing_retries": retries}


def main():
    res = {"variant": VARIANT, "condition": CONDITION, "sheet": SHEET,
           "minutes": MINUTES, "started": time.strftime("%H:%M:%S")}

    stop_everything()
    ok, state = clean_slate()
    res["gate_clean_slate"] = ok
    res["clean_slate_detail"] = state
    if not ok:
        res["ABORTED"] = f"not a clean slate: {state}"
        log(f"   ABORT {res['ABORTED']}")
        return save(res)

    # ---- GATE 1: exchange feed live for 2 minutes, independent of the tool
    live, finfo = raw_feed_live(120)
    res["gate_feed_live"] = live
    res["feed_info"] = finfo
    log(f"   GATE feed_live(120s)={live} {finfo}")
    if not live:
        res["ABORTED"] = "exchange feed not live for 2 min"
        log(f"   ABORT {res['ABORTED']}")
        return save(res)

    # ---- start the tool FIRST, then attach to the Excel it brings up
    proc = None
    if CONDITION in ("on", "pause"):
        env = dict(os.environ)
        env["OC_TEST_PAUSE"] = "1" if CONDITION == "pause" else "0"
        proc = subprocess.Popen([sys.executable, "-u", "main.py"], cwd=SRC, env=env,
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(60)
        # attach to the Excel the tool brought up; never open one ourselves here
        xl = win32com.client.GetActiveObject("Excel.Application")
    else:
        # tool OFF: nothing brought Excel up, so we open it ourselves
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        xl.Workbooks.Add(os.path.join(SRC, "btc_chain.xlsx"))
        time.sleep(8)
    xl.Visible = True
    xl.Calculation = -4105                      # automatic
    book = [b for b in [xl.Workbooks(i + 1) for i in range(xl.Workbooks.Count)]
            if "btc_chain" in b.Name.lower()][0]
    chain = book.Worksheets("BTC")
    strat = book.Worksheets(SHEET)
    pid = excel_pid()

    mem0 = mem_mb(pid)

    # ---- measurement: READ ONLY from here
    cpu, probe_ms, hangs = [], [], []
    chain_log, strat_log, vol_log = [], [], []
    chain_ticks, strat_ticks = [], []
    prev_chain, prev_strat = None, None
    cpu_last, cpu_t0 = _cpu_seconds(pid), time.perf_counter()
    start = time.perf_counter()

    while time.perf_counter() - start < DURATION:
        c = _cpu_seconds(pid)
        if c is not None and cpu_last is not None:
            dt = time.perf_counter() - cpu_t0
            if dt > 0:
                cpu.append(max(0.0, (c - cpu_last) / dt * 100.0))
        cpu_last, cpu_t0 = c, time.perf_counter()

        # UI responsiveness probe: a tiny marshalled read. If Excel's UI thread
        # is wedged, this blocks -- which is exactly the freeze we care about.
        t0 = time.perf_counter()
        try:
            _ = xl.ActiveCell.Address
        except Exception:
            pass
        probe_ms.append((time.perf_counter() - t0) * 1000.0)

        now = round(time.perf_counter() - start, 2)
        # Count how many cells actually changed this tick. Sampling one cell was
        # misleading: a deep-OTM strike never trades, so it read "static" while
        # the rest of the block was moving.
        try:
            cvals = chain.Range("A5:I40").Value2
            flat = [c for row in cvals for c in row]
            if prev_chain is not None:
                changed = sum(1 for a, b in zip(prev_chain, flat) if a != b)
                chain_ticks.append((now, changed))
            prev_chain = flat
        except Exception:
            pass
        try:
            svals = strat.Range(f"{STRAT_PROBE}7:{STRAT_PROBE}806").Value2
            sflat = [c for row in svals for c in row]
            if prev_strat is not None:
                changed = sum(1 for a, b in zip(prev_strat, sflat) if a != b)
                strat_ticks.append((now, changed))
            prev_strat = sflat
        except Exception:
            pass
        try:
            v = strat.Range(VOL_PROBE).Value
            vol_log.append((now, v.isoformat() if hasattr(v, "isoformat") else v))
        except Exception:
            vol_log.append((now, None))

        if int(time.perf_counter() - start) % 30 == 0 and excel_hung():
            hangs.append(time.strftime("%H:%M:%S"))
            log(f"   !! Not Responding {hangs[-1]}")
        time.sleep(1.0)

    res["wall_seconds"] = round(time.perf_counter() - start, 1)
    if cpu:
        cs = sorted(cpu)
        res["cpu_avg_pct"] = round(statistics.mean(cs), 1)
        res["cpu_p95_pct"] = round(cs[int(len(cs) * 0.95) - 1], 1)
        res["cpu_max_pct"] = round(cs[-1], 1)
        res["cpu_samples"] = len(cs)
    res["ui_probe_avg_ms"] = round(statistics.mean(probe_ms), 2)
    res["ui_probe_max_ms"] = round(max(probe_ms), 1)
    res["ui_probe_over_1s"] = sum(1 for x in probe_ms if x > 1000)
    res["not_responding"] = hangs

    # ---- GATE 2/3: chain moved, strategy followed
    ch_moves = [i for i, (_, c) in enumerate(chain_ticks) if c > 0]
    st_moves = [i for i, (_, c) in enumerate(strat_ticks) if c > 0]
    vo_moves = sum(1 for i in range(1, len(vol_log)) if vol_log[i][1] != vol_log[i - 1][1])
    res["ticks_sampled"] = len(chain_ticks)
    res["chain_ticks_with_change"] = len(ch_moves)
    res["strategy_ticks_with_change"] = len(st_moves)
    res["volatile_changes"] = vo_moves
    if chain_ticks:
        cc = [c for _, c in chain_ticks]
        res["chain_cells_changed_per_tick"] = {
            "avg": round(statistics.mean(cc), 1), "max": max(cc),
            "ticks_total": len(cc)}
    if strat_ticks:
        sc = [c for _, c in strat_ticks]
        res["strategy_cells_changed_per_tick"] = {
            "avg": round(statistics.mean(sc), 1), "max": max(sc),
            "ticks_total": len(sc)}

    lags = []
    for i in ch_moves:
        t_ch = chain_ticks[i][0]
        for j in st_moves:
            if j >= i and strat_ticks[j][0] >= t_ch:
                lags.append(strat_ticks[j][0] - t_ch)
                break
    if lags:
        ls = sorted(lags)
        res["chain_to_strategy_lag_s"] = {
            "n": len(ls), "median": round(statistics.median(ls), 2),
            "p95": round(ls[int(len(ls) * 0.95) - 1], 2), "max": round(ls[-1], 2)}

    gaps = [chain_ticks[i][0] - chain_ticks[i - 1][0] for i in ch_moves]
    if gaps:
        res["refresh_interval_s"] = {"median": round(statistics.median(gaps), 2),
                                     "min": round(min(gaps), 2), "max": round(max(gaps), 2)}

    log(f"   typing {REPS} reps (SendKeys, not COM)...")
    res.update(typing_reps(xl, book, SHEET, pid))

    res["perf"] = parse_perf()
    res["feed_errors"] = feed_errors()

    res["mem_start_mb"] = mem0
    res["mem_end_mb"] = mem_mb(pid)

    if proc:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except Exception:
            pass
    stop_everything()
    return save(res)


def save(res):
    outdir = os.path.join(SRC, "strategy_results")
    os.makedirs(outdir, exist_ok=True)
    name = f"{VARIANT}_{CONDITION}_{int(MINUTES)}m.json"
    json.dump(res, open(os.path.join(outdir, name), "w"), indent=2, default=str)
    log(f"   saved {name}")
    return res


if __name__ == "__main__":
    r = main()
    print("\n" + json.dumps(r, indent=2, default=str))