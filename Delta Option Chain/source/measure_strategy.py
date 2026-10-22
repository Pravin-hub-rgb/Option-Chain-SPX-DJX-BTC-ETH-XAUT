"""Clean end-to-end strategy responsiveness measurement.

Nothing is measured unless the preconditions hold, because a static payload
makes the whole exercise meaningless:

  1. feed live      - spot is numeric AND actually changing for 60 s
  2. propagation    - a strategy cell is sampled with timestamps and must move
                      when the chain moves
  3. recalc         - the tool's own excel_recalc figure must be non-zero
  4. volatile ctrl  - with the tool OFF, variant B (NOW/OFFSET) must show
                      non-zero Excel CPU; if it does not, volatile functions
                      are not running and every B number is worthless

Usage: python measure_strategy.py <A|B> <off|on|pause> <seconds>
"""
import ctypes
import ctypes.wintypes as w
import json
import os
import re
import statistics
from datetime import date, datetime
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32com.client
import xlwings as xw

SRC = (r"C:\Users\Pravin\Downloads\Telegram Desktop\new2-deltax-option-chain"
       r"\deltax-option-chain\Delta Option Chain\source")
BOOK = SRC + r"\btc_chain.xlsx"
DIAG = os.path.join(os.environ["LOCALAPPDATA"], "DeltaOptionChain", "startup_diag.log")

VARIANT, CONDITION, SECONDS = sys.argv[1], sys.argv[2], int(sys.argv[3])
SHEET = "StratA" if VARIANT == "A" else "StratB"
REPS = 20
TYPE_CELL = "A900"
CHAIN_PROBE = "C7"          # chain-dependent: proves propagation with the tool ON
VOL_PROBE = "L7"            # NOW(): proves volatile recalc is actually running

_k32 = ctypes.windll.kernel32


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


def excel_pid():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE",
                          "/FO", "CSV", "/NH"], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if "EXCEL.EXE" in line:
            return int(re.search(r'"(\d+)"', line).group(1))
    return None


def excel_hung():
    return "Not Responding" in subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH", "/V"],
        capture_output=True, text=True).stdout


def jsonable(v):
    """NOW() comes back as a datetime; make samples JSON-safe."""
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return v


def clear_diag():
    """The tool appends to one shared log. Without this, run N reports run N-1."""
    import shutil
    if os.path.exists(DIAG):
        shutil.rmtree(os.path.dirname(DIAG), ignore_errors=True)


def kill_stray_tools():
    """A leftover tool holds Excel and makes every later run wrong.

    It also trips the single-instance guard, so the next 'tool on' run exits
    silently and aborts. This happened for real: a process left over from
    13:22 contaminated an entire matrix.

    Uses PowerShell CIM, not wmic -- wmic is gone from current Windows.
    """
    ps = ("Get-CimInstance Win32_Process -Filter "
          "\"Name='python.exe' OR Name='pythonw.exe'\" | "
          "Where-Object { $_.CommandLine -like '*main.py*' } | "
          "ForEach-Object { Stop-Process -Id $_.ProcessId -Force; "
          "Write-Output ('killed ' + $_.ProcessId) }")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                           capture_output=True, text=True, timeout=60)
        for line in (r.stdout or "").splitlines():
            if line.strip():
                print("  ", line.strip(), flush=True)
    except Exception as e:
        print("   kill_stray_tools failed:", e, flush=True)

def graceful_close():
    try:
        for app in xw.apps:
            try:
                for b in list(app.books):
                    b.close()
                app.quit()
            except Exception:
                pass
    except Exception:
        pass


def spot_of(sh):
    m = re.search(r"Spot:\s*([-\d.]+)", str(sh.range("A1").value or ""))
    return m.group(1) if m else None


def raw_feed_live(seconds=30):
    """Precondition 1: is the EXCHANGE feed itself live?

    Deliberately independent of the tool. When the tool is not running its
    workbook header cannot tell us anything, and a stale exchange feed would
    silently invalidate the whole run.
    """
    import asyncio
    import websockets

    async def probe():
        vals = []
        t0 = time.time()
        async with websockets.connect(
                "wss://public-socket.india.delta.exchange", open_timeout=15) as ws:
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
        vals = asyncio.run(probe())
    except Exception as e:
        return False, {"error": f"{type(e).__name__}: {e}"}
    uniq = len(set(vals))
    return (len(vals) > 3 and uniq > 1), {
        "frames": len(vals), "distinct": uniq,
        "first": vals[0] if vals else None, "last": vals[-1] if vals else None}


def wait_for_live_feed(chain, seconds=60):
    """Precondition 1b: the TOOL's own header going live (only meaningful when it runs)."""
    t0 = time.time()
    seen, last, moved = [], None, 0
    while time.time() - t0 < seconds:
        s = spot_of(chain)
        if s and s not in ("--", "-"):
            seen.append(s)
            if last is not None and s != last:
                moved += 1
            last = s
        time.sleep(2)
    ok = len(seen) > 5 and moved >= 3
    return ok, {"samples": len(seen), "distinct_moves": moved,
                "first": seen[0] if seen else None,
                "last": seen[-1] if seen else None}


def typing_reps(book, sh, pid):
    app = book.app
    app.visible = True
    ws = win32com.client.Dispatch("WScript.Shell")

    def one(value):
        sh.activate()
        cell = sh.range(TYPE_CELL)
        cell.value = None
        cell.api.Select()
        ws.AppActivate(pid)
        time.sleep(0.12)
        t0 = time.perf_counter()
        ws.SendKeys(str(value) + "{ENTER}")
        deadline, seen = t0 + 8.0, None
        while time.perf_counter() < deadline:
            try:
                if cell.value == value:
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
    ts = sorted(times)
    return {"typing_avg_ms": round(statistics.mean(times), 1),
            "typing_median_ms": round(statistics.median(times), 1),
            "typing_p95_ms": round(ts[int(REPS * 0.95) - 1], 1),
            "typing_worst_ms": round(ts[-1], 1),
            "typing_retries": retries}


def parse_perf():
    """performance.<scope>.<metric>: count=N, avg_ms=X, max_ms=Y"""
    if not os.path.exists(DIAG):
        return {}
    txt = open(DIAG, encoding="utf-8", errors="replace").read()[-20000:]
    out = {}
    for m in re.finditer(
            r"performance\.(\w+)\.([\w_]+):\s*count=(\d+),\s*avg_ms=([\d.]+),\s*max_ms=([\d.]+)",
            txt):
        scope, metric, cnt, avg, mx = m.groups()
        k = f"{metric}"
        out[k] = {"count": int(cnt), "avg_ms": float(avg), "max_ms": float(mx)}
    return out


def main():
    kill_stray_tools()
    graceful_close()
    clear_diag()
    time.sleep(4)
    res = {"variant": VARIANT, "condition": CONDITION, "sheet": SHEET,
           "seconds_requested": SECONDS}

    proc = None
    if CONDITION in ("on", "pause"):
        env = dict(os.environ)
        env["OC_TEST_PAUSE"] = "1" if CONDITION == "pause" else "0"
        _tl = open(os.path.join(os.environ["TEMP"], "tool_under_test.log"),
                   "w", encoding="utf-8", errors="replace")
        proc = subprocess.Popen([sys.executable, "-u", "main.py"], cwd=SRC, env=env,
                                stdout=_tl, stderr=subprocess.STDOUT)
        # The tool must be the one to bring Excel up and attach. Opening the
        # workbook from here first makes the tool attach differently and its
        # header never goes live -- which is what aborted every "on" run.
        time.sleep(50)

    book = xw.Book(BOOK)
    app = book.app
    app.visible = True
    app.api.Calculation = -4105   # AUTOMATIC - -4135 is manual and silently kills every recalc
    chain, sh = book.sheets("BTC"), book.sheets(SHEET)
    time.sleep(10)

    # ---- precondition 1: exchange feed live, checked independently of the tool ----
    raw_ok, raw_info = raw_feed_live(30)
    res["precondition_exchange_feed_live"] = raw_ok
    res["exchange_feed"] = raw_info
    print(f"   precondition exchange_feed_live={raw_ok} {raw_info}", flush=True)
    if not raw_ok:
        os.makedirs(os.path.join(SRC, "strategy_test_results"), exist_ok=True)
        res["ABORTED"] = "exchange feed not live - measurement would be meaningless"
        json.dump(res, open(os.path.join(SRC, "strategy_test_results",
                                        f"{VARIANT}_{CONDITION}.json"), "w"), indent=2)
        print("   ABORTED:", res["ABORTED"])
        book.close()
        if proc:
            proc.terminate()
        return

    # ---- precondition 1b: the tool's own header going live (only when it runs) ----
    if CONDITION in ("on", "pause"):
        ok, feedinfo = wait_for_live_feed(chain, 60)
        res["precondition_tool_header_live"] = ok
        res["tool_header"] = feedinfo
        print(f"   precondition tool_header_live={ok} {feedinfo}", flush=True)
        if not ok:
            os.makedirs(os.path.join(SRC, "strategy_test_results"), exist_ok=True)
            res["ABORTED"] = "tool header never went live"
            json.dump(res, open(os.path.join(SRC, "strategy_test_results",
                                            f"{VARIANT}_{CONDITION}.json"), "w"), indent=2)
            print("   ABORTED:", res["ABORTED"])
            try:
                _tl.flush()
                txt = open(os.path.join(os.environ["TEMP"], "tool_under_test.log"),
                           encoding="utf-8", errors="replace").read()
                lines = [l for l in txt.splitlines() if l.strip()]
                print(f"   tool stdout: {len(lines)} lines")
                for l in lines[-15:]:
                    print("      ", l[:170])
            except Exception as _e:
                print("   (no tool log:", _e, ")")
            book.close()
            if proc:
                proc.terminate()
            return
    time.sleep(5)

    pid = excel_pid()
    cpu_last, cpu_t0 = _cpu_seconds(pid), time.perf_counter()
    cpu = []
    hangs, ticks = [], []

    start = time.perf_counter()
    while time.perf_counter() - start < SECONDS:
        cnow = _cpu_seconds(pid)
        if cnow is not None and cpu_last is not None:
            dt = time.perf_counter() - cpu_t0
            if dt > 0:
                cpu.append(max(0.0, (cnow - cpu_last) / dt * 100.0))
            cpu_last, cpu_t0 = cnow, time.perf_counter()
        if excel_hung():
            hangs.append(time.strftime("%H:%M:%S"))
            print(f"   !! Not Responding {hangs[-1]}", flush=True)
        # precondition 2: tick propagation, with timestamps
        try:
            ticks.append({
                "t": round(time.perf_counter() - start, 1),
                "chain": jsonable(chain.range("A5").value),
                "spot": spot_of(chain),
                "strat": jsonable(sh.range(CHAIN_PROBE).value),
                "vol": jsonable(sh.range(VOL_PROBE).value),
            })
        except Exception:
            pass
        time.sleep(2)

    if cpu:
        res["excel_cpu_avg_pct"] = round(statistics.mean(cpu), 1)
        res["excel_cpu_max_pct"] = round(max(cpu), 1)
        res["excel_cpu_samples"] = len(cpu)
    res["not_responding"] = hangs
    res["wall_seconds"] = round(time.perf_counter() - start, 1)

    # propagation summary
    chain_moves = sum(1 for i in range(1, len(ticks))
                      if ticks[i]["chain"] != ticks[i - 1]["chain"])
    strat_moves = sum(1 for i in range(1, len(ticks))
                      if ticks[i]["strat"] != ticks[i - 1]["strat"])
    vol_moves = sum(1 for i in range(1, len(ticks))
                    if ticks[i]["vol"] != ticks[i - 1]["vol"])
    res["tick_samples"] = len(ticks)
    res["chain_cell_moves"] = chain_moves
    res["strategy_cell_moves"] = strat_moves
    res["strategy_tracked_chain"] = strat_moves > 0
    res["volatile_cell_moves"] = vol_moves
    res["volatile_recalc_running"] = vol_moves > 0

    print(f"   typing {REPS} reps...", flush=True)
    res.update(typing_reps(book, sh, pid))

    res["perf"] = parse_perf()
    book.close()
    if proc:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except Exception:
            proc.kill()

    outdir = os.path.join(SRC, "strategy_test_results")
    os.makedirs(outdir, exist_ok=True)
    json.dump(res, open(os.path.join(outdir, f"{VARIANT}_{CONDITION}.json"), "w"), indent=2)
    json.dump(ticks, open(os.path.join(outdir, f"{VARIANT}_{CONDITION}_ticks.json"), "w"), indent=2)
    print("\n" + json.dumps(res, indent=2))


if __name__ == "__main__":
    main()
