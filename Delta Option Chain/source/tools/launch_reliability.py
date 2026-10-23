"""Launch the tool 10 times and count how many show a live spot price.

This is the product-bug check: a customer who sees "BTC Spot: --" thinks the
tool is broken. If any launch fails to attach the index feed, that is ours to
fix, not just a quirk of this machine.
"""
import os
import re
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import xlwings as xw

SRC = (r"C:\Users\Pravin\Downloads\Telegram Desktop\new2-deltax-option-chain"
       r"\deltax-option-chain\Delta Option Chain\source")
TMP = os.environ["TEMP"]
LAUNCHES = int(sys.argv[1]) if len(sys.argv) > 1 else 10
WAIT = int(sys.argv[2]) if len(sys.argv) > 2 else 45

results = []
for n in range(1, LAUNCHES + 1):
    # graceful teardown so Excel never raises the "serious error" dialog
    try:
        for app in xw.apps:
            try:
                for b in list(app.books):
                    b.close(save_changes=False)
                app.quit()
            except Exception:
                pass
    except Exception:
        pass
    time.sleep(5)

    log = os.path.join(TMP, f"launch{n}.log")
    with open(log, "w", encoding="utf-8", errors="replace") as fh:
        p = subprocess.Popen([sys.executable, "-u", "main.py"], cwd=SRC,
                             stdout=fh, stderr=subprocess.STDOUT, env=dict(os.environ))
    time.sleep(WAIT)

    header, strikes = None, None
    try:
        s = xw.Book(SRC + r"\btc_chain.xlsx").sheets[0]
        header = s.range("A1").value or ""
        m = re.search(r"Spot:\s*([-\d.]+)", header)
        strikes = re.search(r"(\d+)\s+strikes", header)
    except Exception as e:
        header = f"<could not read: {e}>"

    p.terminate()
    try:
        p.wait(timeout=10)
    except Exception:
        p.kill()

    txt = ""
    try:
        txt = open(log, encoding="utf-8", errors="replace").read()
    except Exception:
        pass

    spot = m.group(1) if (m := re.search(r"Spot:\s*([-\d.]+)", header or "")) else None
    live = bool(spot) and spot not in ("--", "-", "")
    errs = [l for l in txt.splitlines() if "error" in l.lower()][:2]
    results.append({
        "launch": n, "live_spot": live, "spot": spot,
        "strikes": int(strikes.group(1)) if strikes else None,
        "stdout_lines": len([l for l in txt.splitlines() if l.strip()]),
        "errors": errs,
    })
    print(f"   launch {n:2d}: live={live!s:<5} spot={spot!s:<12} "
          f"strikes={results[-1]['strikes']!s:<5} stdout_lines={results[-1]['stdout_lines']}"
          f"  {errs if errs else ''}", flush=True)

good = sum(1 for r in results if r["live_spot"])
print(f"\n   LIVE SPOT: {good}/{len(results)} launches")
print(f"   FAILED   : {[r['launch'] for r in results if not r['live_spot']]}")
print("\n   --- a failing launch's log, if any ---")
for r in results:
    if not r["live_spot"]:
        txt = open(os.path.join(TMP, f"launch{r['launch']}.log"),
                   encoding="utf-8", errors="replace").read()
        print(f"   launch {r['launch']}: {len(txt)} chars")
        for l in [x for x in txt.splitlines() if x.strip()][:12]:
            print("      ", l[:150])
        break