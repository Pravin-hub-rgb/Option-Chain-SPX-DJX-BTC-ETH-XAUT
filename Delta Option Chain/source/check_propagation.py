"""End-to-end: does a market tick propagate into the trader's strategy sheet?

That is the product promise -- the trader writes formulas against the chain,
and every tick updates their P&L by itself. This samples the chain AND the
strategy sheet over time and reports what actually moved, and how quickly the
strategy caught up.
"""
import os
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import xlwings as xw

SRC = (r"C:\Users\Pravin\Downloads\Telegram Desktop\new2-deltax-option-chain"
       r"\deltax-option-chain\Delta Option Chain\source")

SAMPLE_CHAIN = [f"{c}{r}" for r in range(5, 30) for c in ("A", "D", "F", "I")][:120]
SAMPLE_STRAT = [f"{c}{r}" for r in range(7, 60) for c in ("C", "D", "G", "I", "J")][:120]
DURATION = int(sys.argv[1]) if len(sys.argv) > 1 else 90


def snap(sh, addrs):
    return {a: sh.range(a).value for a in addrs}


book = xw.Book(SRC + r"\btc_chain.xlsx")
app = book.app
app.visible = True
chain, sa, sb = book.sheets("BTC"), book.sheets("StratA"), book.sheets("StratB")

print(f"   before: spot header = {chain.range('A1').value}", flush=True)
p = subprocess.Popen([sys.executable, "main.py"], cwd=SRC, env=dict(os.environ),
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(40)                      # let it attach and write

t0 = time.time()
base_chain = snap(chain, SAMPLE_CHAIN)
base_sa = snap(sa, SAMPLE_STRAT)
base_sb = snap(sb, SAMPLE_STRAT)

chain_ticks, sa_ticks, sb_ticks = {}, {}, {}
maxlag_sa, maxlag_sb = [], []

while time.time() - t0 < DURATION:
    time.sleep(2)
    now_c = snap(chain, SAMPLE_CHAIN)
    for a, v in now_c.items():
        if v != base_chain[a]:
            chain_ticks.setdefault(a, time.time() - t0)
    now_a, now_b = snap(sa, SAMPLE_STRAT), snap(sb, SAMPLE_STRAT)
    for a, v in now_a.items():
        if v != base_sa[a]:
            sa_ticks.setdefault(a, time.time() - t0)
    for a, v in now_b.items():
        if v != base_sb[a]:
            sb_ticks.setdefault(a, time.time() - t0)

p.terminate()
final = snap(chain, SAMPLE_CHAIN)
header = chain.range("A1").value

print(f"\n   after:  spot header = {header}")
print(f"\n   chain cells that moved : {len(chain_ticks)}/{len(SAMPLE_CHAIN)}"
      f"   ({100*len(chain_ticks)//len(SAMPLE_CHAIN)}%)")
print(f"   StratA cells that moved: {len(sa_ticks)}/{len(SAMPLE_STRAT)}"
      f"   ({100*len(sa_ticks)//len(SAMPLE_STRAT)}%)")
print(f"   StratB cells that moved: {len(sb_ticks)}/{len(SAMPLE_STRAT)}"
      f"   ({100*len(sb_ticks)//len(SAMPLE_STRAT)}%)")

changed_now = sum(1 for a, v in final.items() if v != base_chain[a])
print(f"   chain cells still different at the end: {changed_now}")

# did strategy follow the chain, and after how long?
common = set(chain_ticks) & set(sa_ticks)
if chain_ticks:
    lags = [sa_ticks[c] - chain_ticks[c] for c in common if c in sa_ticks]
    if lags:
        lags.sort()
        print(f"\n   strategy lag behind chain on {len(lags)} shared rows:")
        print(f"      min {min(lags):.1f}s   median {lags[len(lags)//2]:.1f}s   max {max(lags):.1f}s")
    else:
        print("\n   chain moved but NO matching strategy cell moved within the window")
book.close()