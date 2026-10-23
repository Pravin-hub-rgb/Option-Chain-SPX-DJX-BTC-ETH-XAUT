"""One-shot: merge startup_diagnostics into delta_diagnostics and rename
performance_metrics to delta_performance, so the two products have the same
module names.

  startup_diagnostics.py + the feed log  ->  delta_diagnostics.py
  performance_metrics.py                 ->  delta_performance.py

Run from the Delta Option Chain folder.
"""

import io
import os
import re

SRC = "source/startup_diagnostics.py"
PERF = "source/performance_metrics.py"
DIAG = "source/delta_diagnostics.py"

# ---- delta_diagnostics.py = the startup diagnostics + the feed log --------
startup = io.open(SRC, encoding="utf-8").read()
feed = io.open(DIAG, encoding="utf-8").read()

# The feed log half starts at its own docstring. Keep the startup half's
# imports and add the ones the feed log needs.
feed_body = feed[feed.index("import os"):]

merged = startup.rstrip() + "\n\n\n" + "# " + "-" * 75 + "\n"
merged += "# The feed log: one structured line per event, capped and never raising.\n"
merged += "#\n"
merged += "# The feed log is what support asks the customer to send, so it has to be\n"
merged += "# complete and it has to stay quiet when there is nothing to say. The\n"
merged += "# directory honours DELTA_OPTION_CHAIN_LOG_DIR so tests and support runs\n"
merged += "# never write into the real customer log.\n"
merged += "# " + "-" * 75 + "\n\n"
merged += feed_body.lstrip("\n")

io.open(DIAG, "w", encoding="utf-8", newline="\n").write(merged)
os.remove(SRC)
print("merged startup_diagnostics.py into delta_diagnostics.py")

# ---- performance_metrics.py -> delta_performance.py ------------------------
perf = io.open(PERF, encoding="utf-8").read()
perf = perf.replace(
    '"""',
    '"""Performance metrics for the Delta tool.\n\nMirrors the US tool\'s us_performance.py, so the two products report the\nsame numbers under the same names.\n\n',
    1,
)
io.open("source/delta_performance.py", "w", encoding="utf-8", newline="\n").write(perf)
os.remove(PERF)
print("renamed performance_metrics.py to delta_performance.py")

# ---- update the imports ---------------------------------------------------
for path in ("source/main.py", "source/delta_state.py", "source/test_performance_metrics.py"):
    text = io.open(path, encoding="utf-8").read()
    text = text.replace(
        "from startup_diagnostics import (",
        "from delta_diagnostics import (",
    )
    text = text.replace(
        "from performance_metrics import PerformanceMetrics",
        "from delta_performance import PerformanceMetrics",
    )
    io.open(path, "w", encoding="utf-8", newline="\n").write(text)
    print("updated imports in", path)
