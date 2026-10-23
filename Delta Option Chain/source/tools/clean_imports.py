"""Remove the imports the module split left behind in main.py."""

import io
import re

PATH = "source/main.py"

# name -> the import statement it belongs to, so the right block is edited.
DROP = {
    "from delta_diagnostics import (": ["FEED_LOG_DIR", "FEED_LOG_MAX_BYTES", "FEED_LOG_PATH"],
    "from delta_licensing import (": [
        "LICENSE_PATH", "_license_sig", "_verify_sig", "ask_product_key",
        "key_accepted", "save_license", "show_info",
    ],
    "from delta_instance import (": [
        "_lock_path", "_read_lock_pid", "_take_over_locked_instance", "_write_lock_pid",
    ],
}

text = io.open(PATH, encoding="utf-8").read()
lines = text.splitlines(keepends=True)
out = []
i = 0
while i < len(lines):
    line = lines[i]
    header = line.strip()
    if header in DROP:
        names = DROP[header]
        i += 1
        block = []
        while i < len(lines) and lines[i].strip() != ")":
            block.append(lines[i])
            i += 1
        i += 1  # skip the closing paren
        kept = [b for b in block if b.strip().rstrip(",") not in names]
        if kept:
            out.append(line)
            out.extend(kept)
            out.append(")\n")
        continue
    out.append(line)
    i += 1

text = "".join(out)

# Plain imports the split made redundant.
for name in ("import config\n", "import hashlib\n", "import hmac\n", "import sys\n"):
    text = text.replace(name, "", 1)
text = text.replace("from product_identity import identify_product_directory\n", "", 1)
text = text.replace("from delta_performance import PerformanceMetrics\n", "", 1)
text = text.replace("from datetime import datetime, timezone, timedelta, date\n",
                    "from datetime import datetime, timezone, date\n", 1)
text = text.replace("    MONTHS_REV,\n", "", 1)

io.open(PATH, "w", encoding="utf-8", newline="\n").write(text)
print("cleaned imports in", PATH)
