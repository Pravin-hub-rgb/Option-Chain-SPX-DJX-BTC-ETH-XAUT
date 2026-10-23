"""One-shot: carve Delta's self-contained blocks out of main.py.

The blocks are contiguous groups of functions and constants that share no
mutable state with the rest of the product, so moving them is pure code
motion:

  delta_instance     the single-instance guard
  delta_licensing    the product key and the licence

Run from the Delta Option Chain folder.
"""

import ast
import io

SRC = "source/main.py"

BLOCKS = {
    "source/delta_instance.py": {
        "doc": '''"""Single-instance guard for the Delta Option Chain tool.

Only one copy of this tool may write the three chain workbooks. Two writers
mean two chain sessions and a sheet that shows a mix of two snapshots, so a
duplicate that cannot be stopped is a hard error rather than something to
warn about and carry on from.

The lock file in the product folder is what makes a same-folder duplicate
findable: a PID found in it is by definition an instance of this tool in this
folder, so it can be stopped without having to read the other process's
executable path - which is exactly what fails when the two run at different
privilege levels or when security software blocks the query.
"""

import os
import platform
import time

from delta_state import BASE_DIR, PRODUCT_EXE, PRODUCT_ID

''',
        "names": [
            "_instance_guard_log",
            "_process_image_path",
            "_lock_path",
            "_read_lock_pid",
            "_write_lock_pid",
            "_pid_alive",
            "_terminate_pid",
            "_take_over_locked_instance",
            "kill_previous_instances",
        ],
        "consts": [],
        "imports": [
            "from delta_instance import (",
            "    _lock_path,",
            "    _read_lock_pid,",
            "    _take_over_locked_instance,",
            "    _write_lock_pid,",
            "    kill_previous_instances,",
            ")",
        ],
    },
    "source/delta_licensing.py": {
        "doc": '''"""Product key, activation and the licence file.

The licence lives next to the executable, so it is per-installation and never
travels with the product. The check is deliberately self-contained: it reads
the file, verifies the signature, and says so in plain words.
"""

import os
import platform
import time

from delta_state import MONTHS, MONTHS_REV

''',
        "names": [
            "_key_mask",
            "_vmask",
            "_key_check",
            "key_to_date",
            "key_accepted",
            "_license_sig",
            "_verify_sig",
            "license_is_valid",
            "save_license",
            "ask_product_key",
            "ensure_license",
            "show_error",
            "show_info",
        ],
        "consts": ["LICENSE_PATH", "_SK", "_SE", "_LICENSE_SECRET"],
        "imports": [
            "from delta_licensing import (",
            "    LICENSE_PATH,",
            "    _license_sig,",
            "    _verify_sig,",
            "    ask_product_key,",
            "    ensure_license,",
            "    key_accepted,",
            "    license_is_valid,",
            "    save_license,",
            "    show_error,",
            "    show_info,",
            ")",
        ],
    },
    "source/delta_diagnostics.py": {
        "doc": '''"""The feed log: one structured line per event, capped and never raising.

The feed log is what support asks the customer to send, so it has to be
complete and it has to stay quiet when there is nothing to say. The directory
honours DELTA_OPTION_CHAIN_LOG_DIR so tests and support runs never write
into the real customer log.
"""

import os
from datetime import datetime

''',
        "names": ["_flat", "feed_log"],
        "consts": ["FEED_LOG_DIR", "FEED_LOG_PATH", "FEED_LOG_MAX_BYTES"],
        "imports": [
            "from delta_diagnostics import (",
            "    FEED_LOG_DIR,",
            "    FEED_LOG_MAX_BYTES,",
            "    FEED_LOG_PATH,",
            "    feed_log,",
            ")",
        ],
    },
}


def main():
    src = io.open(SRC, encoding="utf-8").read()
    lines = src.splitlines(keepends=True)
    tree = ast.parse(src)

    found = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            found[node.name] = node
        elif isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name):
                    found[t.id] = node

    taken = set()
    for path, spec in BLOCKS.items():
        missing = [n for n in spec["names"] + spec["consts"] if n not in found]
        if missing:
            raise SystemExit(f"{path}: not found in main.py: {missing}")
        taken.update(spec["names"] + spec["consts"])

    # ---- write each block out ---------------------------------------------
    for path, spec in BLOCKS.items():
        pieces = []
        for name in spec["consts"]:
            node = found[name]
            pieces.append(ast.get_source_segment(src, node).rstrip() + "\n")
        for name in spec["names"]:
            node = found[name]
            pieces.append(ast.get_source_segment(src, node).rstrip() + "\n")
        with io.open(path, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(spec["doc"])
            fh.write("\n\n".join(pieces).rstrip() + "\n")
        print(f"wrote {path}")

    # ---- remove them from main.py -----------------------------------------
    drop = set()
    for name in taken:
        node = found[name]
        for ln in range(node.lineno, node.end_lineno + 1):
            drop.add(ln)

    kept = []
    inserted = set()
    for idx, line in enumerate(lines, start=1):
        for path, spec in BLOCKS.items():
            first = min(
                found[n].lineno for n in spec["names"] + spec["consts"]
            )
            if idx == first and path not in inserted:
                kept.append("\n".join(spec["imports"]) + "\n")
                inserted.add(path)
        if idx in drop:
            continue
        kept.append(line)

    out = "".join(kept)
    with io.open(SRC, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(out)
    print("rewrote", SRC)


if __name__ == "__main__":
    main()
