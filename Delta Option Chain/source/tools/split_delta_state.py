"""One-shot: move Delta's shared state into delta_state.py.

Pure code motion. Every definition is moved verbatim, main.py imports the
same objects, and the few places that REBIND a name (rather than mutate it)
are rewritten to delta_state.X = ... so there is still exactly one copy of
each piece of state.

Run from the Delta Option Chain folder.
"""

import ast
import io
import os
import sys

SRC = "source/main.py"
STATE_SRC = "source/delta_state.py"

# Names that become shared state. Everything here is either a constant the
# whole product reads, or a container the whole product mutates.
STATE_NAMES = [
    "BASE_DIR",
    "PRODUCT_ID",
    "PRODUCT_EXE",
    "REST_BASE",
    "WS_URL",
    "INDEX_WS_URL",
    "live_data",
    "spot_cmp",
    "symbol_expiry",
    "expiry_symbols",
    "active_expiries",
    "active_expiries_dmy",
    "user_pinned",
    "cfg",
    "HTTP_TIMEOUT",
    "OI_POLL_SECONDS",
    "EXPIRY_REFRESH_SECONDS",
    "SPOT_POLL_SECONDS",
    "NET_FAIL_LIMIT",
    "_net",
    "_net_drawn",
    "_pending_date_loads",
    "_date_req_sig",
    "_book_sheet_cache",
    "_sheet_cache",
    "_app_cache",
    "_last_rows",
    "_FIRST_SIGHT_BOUND",
    "_block_payload",
    "_block_values",
    "_block_hidden",
    "_colour_state",
    "available_dates",
    "_dropdown_sig",
    "_usd_fmt_done",
    "_performance_metrics",
    "_last_performance_flush",
    "_performance_log_error_shown",
    "_excel_connection_status",
    "_last_excel_exception",
    "MONTHS",
    "MONTHS_REV",
    "BLOCK_COLS",
    "STALE_FLAG_SECONDS",
    "feed_state",
    "SPOT_TICKER_SYMBOLS",
    "INDEX_SYMBOLS",
    "INDEX_CHANNEL",
    "INDEX_ASSET",
    "last_ws_recv",
    "last_spot_change",
    "spot_source",
    "STALE_AFTER_S",
    "excel_fail_since",
    "excel_fail_notice_shown",
    "excel_connected",
    "excel_miss_counts",
    "excel_first_attempt",
    "excel_startup_notice_shown",
    "_auto_open_done",
    "_sheet_altered_notice",
    "_excel_liveness_cache",
    "_bulk_owning_manual",
    "_calc_state_probe",
    "_last_window_status",
    "_SHORT_BUILD",
    "_workbook_missing",
    "_manual_since",
    "MANUAL_WATCHDOG_SECONDS",
    "_CALC_IDS",
    "_CALC_ID_NAMES",
    "_last_observed_calc_raw",
    "_calc_churn_count",
    "_calc_churn_summary_after",
    "_CALC_CHURN_SUMMARY_SECONDS",
    "_calc_watchdog_skip_logged_at",
    "_cycle_wrote",
    "_bulk_will_recalc",
    "_last_recalc_ms",
    "_speed_log_at",
    "_last_bulk_restore",
    "_excel_gone_logged",
]

STATE_HEADER = '''"""Shared state for the Delta Option Chain tool.

Every piece of state the product has lives here, so there is exactly one copy
of each. The modules import the objects they need from here and mutate them
in place; the few places that REBIND a name rather than mutate it write to
delta_state.<name> instead, so the change is seen by everyone.

This mirrors the US tool's us_state.py exactly, so the two products have the
same shape and a fix to one reads the same in the other.
"""

import os
import time

import config
from performance_metrics import PerformanceMetrics

'''

IMPORTS = None  # built at run time


def main():
    src = io.open(SRC, encoding="utf-8").read()
    lines = src.splitlines(keepends=True)
    tree = ast.parse(src)

    # A name that is assigned inside a function is rebound, not mutated. It
    # must never be imported by value, or the reader and the writer end up
    # looking at two different objects.
    rebound = compute_rebound(tree)

    imported = [n for n in sorted(STATE_NAMES) if n not in rebound]
    imports = (
        "from delta_state import (\n"
        + "".join(f"    {name},\n" for name in imported)
        + ")\n"
    )

    # ---- collect the definitions to move -----------------------------------
    segments = []  # (start_line, end_line, text)
    found = set()
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        names = [t.id for t in node.targets if isinstance(t, ast.Name)]
        picked = [n for n in names if n in STATE_NAMES]
        if not picked:
            continue
        text = ast.get_source_segment(src, node)
        if text is None:
            raise SystemExit(f"no source segment for line {node.lineno}")
        segments.append((node.lineno, node.end_lineno, text.rstrip() + "\n"))
        found.update(picked)

    missing = [n for n in STATE_NAMES if n not in found]
    if missing:
        print("NOT FOUND AT MODULE LEVEL:", missing)

    # ---- write delta_state.py ---------------------------------------------
    ordered = sorted(segments, key=lambda s: s[0])
    with io.open(STATE_SRC, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(STATE_HEADER)
        fh.write("\n".join(text for _, _, text in ordered))
        fh.write("\n")
    print(f"wrote {STATE_SRC} with {len(ordered)} definitions")
    print(f"imported by name: {len(imported)}, routed through delta_state: {len(rebound)}")

    # ---- rewrite main.py ---------------------------------------------------
    drop = set()
    for start, end, _ in segments:
        for ln in range(start, end + 1):
            drop.add(ln)

    first_line = min(s for s, _, _ in segments)
    kept = []
    for idx, line in enumerate(lines, start=1):
        if idx == first_line:
            kept.append(imports)
        if idx in drop:
            continue
        kept.append(line)
    out = "".join(kept)

    # ---- rewrite the rebinds ----------------------------------------------
    out = rewrite_rebinds(out)

    with io.open(SRC, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(out)
    print("rewrote", SRC)


def _scope_chain(tree):
    """Map every function node to the chain of functions that contain it."""
    parents = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parents[child] = node
    return parents


def compute_rebound(tree):
    """State names that are rebound rather than mutated.

    A name declared global anywhere in the file is rebound, not mutated. If
    main.py imported it by value, the reader and the writer would end up
    looking at two different objects.

    A nested function inherits the enclosing function's global names, so an
    assignment there is meant to update the module-level value. The restore()
    closure assigns _last_bulk_restore that way, and without it
    _emergency_restore() would never find a restore to call.
    """
    rebound = set()
    for func in ast.walk(tree):
        if not isinstance(func, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        for stmt in ast.walk(func):
            if isinstance(stmt, ast.Global):
                rebound.update(n for n in stmt.names if n in STATE_NAMES)
    return rebound


def rewrite_rebinds(src):
    """Route every use of a REBOUND name through delta_state.

    A rebound name is never imported by name: it is always written and read
    as delta_state.<name>, everywhere in the file.
    """
    tree = ast.parse(src)
    lines = src.splitlines(keepends=True)
    statements = {}  # (start_idx, end_idx) -> replacement
    names = {}       # line index -> set of (col, length, text)

    rebound = compute_rebound(tree)
    if not rebound:
        return src

    for func in ast.walk(tree):
        if not isinstance(func, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        for stmt in ast.walk(func):
            if isinstance(stmt, ast.Global):
                keep = [n for n in stmt.names if n not in STATE_NAMES]
                if len(keep) == len(stmt.names):
                    continue
                text = "    global " + ", ".join(keep) + "\n" if keep else ""
                statements[(stmt.lineno - 1, stmt.end_lineno)] = text

    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and node.id in rebound:
            names.setdefault(node.lineno - 1, set()).add(
                (node.col_offset, len(node.id), f"delta_state.{node.id}")
            )

    # Name edits first, right to left within each line so the columns stay
    # valid while a line is being rewritten.
    for lineno, edits in names.items():
        line = lines[lineno]
        for col, length, text in sorted(edits, key=lambda e: -e[0]):
            line = line[:col] + text + line[col + length:]
        lines[lineno] = line

    # Then the global statements, from the bottom up. The dictionary keys are
    # unique, so a nested function's global statement is only removed once.
    for (start, end), text in sorted(statements.items(), key=lambda e: (-e[0][0], -e[0][1])):
        lines[start:end] = [text] if text else []
    return "".join(lines)


if __name__ == "__main__":
    main()
