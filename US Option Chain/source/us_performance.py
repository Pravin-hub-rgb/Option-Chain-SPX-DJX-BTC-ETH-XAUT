"""Bounded timings and behavior-preserving batched ITM color ranges."""

import asyncio
import time
from collections import defaultdict
from bisect import bisect_left, bisect_right

from us_diagnostics import append_performance_summary


class PerformanceMetrics:
    def __init__(self):
        self._samples = defaultdict(
            lambda: {"count": 0, "total_ms": 0.0, "max_ms": 0.0, "rows": 0}
        )
        self._started = time.monotonic()

    def record(self, scope, phase, duration_seconds, rows=0):
        duration_ms = max(0.0, float(duration_seconds) * 1000)
        sample = self._samples[(scope, phase)]
        sample["count"] += 1
        sample["total_ms"] += duration_ms
        sample["max_ms"] = max(sample["max_ms"], duration_ms)
        sample["rows"] += max(0, int(rows))

    def summary_lines(self, now=None):
        now = time.monotonic() if now is None else now
        lines = [
            f"performance.interval_seconds={max(0.0, now - self._started):.1f}; "
            "scopes are roles or ALL"
        ]
        for (scope, phase), sample in sorted(self._samples.items()):
            average_ms = sample["total_ms"] / sample["count"]
            lines.append(
                f"performance.{scope}.{phase}: count={sample['count']}, "
                f"avg_ms={average_ms:.2f}, max_ms={sample['max_ms']:.2f}, "
                f"rows={sample['rows']}"
            )
        return lines

    def reset(self, now=None):
        self._samples.clear()
        self._started = time.monotonic() if now is None else now


def color_signature(strikes, spot, n_rows):
    """Fingerprint of everything the ITM/ATM colouring depends on.

    Two equal signatures mean the picture on screen is identical, so the block
    does not need repainting. Spot is part of it because the call and put
    boundaries are derived from spot, not only from the strike ladder — that is
    what used to leave stale colours behind when the price moved.
    """
    if not strikes:
        return (n_rows, 0, 0, 0)
    calls_below = bisect_left(strikes, spot)
    puts_above = bisect_right(strikes, spot)
    atm_index = min(range(len(strikes)), key=lambda i: abs(strikes[i] - spot))
    return (n_rows, calls_below, puts_above, atm_index)


def itm_color_ranges(block, strikes, spot):
    """Return contiguous ranges matching the old per-row ITM coloring logic."""
    if not strikes:
        return []
    calls_below_spot = bisect_left(strikes, spot)
    puts_above_spot = bisect_right(strikes, spot)
    ranges = []
    if calls_below_spot:
        ranges.append(
            (
                f"{block['call_start']}5:{block['call_end']}{4 + calls_below_spot}",
                (198, 224, 180),
            )
        )
    if puts_above_spot < len(strikes):
        ranges.append(
            (
                f"{block['put_start']}{5 + puts_above_spot}:"
                f"{block['put_end']}{4 + len(strikes)}",
                (255, 200, 200),
            )
        )
    return ranges


async def flush_performance_metrics():
    import us_state
    from us_licensing import show_error

    now = time.monotonic()
    if now - _last_performance_flush < 30:
        return
    lines = _performance_metrics.summary_lines(now)
    try:
        await asyncio.to_thread(append_performance_summary, lines)
    except OSError as e:
        if not us_state._performance_log_error_shown:
            show_error(
                "Could not append performance diagnostics to startup_diag.log.\n\n"
                "Check that your user profile is writable and has free space.\n\n"
                f"Details: {e}"
            )
            us_state._performance_log_error_shown = True
    else:
        us_state._performance_metrics.reset(now)
        us_state._last_performance_flush = now
