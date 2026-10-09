"""Bounded timings and behavior-preserving batched ITM color ranges."""

import time
from collections import defaultdict
from bisect import bisect_left, bisect_right


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
