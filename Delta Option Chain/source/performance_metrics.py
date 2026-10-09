"""Bounded, aggregate-only performance metrics for the Excel refresh loop."""

import time
from collections import defaultdict


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
        """Return a fixed-size snapshot without exposing any underlying values."""
        now = time.monotonic() if now is None else now
        interval_seconds = max(0.0, now - self._started)
        lines = [
            "performance.interval_seconds="
            f"{interval_seconds:.1f}; scopes are assets or ALL"
        ]
        for (scope, phase), sample in sorted(self._samples.items()):
            average_ms = sample["total_ms"] / sample["count"]
            lines.append(
                "performance."
                f"{scope}.{phase}: count={sample['count']}, "
                f"avg_ms={average_ms:.2f}, max_ms={sample['max_ms']:.2f}, "
                f"rows={sample['rows']}"
            )
        return lines

    def reset(self, now=None):
        """Start a fresh aggregation interval after a summary has been persisted."""
        now = time.monotonic() if now is None else now
        self._samples.clear()
        self._started = now
