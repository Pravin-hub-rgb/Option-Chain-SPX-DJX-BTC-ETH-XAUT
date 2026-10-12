"""Controls the trader can use without asking anyone for help.

Two cells, both read through COM once per cycle:

    G1   type ON  to pause live writes, OFF to resume
    I1   type a number for the minimum refresh interval in seconds

They exist because the customers using this are not technical. Asking them to
edit a config file, or to call someone to change an interval, is not a
realistic support story. Typing a word into a labelled cell is.

Both reads are deliberately forgiving: the cell may hold text mid-edit, a
half-typed number, or something else entirely. Anything unrecognised falls
back to the default rather than raising, because a raised exception here would
stop the chain.

These cells sit in row 1 past the PCR readout. Nothing else writes there, and
the chain writes, colour repaints and number formatting all start at row 5, so
the controls cannot be overwritten by the tool.
"""

import time

PAUSE_LABEL_CELL = "F1"
PAUSE_CELL = "G1"
REFRESH_LABEL_CELL = "H1"
REFRESH_CELL = "I1"

PAUSE_LABEL = "Pause:"
REFRESH_LABEL = "Secs:"

DEFAULT_REFRESH = 1.0
MIN_REFRESH = 0.5
MAX_REFRESH = 30.0

_TRUTHY = {"on", "yes", "1", "true", "pause", "paused", "y"}
_FALSY = {"off", "no", "0", "false", "run", "running", ""}


def read_controls(sheet):
    """Return (paused, refresh_seconds). Never raises."""
    paused = False
    refresh = DEFAULT_REFRESH

    try:
        raw = sheet.range(PAUSE_CELL).value
        token = str(raw).strip().lower() if raw is not None else ""
        if token in _TRUTHY:
            paused = True
        elif token in _FALSY:
            paused = False
    except Exception:
        # Mid-edit, or the cell is unavailable. Keep the previous behaviour.
        pass

    try:
        raw = sheet.range(REFRESH_CELL).value
        if raw is not None:
            value = float(str(raw).strip())
            if MIN_REFRESH <= value <= MAX_REFRESH:
                refresh = value
    except (TypeError, ValueError):
        pass

    return paused, refresh


def apply_labels(sheet, cache):
    """Write the labels and the dropdown once per sheet.

    Only on first sighting: writing them every cycle would dirty the sheet and
    trigger needless recalculation on the customer's dependent formulas.
    """
    key = id(sheet)
    if cache.get(key):
        return
    cache[key] = True
    try:
        sheet.range(PAUSE_LABEL_CELL).value = PAUSE_LABEL
        sheet.range(REFRESH_LABEL_CELL).value = REFRESH_LABEL
        for cell in (PAUSE_LABEL_CELL, REFRESH_LABEL_CELL):
            try:
                sheet.range(cell).font.bold = True
                sheet.range(cell).font.color = (89, 89, 89)
            except Exception:
                pass
    except Exception:
        pass
    try:
        cell = sheet.range(PAUSE_CELL).api.Validation
        cell.Delete()
        cell.Add(Type=3, Formula1="OFF,ON")
        cell.ShowError = False      # a dropdown is a hint; typing still works
    except Exception:
        # Excel raises when the cell has no existing validation, which is the
        # normal case here.
        pass


def describe(paused, refresh, recalc_ms=None):
    """One line for the sheet header and the window title."""
    stamp = time.strftime("%H:%M:%S")
    if paused:
        return f"PAUSED  |  writes on hold, resume by typing OFF in {PAUSE_CELL}  |  {stamp}"
    text = f"Live every {refresh:g}s"
    if recalc_ms is not None and recalc_ms > 0:
        text = f"Refresh slowed to {max(refresh, recalc_ms / 1000.0 * 3):g}s (workbook recalc {recalc_ms / 1000.0:.1f}s)"
    return f"{text}  |  {stamp}"