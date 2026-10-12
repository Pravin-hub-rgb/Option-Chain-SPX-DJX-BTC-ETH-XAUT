"""Crash-safe record of Excel's calculation mode.

Both Option Chain tools can run inside the same Excel instance. Suspending
calculation is what keeps a customer's strategy workbook usable, but if this
process dies while Excel is suspended -- killed from Task Manager, power loss,
or Excel itself crashing -- Excel is left on manual and their formulas silently
stop updating. That is worse than the freeze this exists to help with.

So the mode we found is written to a small shared file BEFORE anything is
changed. A clean restore removes it. If it is still there on the next launch,
the previous run died mid-window and the mode is put back.

The file is shared by both tools, with one entry per product, because they act
on the same Excel instance. Without that, whichever tool looked last would
treat the other tool's temporary manual mode as the customer's own setting.
"""

import json
import os
import platform
import tempfile
import time

RECORD_NAME = "calc_mode.json"

# A record older than this is history, not a live crash. The customer may have
# deliberately chosen manual in the meantime, and we must not undo that.
MAX_AGE_SECONDS = 600

XL_MANUAL = -4135


def _state_dir():
    base = os.environ.get("LOCALAPPDATA")
    if not base:
        return None
    path = os.path.join(base, "OptionChain")
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        return None
    return path


def _record_path():
    directory = _state_dir()
    return os.path.join(directory, RECORD_NAME) if directory else None


def _read():
    path = _record_path()
    if not path:
        return None
    try:
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict) or not data.get("original_mode"):
        return None
    return data


def _write(data):
    """Atomic, so a kill part way through cannot leave a half-written record."""
    directory = _state_dir()
    if not directory:
        return
    handle = None
    tmp = None
    try:
        fd, tmp = tempfile.mkstemp(dir=directory, suffix=".tmp")
        handle = os.fdopen(fd, "w", encoding="utf-8")
        json.dump(data, handle)
        handle.flush()
        os.fsync(handle.fileno())
        handle.close()
        handle = None
        os.replace(tmp, _record_path())
    except OSError:
        if handle is not None:
            try:
                handle.close()
            except OSError:
                pass
        if tmp:
            try:
                os.remove(tmp)
            except OSError:
                pass


def clear_record():
    path = _record_path()
    if not path:
        return
    try:
        os.remove(path)
    except OSError:
        pass


def _process_alive(pid):
    """True if the pid is running.

    Deliberately not os.kill(pid, 0): on Windows that maps to TerminateProcess,
    so it would kill the very process we were only trying to look at. Both
    tools run in the same Excel instance, so that would be fatal.
    """
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    if platform.system() != "Windows":
        try:
            os.kill(pid, 0)
        except OSError:
            return False
        return True
    try:
        import ctypes
        from ctypes import wintypes
    except ImportError:
        return True  # cannot tell: assume alive, leave the record alone
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    STILL_ACTIVE = 259
    ERROR_INVALID_PARAMETER = 87
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        # OpenProcess fails for a dead pid as well as for one we may not open,
        # so the error decides it. Guessing "alive" here would block the very
        # recovery this module exists for.
        # GetLastError from the DLL, not ctypes.get_last_error: windll is not
        # created with use_last_error, so the ctypes copy is stale.
        error = int(kernel32.GetLastError())
        if error == ERROR_INVALID_PARAMETER:
            return False        # no such process: the run really is dead
        return True             # exists but out of reach: assume alive
    try:
        code = wintypes.DWORD()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
            return True
        return code.value == STILL_ACTIVE
    finally:
        kernel32.CloseHandle(handle)


def _someone_alive(data):
    """True if a process still listed in the record is running."""
    holders = data.get("holders")
    if not isinstance(holders, dict) or not holders:
        # No usable info: assume another tool owns this window rather than
        # restoring a mode that is still in use.
        return bool(holders)
    for pid in holders.values():
        if _process_alive(pid):
            return True
    return False


def acquire(original_mode, product):
    """Record the mode to return to, before Excel is changed.

    Only an automatic workbook is recorded. If Excel is already manual that is
    either the customer's own choice or another tool's window, and neither is
    ours to undo.
    """
    existing = _read()
    if existing is not None and existing.get("original_mode") == "automatic":
        holders = existing.get("holders")
        if not isinstance(holders, dict):
            holders = {}
        holders[str(product)] = os.getpid()
        existing["holders"] = holders
        existing["seen_at"] = time.time()
        _write(existing)
        return True
    _write(
        {
            "original_mode": original_mode,
            "created_at": time.time(),
            "seen_at": time.time(),
            "holders": {str(product): os.getpid()},
        }
    )
    return True


def release(product):
    """Drop our claim. If nobody else holds it, the record is no longer needed."""
    data = _read()
    if data is None:
        return
    holders = data.get("holders")
    if isinstance(holders, dict):
        holders.pop(str(product), None)
        data["holders"] = holders
        if holders:
            _write(data)
            return
    clear_record()


def pending_recovery():
    """(should_restore, mode, reason) for a record left behind by a dead run."""
    data = _read()
    if data is None:
        return False, None, ""
    if _someone_alive(data):
        return False, None, ""
    try:
        age = time.time() - float(data.get("seen_at") or data.get("created_at") or 0)
    except (TypeError, ValueError):
        age = 0
    if age > MAX_AGE_SECONDS:
        return False, None, f"record is {int(age)}s old, left alone"
    return (
        True,
        data.get("original_mode"),
        "previous run ended while Excel was on manual calculation",
    )