"""Instance guard and calculation-mode lifecycle helpers.

One process owns this folder: the lock file plus a product-exe process scan
lets a double-click take over from an older copy without touching a tool from
the other product folder. The calc-mode restore/forget logic rides along
because it belongs to the same process lifecycle.
"""

import atexit
import os
import platform
import time

import calc_mode_guard

from us_state import (
    BASE_DIR,
    PRODUCT_EXE,
    PRODUCT_ID,
    ROLE_ORDER,
    _calc_session,
    _last_bulk_restore,
)
from us_workbook import excel_path
from us_excel import find_open_book, _set_calc_mode


def _instance_guard_log(message):
    try:
        path = os.path.join(BASE_DIR, "instance_guard.log")
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except OSError:
        pass


def _same_folder(first, second):
    """True when two paths name the same directory."""
    try:
        return os.path.normcase(os.path.abspath(first)) == os.path.normcase(
            os.path.abspath(second)
        )
    except Exception:
        return False


def _lock_path():
    return os.path.join(BASE_DIR, "instance.lock")


def _read_lock_pid():
    """PID recorded in this folder's lock file, or None."""
    try:
        with open(_lock_path(), encoding="ascii") as fh:
            return int(fh.read().strip().split()[0])
    except (OSError, ValueError, IndexError):
        return None


def _write_lock_pid(pid):
    try:
        with open(_lock_path(), "w", encoding="ascii") as fh:
            fh.write(str(int(pid)))
    except OSError:
        pass


def _pid_alive(pid):
    """True only if the process is genuinely still running.

    OpenProcess succeeds on a process that has been terminated but whose
    object has not been released yet, so opening it is not proof of life. The
    exit code is: STILL_ACTIVE (259) means it is running, anything else means
    it has finished.
    """
    if not pid or pid <= 0:
        return False
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            SYNCHRONIZE = 0x00100000
            STILL_ACTIVE = 259
            handle = ctypes.windll.kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, False, int(pid)
            )
            if not handle:
                return False
            try:
                code = wintypes.DWORD()
                if not ctypes.windll.kernel32.GetExitCodeProcess(
                    handle, ctypes.byref(code)
                ):
                    return False
                return code.value == STILL_ACTIVE
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        except Exception:
            return False
    try:
        os.kill(int(pid), 0)
    except OSError:
        return False
    except Exception:
        return True
    return True


def _terminate_pid(pid):
    """Terminate a process. True when it is gone afterwards."""
    if os.name == "nt":
        try:
            import ctypes
            PROCESS_TERMINATE = 0x0001
            SYNCHRONIZE = 0x00100000
            handle = ctypes.windll.kernel32.OpenProcess(
                PROCESS_TERMINATE | SYNCHRONIZE, False, int(pid)
            )
            if not handle:
                return not _pid_alive(pid)
            try:
                ctypes.windll.kernel32.TerminateProcess(handle, 1)
                # Wait for it to actually go, rather than guessing.
                ctypes.windll.kernel32.WaitForSingleObject(handle, 5000)
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        except Exception:
            return not _pid_alive(pid)
    else:
        try:
            import signal
            os.kill(int(pid), signal.SIGTERM)
        except Exception:
            return not _pid_alive(pid)
    for _ in range(30):
        if not _pid_alive(pid):
            return True
        time.sleep(0.1)
    return not _pid_alive(pid)


def _take_over_locked_instance():
    """Stop an older instance of this tool running from this same folder.

    The lock file lives in this folder, so a PID found in it is by definition
    an instance of this tool in this folder. That is what lets a second
    launch take over even when it cannot read the other process's executable
    path - which is what happens when the two run at different privilege
    levels or when security software blocks the query.

    Returns True when another instance was found and stopped, False when there
    was none, and raises RuntimeError when one was found but could not be
    stopped, because two writers on the same workbooks is worse than one
    writer refusing to start.
    """
    pid = _read_lock_pid()
    if not pid or pid == os.getpid():
        return False
    if not _pid_alive(pid):
        _instance_guard_log(f"  lock pid={pid} is gone, taking over")
        return False
    _instance_guard_log(f"  lock pid={pid} is alive in this folder, stopping it")
    if _terminate_pid(pid):
        _instance_guard_log(f"  lock pid={pid} stopped; this launch takes over")
        return True
    raise RuntimeError(
        f"Another copy of {PRODUCT_EXE} (pid {pid}) is still running from "
        f"this folder and could not be stopped."
    )


def _process_image_path(kernel32, pid):
    """Full path of a running process, or None if it cannot be read."""
    import ctypes
    from ctypes import wintypes

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return None
    try:
        size = wintypes.DWORD(32768)
        buf = ctypes.create_unicode_buffer(size.value)
        if not kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return None
        return buf.value
    finally:
        kernel32.CloseHandle(h)


def kill_previous_instances():
    """Stop an older copy of *this* tool so a double-click takes over cleanly.

    The executable is named for the product - us_option_chain.exe - so the
    process name alone says which tool it is. That removes the need to guess
    the product from the folder, which is what went wrong when both packages
    were extracted into the same place.

    The lock file in this folder is checked first. It proves that a process
    running from this folder is an instance of this tool, so it can be stopped
    without having to read its executable path - which is exactly what fails
    when the two run at different privilege levels or when security software
    blocks the query.

    Never let the process scan break startup: any failure there is logged and
    ignored. A duplicate that cannot be stopped is a hard error, because two
    writers on the same workbooks is worse than one writer refusing to start.
    """
    if platform.system() != "Windows":
        return
    my_pid = os.getpid()
    my_dir = os.path.normcase(os.path.abspath(BASE_DIR))
    _instance_guard_log(
        f"pid={my_pid} product={PRODUCT_ID} exe={PRODUCT_EXE} dir={my_dir}"
    )

    # A previous instance of this tool in this folder, stopped without needing
    # to read its path.
    try:
        _take_over_locked_instance()
    except RuntimeError as e:
        _instance_guard_log(f"  duplicate could not be stopped: {e}")
        raise
    except Exception as e:
        _instance_guard_log(f"  lock check failed: {type(e).__name__}: {e}")

    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
        TH32CS_SNAPPROCESS = 0x2

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.POINTER(wintypes.ULONG)),
                ("th32ModuleID", wintypes.DWORD),
                ("th32Threads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_char * 260),
            ]

        snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if snap == -1 or snap is None:
            _instance_guard_log(f"snapshot failed; my_dir={my_dir}")
            return
        try:
            entry = PROCESSENTRY32()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
            matches = []
            if kernel32.Process32First(snap, ctypes.byref(entry)):
                while True:
                    # The executable is named for the product, so the process
                    # name alone identifies it. The other product's tool has a
                    # different name and is never a candidate.
                    if entry.szExeFile.decode(errors="ignore").lower() == PRODUCT_EXE:
                        matches.append(int(entry.th32ProcessID))
                    if not kernel32.Process32Next(snap, ctypes.byref(entry)):
                        break
        finally:
            kernel32.CloseHandle(snap)

        others = [pid for pid in matches if pid != my_pid]
        _instance_guard_log(
            f"pid={my_pid} product={PRODUCT_ID} dir={my_dir} "
            f"matched={matches} others={others}"
        )

        for pid in others:
            path = _process_image_path(kernel32, pid)
            if not path:
                # Cannot read the path, so cannot tell which folder it is in.
                # The lock file already covers the same-folder case, so leave
                # this one alone rather than killing another product's tool.
                _instance_guard_log(f"  pid={pid} path unreadable, left alone")
                continue
            other_dir = os.path.normcase(os.path.dirname(os.path.abspath(path)))
            if not _same_folder(other_dir, my_dir):
                _instance_guard_log(
                    f"  pid={pid} at {other_dir}, different folder, left alone"
                )
                continue
            if _terminate_pid(pid):
                _instance_guard_log(
                    f"  pid={pid} at {other_dir}, stopped; this launch takes over"
                )
    except Exception as e:  # never block startup on this
        _instance_guard_log(f"guard error: {type(e).__name__}: {e}")
    finally:
        # Record ourselves as the instance holding this folder, so the next
        # launch can take over without reading process paths.
        _write_lock_pid(my_pid)

def recover_calc_mode():
    """Put Excel back on automatic if a previous run died mid-write-window.

    A crash, a kill from Task Manager or a power cut while calculation was
    suspended would otherwise leave the customer's formulas silently stopped.
    The record is consulted first; Excel's current mode is only read if there
    is something to undo.
    """
    should, mode, reason = calc_mode_guard.pending_recovery()
    if not should:
        return False
    if reason:
        print(f"  calculation-mode record found: {reason}", flush=True)
    try:
        import xlwings as xw

        app = None
        for _ in range(10):
            for role in ROLE_ORDER:
                book = find_open_book(xw, excel_path(role))
                if book is not None:
                    app = book.app
                    break
            if app is not None:
                break
            time.sleep(0.5)
        if app is None:
            print("  calc recovery: no chain workbook is reachable", flush=True)
            return False          # keep the record so a later launch can retry
        if int(app.api.Calculation) != calc_mode_guard.XL_MANUAL:
            # Excel is not on manual, so there is nothing to undo. Clear the
            # record rather than touching a mode the customer may have set.
            calc_mode_guard.clear_record()
            return False
        app.calculation = mode or "automatic"
        calc_mode_guard.clear_record()
        print(f"  restored Excel calculation to {mode or 'automatic'}", flush=True)
        return True
    except Exception as e:
        print(f"  could not restore calculation mode: {e}", flush=True)
        return False


def _emergency_restore():
    """Put Excel back the way we found it, on a clean exit.

    The calculation mode is held for the whole session, so it has to be
    released here. Only if this tool set it in the first place - if it was
    already on manual calculation when we started, that belongs to the
    customer or to the other OptionChain tool and is left alone.
    """
    restore = _last_bulk_restore
    if restore is not None:
        try:
            restore()
        except Exception:
            pass
    if not _calc_session.get("owned"):
        return
    app = _calc_session.get("app")
    mode = _calc_session.get("mode")
    if app is None or not mode:
        return
    try:
        _ = app.version
    except Exception:
        return
    try:
        for _attempt in range(3):
            try:
                if _set_calc_mode(app, mode):
                    break
            except Exception:
                pass
            time.sleep(0.5)
        calc_mode_guard.release(PRODUCT_ID)
    except Exception:
        pass
    finally:
        _calc_session["owned"] = False


atexit.register(_emergency_restore)
