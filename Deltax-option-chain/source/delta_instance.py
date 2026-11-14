"""Single-instance guard for the Delta Option Chain tool.

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

import delta_state
from product_identity import identify_product_directory


def _instance_guard_log(message):
    try:
        path = os.path.join(delta_state.BASE_DIR, "instance_guard.log")
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except OSError:
        pass


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


def _lock_path():
    return os.path.join(delta_state.BASE_DIR, "instance.lock")


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
    an instance of this tool in this folder. That is what lets a second launch
    take over even when it cannot read the other process's executable path -
    which is what happens when the two run at different privilege levels or
    when security software blocks the query.

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
        f"Another copy of {delta_state.PRODUCT_EXE} (pid {pid}) is still running from "
        f"this folder and could not be stopped."
    )


def kill_previous_instances():
    """Stop an older copy of *this* tool so a double-click takes over cleanly.

    The executable is named for the product - delta_option_chain.exe - so the
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
    my_dir = os.path.normcase(os.path.abspath(delta_state.BASE_DIR))
    _instance_guard_log(
        f"pid={my_pid} product={delta_state.PRODUCT_ID} exe={delta_state.PRODUCT_EXE} dir={my_dir}"
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
        my_pid = os.getpid()
        my_dir = os.path.normcase(os.path.abspath(delta_state.BASE_DIR))
        my_product = identify_product_directory(delta_state.BASE_DIR)
        if my_product is None:
            _instance_guard_log(
                f"pid={my_pid} dir={my_dir} product unknown; leaving other instances alone"
            )
            return

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.POINTER(wintypes.ULONG)),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
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
                    if entry.szExeFile.decode(errors="ignore").lower() == delta_state.PRODUCT_EXE:
                        matches.append(int(entry.th32ProcessID))
                    if not kernel32.Process32Next(snap, ctypes.byref(entry)):
                        break
        finally:
            kernel32.CloseHandle(snap)

        others = [pid for pid in matches if pid != my_pid]
        _instance_guard_log(
            f"pid={my_pid} product={my_product} dir={my_dir} "
            f"matched={matches} others={others}"
        )

        for pid in others:
            path = _process_image_path(kernel32, pid)
            if not path:
                _instance_guard_log(f"  pid={pid} path unreadable, left alone")
                continue
            other_dir = os.path.normcase(os.path.dirname(os.path.abspath(path)))
            other_product = identify_product_directory(other_dir)
            if other_product != my_product:
                _instance_guard_log(
                    f"  pid={pid} product={other_product or 'unknown'} "
                    f"at {other_dir}, left alone"
                )
                continue
            if _terminate_pid(pid):
                _instance_guard_log(
                    f"  pid={pid} product={other_product} at {other_dir}, "
                    "stopped; this launch takes over"
                )
            else:
                _instance_guard_log(
                    f"  pid={pid} product={other_product} at {other_dir}, "
                    "could not be stopped; left running"
                )
    except Exception as e:  # never block startup on this
        _instance_guard_log(f"guard error: {type(e).__name__}: {e}")
    finally:
        # Record ourselves as the instance holding this folder, so the next
        # launch can take over without reading process paths.
        _write_lock_pid(my_pid)
