"""Per-user startup, Excel connection and aggregate performance diagnostics."""

import os
import platform
import struct
import sys
import threading
import traceback
from datetime import datetime

_LOG_LOCK = threading.Lock()
_MAX_LOG_BYTES = 1024 * 1024


def _log_directory():
    root = os.environ.get("LOCALAPPDATA")
    if not root:
        raise OSError("LOCALAPPDATA is not set")
    directory = os.path.join(root, "USOptionChain")
    os.makedirs(directory, exist_ok=True)
    return directory


def _prepare(path):
    if os.path.exists(path) and os.path.getsize(path) >= _MAX_LOG_BYTES:
        rotated = path + ".1"
        if os.path.exists(rotated):
            os.remove(rotated)
        os.replace(path, rotated)
    return path


def _build_id():
    """Identify this build, so support knows what the customer is running.

    build_info.txt is written by build_exe.bat next to the executable. It
    replaces the separate .sha256 file that used to be delivered beside the
    ZIP: customers could not fetch that from Drive, and even if they could it
    told them nothing about which build they had. This is self-reported, so it
    works whether or not anyone checks anything.
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    try:
        with open(os.path.join(base, "build_info.txt"), encoding="utf-8") as handle:
            parts = [line.strip() for line in handle if line.strip()]
    except OSError:
        return "not stamped (rebuilt outside build_exe.bat)"
    if not parts:
        return "not stamped (build_info.txt was empty)"
    return " ".join(parts)


def write_startup_diagnostics():
    values = {
        "build_id": _build_id(),
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "platform": platform.platform(),
        "windows_release": platform.release(),
        "windows_version": platform.version(),
        "machine": platform.machine(),
        "python_version": sys.version.replace("\n", " "),
        "python_process_bitness": f"{struct.calcsize('P') * 8}-bit",
        "application_path": os.path.abspath(sys.argv[0]),
    }
    try:
        import ctypes
        values["application_process_is_administrator"] = str(
            bool(ctypes.windll.shell32.IsUserAnAdmin())
        )
    except (AttributeError, OSError) as exc:
        values["application_process_is_administrator"] = f"unavailable: {exc}"
    try:
        import xlwings
        values["xlwings_version"] = getattr(xlwings, "__version__", "version unknown")
    except ImportError as exc:
        values["xlwings_version"] = f"import failed: {exc}"

    if platform.system() == "Windows":
        try:
            import winreg

            release_values = []
            for view_name, view in (
                ("64-bit", winreg.KEY_WOW64_64KEY),
                ("32-bit", winreg.KEY_WOW64_32KEY),
            ):
                try:
                    with winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE,
                        r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full",
                        0,
                        winreg.KEY_READ | view,
                    ) as key:
                        release, _ = winreg.QueryValueEx(key, "Release")
                        release_values.append(f"{view_name}:{release}")
                except FileNotFoundError:
                    continue
            values["dotnet_release"] = ",".join(release_values) or "not detected"
            office_fields = {}
            for view_name, view in (
                ("64-bit", winreg.KEY_WOW64_64KEY),
                ("32-bit", winreg.KEY_WOW64_32KEY),
            ):
                try:
                    with winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE,
                        r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                        0,
                        winreg.KEY_READ | view,
                    ) as key:
                        for name in ("VersionToReport", "Platform"):
                            try:
                                office_fields[f"{view_name}.{name}"] = str(
                                    winreg.QueryValueEx(key, name)[0]
                                )
                            except FileNotFoundError:
                                continue
                except FileNotFoundError:
                    continue
                for office_version in ("16.0", "15.0", "14.0"):
                    try:
                        with winreg.OpenKey(
                            winreg.HKEY_LOCAL_MACHINE,
                            rf"SOFTWARE\Microsoft\Office\{office_version}\Excel",
                            0,
                            winreg.KEY_READ | view,
                        ) as key:
                            bitness, _ = winreg.QueryValueEx(key, "Bitness")
                            office_fields[f"{view_name}.Office{office_version}.Bitness"] = str(
                                bitness
                            )
                    except FileNotFoundError:
                        continue
            values["excel_registry"] = (
                ";".join(f"{key}={value}" for key, value in office_fields.items())
                or "not identified"
            )
        except ImportError as exc:
            values["dotnet_release"] = f"registry unavailable: {exc}"

    directory = _log_directory()
    path = os.path.join(directory, "startup_diag.log")
    with _LOG_LOCK:
        _prepare(path)
        with open(path, "a", encoding="utf-8") as log_file:
            log_file.write("\n--- US Option Chain startup ---\n")
            log_file.writelines(f"{key}={value}\n" for key, value in values.items())
    return path


def excel_process_details(hwnd):
    """Return bitness and elevation details for the Excel process."""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)

    process_id = wintypes.DWORD()
    user32.GetWindowThreadProcessId.argtypes = (
        wintypes.HWND,
        ctypes.POINTER(wintypes.DWORD),
    )
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    if not user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id)):
        raise ctypes.WinError(ctypes.get_last_error())

    kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    process = kernel32.OpenProcess(0x1000, False, process_id.value)
    if not process:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        image_path = "unknown"
        image_bitness = "unknown"
        image_buffer = ctypes.create_unicode_buffer(32768)
        image_size = wintypes.DWORD(len(image_buffer))
        kernel32.QueryFullProcessImageNameW.argtypes = (
            wintypes.HANDLE,
            wintypes.DWORD,
            wintypes.LPWSTR,
            ctypes.POINTER(wintypes.DWORD),
        )
        kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
        if kernel32.QueryFullProcessImageNameW(
            process, 0, image_buffer, ctypes.byref(image_size)
        ):
            image_path = image_buffer.value
            try:
                with open(image_path, "rb") as image:
                    header = image.read(64)
                    if header[:2] == b"MZ" and len(header) >= 64:
                        pe_offset = int.from_bytes(header[60:64], "little")
                        image.seek(pe_offset)
                        pe_header = image.read(6)
                        if pe_header[:4] == b"PE\0\0":
                            machine = int.from_bytes(pe_header[4:6], "little")
                            image_bitness = {
                                0x014C: "32-bit",
                                0x8664: "64-bit",
                                0xAA64: "ARM64",
                            }.get(machine, f"unknown-0x{machine:04X}")
            except OSError:
                pass

        token = wintypes.HANDLE()
        advapi32.OpenProcessToken.argtypes = (
            wintypes.HANDLE,
            wintypes.DWORD,
            ctypes.POINTER(wintypes.HANDLE),
        )
        advapi32.OpenProcessToken.restype = wintypes.BOOL
        if not advapi32.OpenProcessToken(process, 0x0008, ctypes.byref(token)):
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            elevation = wintypes.DWORD()
            returned = wintypes.DWORD()
            advapi32.GetTokenInformation.argtypes = (
                wintypes.HANDLE,
                ctypes.c_int,
                ctypes.c_void_p,
                wintypes.DWORD,
                ctypes.POINTER(wintypes.DWORD),
            )
            advapi32.GetTokenInformation.restype = wintypes.BOOL
            if not advapi32.GetTokenInformation(
                token,
                20,
                ctypes.byref(elevation),
                ctypes.sizeof(elevation),
                ctypes.byref(returned),
            ):
                raise ctypes.WinError(ctypes.get_last_error())
            return {
                "process_id": process_id.value,
                "bitness": image_bitness,
                "is_administrator": bool(elevation.value),
            }
        finally:
            kernel32.CloseHandle(token)
    finally:
        kernel32.CloseHandle(process)


def _append(section, lines, filename="startup_diag.log"):
    path = os.path.join(_log_directory(), filename)
    with _LOG_LOCK:
        _prepare(path)
        with open(path, "a", encoding="utf-8") as log_file:
            log_file.write(f"\n--- US Option Chain {section} ---\n")
            log_file.writelines(f"{line}\n" for line in lines)
    return path


def append_connection_event(role, status, detail=""):
    lines = [
        f"timestamp={datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"role={role}",
        f"status={status}",
    ]
    if detail:
        lines.append(f"detail={detail}")
    return _append("Excel connection", lines)


def append_performance_summary(lines):
    return _append("performance", lines)


def append_app_exception(context, exc):
    rendered = traceback.format_exception(type(exc), exc, exc.__traceback__)
    lines = [
        f"timestamp={datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"context={context}",
        *"".join(rendered).splitlines(),
    ]
    return _append("application exception", lines, "app.log")
