"""Collect Windows/Excel environment details for remote startup troubleshooting."""

import os
import platform
import struct
import sys
import threading
import traceback
from datetime import datetime


def _registry_diagnostics():
    import winreg

    values = {}
    errors = []
    views = (
        ("64-bit", winreg.KEY_WOW64_64KEY),
        ("32-bit", winreg.KEY_WOW64_32KEY),
    )

    def read_key(root, path, view, label):
        try:
            with winreg.OpenKey(root, path, 0, winreg.KEY_READ | view) as key:
                index = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, index)
                    except OSError as exc:
                        if getattr(exc, "winerror", None) != 259:
                            errors.append(f"{label}: {exc}")
                        break
                    values[f"{label}.{name or '(Default)'}"] = str(value)
                    index += 1
        except FileNotFoundError:
            return
        except OSError as exc:
            errors.append(f"{label}: {exc}")

    windows_key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    for view_name, view in views:
        read_key(
            winreg.HKEY_LOCAL_MACHINE,
            windows_key,
            view,
            f"windows.registry_{view_name}",
        )

    office_versions = ("16.0", "15.0", "14.0")
    for view_name, view in views:
        click_to_run = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        read_key(
            winreg.HKEY_LOCAL_MACHINE,
            click_to_run,
            view,
            f"excel.click_to_run_{view_name}",
        )
        for version in office_versions:
            install_root = (
                rf"SOFTWARE\Microsoft\Office\{version}\Excel\InstallRoot"
            )
            read_key(
                winreg.HKEY_LOCAL_MACHINE,
                install_root,
                view,
                f"excel.office_{version}_{view_name}",
            )

    dotnet_key = r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    release_found = False
    for view_name, view in views:
        try:
            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE, dotnet_key, 0, winreg.KEY_READ | view
            ) as key:
                release, _ = winreg.QueryValueEx(key, "Release")
                values[f"dotnet.release_{view_name}"] = str(release)
                release_found = True
        except FileNotFoundError:
            continue
        except OSError as exc:
            errors.append(f"dotnet.{view_name}: {exc}")
    if not release_found:
        values["dotnet.release"] = "not detected"

    click_to_run_platform = next(
        (
            value
            for key, value in values.items()
            if key.startswith("excel.click_to_run_") and key.endswith(".Platform")
        ),
        None,
    )
    excel_install = next(
        (
            (key, value)
            for key, value in values.items()
            if key.startswith("excel.office_") and key.endswith(".Path")
        ),
        None,
    )
    values["excel.version"] = next(
        (
            value
            for key, value in values.items()
            if key.startswith("excel.click_to_run_")
            and key.endswith(".VersionToReport")
        ),
        "not identified",
    )
    if values["excel.version"] == "not identified" and excel_install:
        install_key, _ = excel_install
        version_prefix = install_key.removeprefix("excel.office_")
        values["excel.version"] = (
            f"{version_prefix.split('_', 1)[0]} (registry major version)"
        )
    if click_to_run_platform:
        values["excel.bitness"] = click_to_run_platform
    elif platform.machine().lower() in ("x86", "i386", "i486", "i586", "i686"):
        values["excel.bitness"] = "32-bit (32-bit Windows)"
    elif excel_install:
        install_key, _ = excel_install
        values["excel.bitness"] = (
            "inferred 64-bit"
            if "_64-bit" in install_key
            else "inferred 32-bit"
        )
    else:
        values["excel.bitness"] = "not identified"

    return values, errors


def _build_id():
    """Identify this build, so support knows what the customer is running.

    build_info.txt is written by build_final.bat next to the executable. It
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
        return "not stamped (rebuilt outside build_final.bat)"
    if not parts:
        return "not stamped (build_info.txt was empty)"
    return " ".join(parts)


def _diagnostic_lines():
    data = {
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

        data["application_process_is_administrator"] = str(
            bool(ctypes.windll.shell32.IsUserAnAdmin())
        )
    except (AttributeError, OSError) as exc:
        data["application_process_is_administrator"] = f"unavailable: {exc}"

    try:
        import xlwings

        data["xlwings_version"] = getattr(xlwings, "__version__", "version unknown")
    except ImportError as exc:
        data["xlwings_version"] = f"import failed: {exc}"

    if platform.system() == "Windows":
        try:
            registry_values, registry_errors = _registry_diagnostics()
            data.update(registry_values)
            if registry_errors:
                data["registry_read_errors"] = " | ".join(registry_errors)
        except ImportError as exc:
            data["registry_read_errors"] = f"Windows registry unavailable: {exc}"
    else:
        data["excel_registry"] = "unavailable (not running on Windows)"

    return [f"{key}={value}" for key, value in data.items()]


_log_lock = threading.Lock()
_LOG_MAX_BYTES = 1024 * 1024


def _prepare_log_path(path):
    if os.path.exists(path) and os.path.getsize(path) >= _LOG_MAX_BYTES:
        rotated_path = path + ".1"
        if os.path.exists(rotated_path):
            os.remove(rotated_path)
        os.replace(path, rotated_path)
    return path


def _diagnostics_log_path():
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise OSError("LOCALAPPDATA is not set")

    log_dir = os.path.join(local_app_data, "DeltaOptionChain")
    os.makedirs(log_dir, exist_ok=True)
    return _prepare_log_path(os.path.join(log_dir, "startup_diag.log"))


def write_startup_diagnostics():
    """Append startup details to the current user's writable application data."""
    with _log_lock:
        log_path = _diagnostics_log_path()
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write("\n--- Delta Option Chain startup ---\n")
            log_file.write("\n".join(_diagnostic_lines()))
            log_file.write("\n")
    return log_path


def append_performance_summary(lines):
    """Append aggregate metrics to the bounded per-user diagnostics log."""
    with _log_lock:
        log_path = _diagnostics_log_path()
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write("\n--- Delta Option Chain performance ---\n")
            log_file.write("\n".join(lines))
            log_file.write("\n")
    return log_path


def append_connection_event(asset, status, detail=""):
    """Record only Excel attachment state transitions, never workbook cell data."""
    with _log_lock:
        log_path = _diagnostics_log_path()
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write("\n--- Delta Option Chain Excel connection ---\n")
            log_file.write(f"timestamp={datetime.now().astimezone().isoformat(timespec='seconds')}\n")
            log_file.write(f"asset={asset}\nstatus={status}\n")
            if detail:
                log_file.write(f"detail={detail}\n")
    return log_path


def append_app_exception(context, exc):
    """Persist an actionable exception and traceback in a bounded per-user log."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise OSError("LOCALAPPDATA is not set")

    log_dir = os.path.join(local_app_data, "DeltaOptionChain")
    os.makedirs(log_dir, exist_ok=True)
    rendered_traceback = "".join(
        traceback.format_exception(type(exc), exc, exc.__traceback__)
    )
    with _log_lock:
        log_path = _prepare_log_path(os.path.join(log_dir, "app.log"))
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(
                f"\n--- {datetime.now().astimezone().isoformat(timespec='seconds')} ---\n"
                f"context={context}\n"
                f"{rendered_traceback}"
            )
    return log_path
