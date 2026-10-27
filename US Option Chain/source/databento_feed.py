"""Databento OPRA feed and API-key handling.

The Databento key is never embedded in the application or shipped in the
package. It is resolved in this order:

1. ``DATABENTO_API_KEY`` in the process environment (development/support).
2. ``US Option Chain/.env`` when running from source (development only).
3. ``%LOCALAPPDATA%\\USOptionChain\\databento.key`` (production, entered once
   through the first-run dialog).

The key is stored per Windows user, outside the application folder, so copying
or distributing the application folder does not copy the key.
"""

from __future__ import annotations

import json
import math
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
import tempfile
import threading
import time
import traceback
from datetime import datetime, timedelta, timezone

APP_DIR_NAME = "USOptionChain"
KEY_FILE_NAME = "databento.key"
ENV_VAR = "DATABENTO_API_KEY"
KEY_PREFIX = "db-"

DATASET = "OPRA.PILLAR"
LIVE_SCHEMA = "cbbo-1s"
DEFINITION_LOOKBACK_DAYS = 7
STATISTICS_LOOKBACK_DAYS = 5
DEFAULT_BAND = 120
# Extra strikes subscribed on each side of the displayed band, so the index
# has to travel a long way before a new subscription request is needed.
SUBSCRIBE_MARGIN = 40
MAX_SUBSCRIBED = 12000
SUBSCRIBE_CHUNK = 200
SUBSCRIBE_PAUSE = 0.12
STALE_SECONDS = 15.0
# The gateway is asked for a heartbeat every five seconds. If nothing at all
# arrives for one heartbeat interval plus ten seconds the session is hung and
# has to be dropped and rebuilt.
HEARTBEAT_INTERVAL_S = 5
HUNG_AFTER_SECONDS = HEARTBEAT_INTERVAL_S + 10

RTYPE_CBBO = 192
RTYPE_SYMBOL_MAPPING = 22
RTYPE_ERROR = 21
RTYPE_SYSTEM = 23

PARENT_SUFFIXES = {
    "SPX": "W",
    "NDX": "P",
    "RUT": "W",
    "VIX": "W",
    "DJX": "W",
}


def _base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _key_dir() -> str:
    root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(root, APP_DIR_NAME)


def key_path() -> str:
    return os.path.join(_key_dir(), KEY_FILE_NAME)


def _load_dev_env() -> None:
    """Load ``.env`` from the source or product folder; frozen builds ignore it."""
    if getattr(sys, "frozen", False):
        return
    source_dir = _base_dir()
    paths = [
        os.path.join(source_dir, ".env"),
        os.path.join(os.path.dirname(source_dir), ".env"),
    ]
    lines = None
    for path in dict.fromkeys(paths):
        try:
            with open(path, encoding="utf-8-sig") as fh:
                lines = fh.readlines()
            break
        except OSError:
            continue
    if lines is None:
        return
    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        name, sep, value = line.partition("=")
        name = name.strip()
        if not sep or not name or name in os.environ:
            continue
        value = value.strip().strip('"').strip("'")
        if value:
            os.environ[name] = value


def read_api_key() -> str | None:
    _load_dev_env()
    env = os.environ.get(ENV_VAR, "").strip()
    if env:
        return env
    try:
        with open(key_path(), encoding="utf-8-sig") as fh:
            key = fh.readline().strip()
    except OSError:
        return None
    return key or None


def _hide_file(path: str) -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.kernel32.SetFileAttributesW(path, 0x2)
    except Exception:
        pass


def save_api_key(key: str) -> str:
    key = str(key or "").strip().lstrip("\ufeff")
    if not key.startswith(KEY_PREFIX):
        raise ValueError("Databento keys start with 'db-'.")
    directory = _key_dir()
    os.makedirs(directory, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".databento-", suffix=".tmp", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(key + "\n")
        os.replace(tmp, key_path())
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass
    _hide_file(key_path())
    return key_path()


def mask_key(key: str | None) -> str:
    key = str(key or "")
    if len(key) <= 8:
        return "..." if key else ""
    return f"{key[:4]}...{key[-4:]}"


def check_api_key(key: str) -> tuple[bool, str, bool]:
    """(accepted, detail, rejected)

    rejected is True only when Databento itself said no - HTTP 401 or 403.
    It is False when Databento could not be reached at all, or answered with
    something that is not a verdict on the key.

    The distinction matters. A key that was refused is useless and has to be
    replaced. A key that could not be checked because the connection dropped
    is perfectly good, and throwing it away means a customer with a flaky
    internet connection has to dig out a new key for no reason at all.

    Uses the raw HTTP endpoint rather than the SDK client so the request has a
    hard timeout - the SDK's Historical client does not accept one and its
    default wait is 100s, which would leave the first-run dialog hanging.

    The first request to Databento is often slow: cold DNS, the TLS handshake
    and anything a proxy adds can push a normal request past a short timeout.
    That is a network hiccup, not a bad key, so retry once automatically
    instead of making the customer click Save again.
    """
    try:
        import requests
    except ImportError:
        return False, "requests library is not available.", False
    url = "https://hist.databento.com/v0/metadata.get_dataset_condition"
    for attempt in range(2):
        try:
            resp = requests.get(
                url,
                params={"dataset": DATASET},
                auth=(key, ""),
                timeout=20,
            )
        except Exception as exc:
            if attempt == 0:
                time.sleep(1.5)
                continue
            return False, (
                f"Cannot reach Databento ({type(exc).__name__}). "
                "Check the internet connection and try again."
            ), False
        if resp.status_code == 200:
            return True, "", False
        if resp.status_code in (401, 403):
            return False, "Databento rejected this key. Paste a valid live key.", True
        return False, f"Databento returned HTTP {resp.status_code}. Try again.", False
    return False, "Cannot reach Databento. Check the internet connection and try again.", False


def validate_api_key(key: str) -> tuple[bool, str]:
    """Kept for callers that only want the verdict."""
    accepted, detail, _rejected = check_api_key(key)
    return accepted, detail


def ask_api_key() -> str | None:
    """Clean masked first-run dialog. Must be called from the main thread."""
    try:
        import tkinter as tk
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    result = {"key": None}

    win = tk.Toplevel(root)
    win.title("Databento API Key")
    win.geometry("520x250")
    win.resizable(False, False)
    win.configure(bg="#14181f")
    win.attributes("-topmost", True)

    tk.Label(
        win,
        text="Databento API Key",
        font=("Segoe UI", 13, "bold"),
        fg="#ffffff", bg="#14181f",
    ).pack(pady=(22, 4))
    tk.Label(
        win,
        text="Paste your live key (starts with db-):",
        font=("Segoe UI", 10),
        fg="#c7d2fe", bg="#14181f",
    ).pack(pady=(0, 10))

    entry = tk.Entry(
        win,
        font=("Consolas", 12),
        width=44,
        justify="left",
        relief="flat", bd=0,
        bg="#232b36", fg="#ffffff",
        insertbackground="#ffffff", show="*",
        highlightthickness=1, highlightbackground="#3a4555", highlightcolor="#5a97ff",
    )
    entry.pack(padx=24, ipady=7)
    entry.focus_set()

    show_var = tk.BooleanVar(value=False)

    def toggle_show():
        entry.config(show="" if show_var.get() else "*")

    tk.Checkbutton(
        win,
        text="Show key",
        variable=show_var,
        command=toggle_show,
        font=("Segoe UI", 9),
        fg="#8aa0bf", bg="#14181f",
        selectcolor="#232b36", activebackground="#14181f", activeforeground="#8aa0bf",
    ).pack(anchor="w", padx=24, pady=(8, 0))

    error = tk.Label(win, text="", fg="#ff6b6b", font=("Segoe UI", 9), bg="#14181f")
    error.pack(anchor="w", padx=24, pady=(6, 0))

    def on_ok():
        key = entry.get().strip()
        if not key.startswith(KEY_PREFIX):
            error.config(text="Enter a valid Databento key (starts with db-).")
            return
        save_btn.config(state="disabled", text="Verifying...")
        cancel_btn.config(state="disabled")
        error.config(text="")

        verify_done = threading.Event()
        verify_result = {}

        def _verify():
            try:
                verify_result["ok"], verify_result["detail"] = validate_api_key(key)
            except Exception as exc:
                verify_result["ok"] = False
                verify_result["detail"] = str(exc)[:200]
            finally:
                verify_done.set()

        def _poll():
            if verify_done.is_set():
                if verify_result.get("ok"):
                    result["key"] = key
                    win.destroy()
                else:
                    save_btn.config(state="normal", text="Save")
                    cancel_btn.config(state="normal")
                    error.config(text=verify_result.get("detail") or "Databento rejected this key. Paste a valid key.")
            else:
                win.after(500, lambda: _poll())

        def _timeout():
            if not verify_done.is_set():
                save_btn.config(state="normal", text="Save")
                cancel_btn.config(state="normal")
                error.config(text="Verification timed out. Check your internet and try again.")

        threading.Thread(target=_verify, daemon=True).start()
        win.after(500, _poll)
        # Must outlast the worst case in validate_api_key (two 20s attempts
        # plus a 1.5s pause), or this fires while the retry is still running.
        win.after(50000, _timeout)

    def on_cancel():
        result["key"] = None
        win.destroy()

    btns = tk.Frame(win, bg="#14181f")
    btns.pack(pady=(18, 0))

    save_btn = tk.Button(
        btns, text="Save", width=14, font=("Segoe UI", 10, "bold"),
        bg="#2f7cf6", fg="#ffffff", activebackground="#5a97ff", activeforeground="#ffffff",
        # While the key is being verified the button is disabled, and Tk's
        # default disabledforeground is a dim grey that is hard to read on the
        # blue background. Keep the label white. (Only -disabledforeground is
        # valid on tk.Button here; -disabledbackground raises a TclError.)
        disabledforeground="#ffffff",
        relief="flat", borderwidth=0, cursor="hand2", command=on_ok,
    )
    save_btn.pack(side="left", padx=8)
    save_btn.bind("<Enter>", lambda e: save_btn.config(bg="#5a97ff"))
    save_btn.bind("<Leave>", lambda e: save_btn.config(bg="#2f7cf6"))

    cancel_btn = tk.Button(
        btns, text="Cancel", width=14, font=("Segoe UI", 10),
        bg="#232b36", fg="#ffffff", activebackground="#2e3947", activeforeground="#ffffff",
        disabledforeground="#c3ccdb",
        relief="flat", borderwidth=0, cursor="hand2", command=on_cancel,
    )
    cancel_btn.pack(side="left", padx=8)
    cancel_btn.bind("<Enter>", lambda e: cancel_btn.config(bg="#2e3947"))
    cancel_btn.bind("<Leave>", lambda e: cancel_btn.config(bg="#232b36"))

    win.bind("<Return>", lambda e: on_ok())
    win.bind("<Escape>", lambda e: on_cancel())
    win.protocol("WM_DELETE_WINDOW", on_cancel)

    win.update_idletasks()
    # Center on the primary display. winfo_screenwidth() reports the primary
    # monitor only, but Windows may still restore a remembered position on a
    # secondary monitor, where the customer never sees the dialog. Force the
    # position and pull the window to the front so it cannot hide.
    try:
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
    except Exception:
        sw, sh = 1920, 1080
    x = max(0, (sw - 520) // 2)
    y = max(0, (sh - 250) // 2)
    win.geometry(f"+{x}+{y}")
    try:
        win.lift()
        win.focus_force()
    except Exception:
        pass
    entry.focus_set()

    root.wait_window(win)
    try:
        root.destroy()
    except Exception:
        pass
    return result["key"]


def ensure_api_key() -> str | None:
    """Return a key from env/.env/key file, prompting once if missing."""
    key = read_api_key()
    if key:
        return key
    key = ask_api_key()
    if not key:
        return None
    save_api_key(key)
    return key


def clear_saved_key() -> None:
    """Forget the key stored on disk so the next ensure asks for a new one.

    This fires when the saved key is rejected by Databento's API (expired or
    revoked). Without it, ensure_api_key would hand the same dead key back
    on every call and the customer would never be asked again.
    """
    path = key_path()
    try:
        if sys.platform == "win32":
            import ctypes

            ctypes.windll.kernel32.SetFileAttributesW(path, 0x80)
        os.remove(path)
    except FileNotFoundError:
        pass
    except OSError:
        try:
            with open(path, "w", encoding="utf-8", newline="\n") as fh:
                fh.write("\n")
        except OSError:
            pass


def _db():
    import databento

    return databento


def _to_float(value, scale_ns: bool = False) -> float:
    try:
        if int(value) == 9223372036854775807:
            return 0.0
    except (TypeError, ValueError, OverflowError):
        pass
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0.0
    if math.isnan(number):
        return 0.0
    if number == 9223372036854775807:
        return 0.0
    if scale_ns and abs(number) > 1_000_000:
        number /= 1_000_000_000.0
    return number


def _price(value) -> float:
    number = _to_float(value)
    if abs(number) > 1_000_000:
        number /= 1_000_000_000.0
    return number if number > 0 else 0.0


def _strike(value) -> float:
    number = _to_float(value)
    if abs(number) > 1_000_000:
        number /= 1_000_000_000.0
    return round(number, 3)


def _timestamp(value) -> float:
    if value is None:
        return 0.0
    try:
        if int(value) == 18446744073709551615:
            return 0.0
    except (TypeError, ValueError, OverflowError):
        pass
    if hasattr(value, "timestamp"):
        try:
            return float(value.timestamp())
        except (OverflowError, OSError, ValueError):
            return 0.0
    number = _to_float(value)
    if number > 1_000_000_000_000:
        number /= 1_000_000_000.0
    return number


def _et_now():
    try:
        from us_time import eastern_now

        return eastern_now()[0]
    except Exception:
        return datetime.now()


def _et_today() -> str:
    return _et_now().strftime("%Y-%m-%d")


def _is_rth(now=None) -> bool:
    now = now or _et_now()
    if now.weekday() >= 5:
        return False
    minutes = now.hour * 60 + now.minute
    return 9 * 60 + 30 <= minutes < 16 * 60


def _norm(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def _root_matches(symbol: str, root: str) -> bool:
    user = _norm(symbol)
    got = _norm(root)
    if not user or not got:
        return False
    if got == user:
        return True
    suffix = PARENT_SUFFIXES.get(user)
    return bool(suffix and got == user + suffix)


def _root_of(raw_symbol: str) -> str:
    return str(raw_symbol or "")[:6].strip().upper()


def _option_type(value) -> str | None:
    text = str(value or "").upper()
    if text.endswith("C") or "CALL" in text:
        return "call"
    if text.endswith("P") or "PUT" in text:
        return "put"
    return None


def _definition_chunks(start, end):
    """Split a definition request into one calendar day per piece.

    The definition feed repeats an instrument's record every day it is live,
    so a multi-day window multiplies the payload for no extra information and
    the gateway answers a large request with "504 The remote gateway timed
    out". One day per request keeps every request small.
    """
    chunks = []
    cursor = start
    while cursor < end:
        midnight = cursor.replace(hour=0, minute=0, second=0, microsecond=0)
        nxt = midnight + timedelta(days=1)
        if nxt > end:
            chunks.append((cursor, end))
            break
        chunks.append((cursor, nxt))
        cursor = nxt
    return chunks or [(start, end)]


def _expiry_iso(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
        if text:
            try:
                return datetime.fromisoformat(text.replace("Z", "+00:00")).date().isoformat()
            except ValueError:
                pass
    if hasattr(value, "date") and not isinstance(value, (int, float)):
        try:
            return value.date().isoformat()
        except (AttributeError, ValueError):
            pass
    if hasattr(value, "isoformat") and not isinstance(value, (int, float)):
        try:
            return str(value.isoformat())[:10]
        except (AttributeError, ValueError):
            pass
    number = _to_float(value)
    if number <= 0:
        return None
    try:
        return datetime.fromtimestamp(number / 1_000_000_000.0, tz=timezone.utc).date().isoformat()
    except (OverflowError, OSError, ValueError):
        return None


def _is_auth_failure(exc: Exception) -> bool:
    """True when the failure is the key being rejected, not a transient fault.

    Databento's SDK raises BentoHttpError with the status attached; a raw
    request raises requests.HTTPError with the response attached. Both carry
    401 when the key is wrong and 403 when the key is valid but not entitled
    to the data. Neither gets better by being tried again.
    """
    status = getattr(exc, "http_status", None)
    if isinstance(status, int) and status in (401, 403):
        return True
    response = getattr(exc, "response", None)
    if response is not None:
        code = getattr(response, "status_code", None)
        if isinstance(code, int) and code in (401, 403):
            return True
    text = str(exc).lower()
    return (
        "authentication" in text
        or "unauthorized" in text
        or "api key" in text
    )


class DatabentoFeed:
    def __init__(self, key: str, band: int = DEFAULT_BAND, event_log=None):
        self.key = key
        self.band = max(1, int(band))
        self._event_log = event_log
        self._lock = threading.RLock()
        self._load_lock = threading.Lock()
        self._sub_lock = threading.RLock()
        self._defs: dict[str, list[dict]] = {}
        self._quotes: dict[str, dict] = {}
        self._id_map: dict[int, str] = {}
        self._errors: list[str] = []
        # Set when the key itself was rejected. A rejected key does not fix
        # itself, so the historical feed is latched off instead of being
        # retried every few seconds for as long as the app is open.
        self._auth_error: str | None = None
        self._loading: set[str] = set()
        self._universe_failures: dict[str, int] = {}
        self._universe_retry_at: dict[str, float] = {}
        self._oi_loading: set[str] = set()
        self._parents_by_symbol: dict[str, list[str]] = {}
        self._live = None
        self._thread = None
        self._stop = threading.Event()
        self._started = threading.Event()
        self._subscribed: set[str] = set()
        self._target: set[str] = set()
        self._last_record = 0.0
        # Any message at all, including the gateway's heartbeats. The live API
        # sends a heartbeat when no data is flowing, so this is what proves
        # the session is alive even when the market is quiet.
        self._last_message = 0.0
        self._last_connect = 0.0
        self._last_reconnect = 0.0
        self._unmapped = 0
        self._session_symbols = 0
        self._cache_dir = _key_dir()

    def _redact(self, value) -> str:
        text = str(value)
        if self.key:
            text = text.replace(self.key, "<key>")
        return text

    def _log(self, event: str, **fields) -> None:
        safe = {key: self._redact(value) for key, value in fields.items()}
        if self._event_log is not None:
            try:
                self._event_log(event, **safe)
                return
            except Exception:
                pass
        try:
            body = " ".join(f"{k}={v}" for k, v in safe.items())
            print(f"[databento] {event} {body}".strip(), flush=True)
        except Exception:
            pass

    def _record_error(self, message: str) -> None:
        with self._lock:
            self._errors.append(self._redact(message)[:300])
            self._errors = self._errors[-5:]

    def auth_error(self) -> str | None:
        """The reason the key was rejected, or None when it was accepted."""
        with self._lock:
            # Read defensively: this is called from the render path, and a
            # missing attribute must not be able to take the render pass down.
            return getattr(self, "_auth_error", None)

    def _set_auth_error(self, symbol: str, exc: Exception) -> None:
        """Latch the feed off: the key was rejected, so stop asking.

        This is written once. The old code treated a 401 like a hiccup and
        retried it three times per chunk, for every parent, for every symbol,
        and then started the whole thing again with a five-to-one-hundred-and-
        twenty-second backoff. Seventeen minutes of the log was nothing but
        401s, and the customer saw a blank sheet with no explanation.
        """
        detail = self._redact(str(exc))[:200]
        with self._lock:
            if self._auth_error is not None:
                return
            self._auth_error = detail or "the key was rejected"
        self._log(
            "databento_auth_failed",
            symbol=symbol,
            error=type(exc).__name__,
            detail=detail,
            note="the historical feed is now off until the key changes",
        )
        self._record_error(f"Databento rejected the key: {detail}")

    def errors(self) -> list[str]:
        with self._lock:
            return list(self._errors)

    def ready(self, symbol: str) -> bool:
        with self._lock:
            return bool(self._defs.get(symbol))

    def is_started(self) -> bool:
        return self._started.is_set()

    def connected(self) -> bool:
        live = self._live
        if live is None:
            return False
        try:
            return bool(live.is_connected())
        except Exception:
            return False

    def last_record_age(self) -> float:
        with self._lock:
            last = self._last_record
        if not last:
            return float("inf")
        return max(0.0, time.monotonic() - last)

    def last_message_age(self) -> float:
        """How long since the gateway last sent anything, heartbeats included."""
        with self._lock:
            last = self._last_message
        if not last:
            return float("inf")
        return max(0.0, time.monotonic() - last)

    def healthy(self) -> bool:
        if not self.connected():
            return False
        return self.last_record_age() <= STALE_SECONDS

    def status_text(self) -> str:
        try:
            rejected = self.auth_error()
        except Exception:
            rejected = None
        if rejected:
            # The key was rejected. Say so, plainly, instead of leaving the
            # customer to stare at a sheet that says LOADING forever.
            return "DATABENTO KEY REJECTED"
        if not self.is_started():
            if self._defs:
                return "DATABENTO CONNECTING"
            return "DATABENTO LOADING"
        if not self.connected():
            return "DATABENTO RECONNECTING"
        age = self.last_record_age()
        if age > STALE_SECONDS:
            # last_record_age() is infinity when no record has ever arrived.
            # int(infinity) raises OverflowError, which would take the whole
            # render pass down, so format it safely.
            shown = "no data yet" if not math.isfinite(age) else f"{int(age)}s"
            if _is_rth():
                return f"DATABENTO STALE {shown}"
            # Outside trading hours the exchange stops sending quotes for
            # these contracts. That is not a fault, and the sheet is still
            # worth reading (strikes and open interest are real), so say
            # plainly that the prices are not live.
            return f"NO LIVE QUOTES ({shown})"
        return "LIVE"

    def _cache_path(self, symbol: str) -> str:
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", str(symbol).upper())
        return os.path.join(self._cache_dir, f"databento_{safe}.json")

    def _load_cache(self, symbol: str) -> list[dict]:
        try:
            with open(self._cache_path(symbol), encoding="utf-8") as fh:
                payload = json.load(fh)
        except (OSError, ValueError):
            return []
        if payload.get("date") != _et_today():
            return []
        rows = payload.get("rows") or []
        if not isinstance(rows, list) or not rows:
            return []
        return [row for row in rows if self._valid_cached_row(row)]

    @staticmethod
    def _valid_cached_row(row) -> bool:
        return (
            isinstance(row, dict)
            and isinstance(row.get("r"), str)
            and isinstance(row.get("e"), str)
            and isinstance(row.get("t"), str)
            and row.get("root")
            and _to_float(row.get("k")) > 0
        )

    def _save_cache(self, symbol: str, rows: list[dict]) -> None:
        try:
            os.makedirs(self._cache_dir, exist_ok=True)
            path = self._cache_path(symbol)
            fd, tmp = tempfile.mkstemp(prefix=".databento-", suffix=".tmp", dir=self._cache_dir)
            try:
                with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as fh:
                    json.dump({"date": _et_today(), "rows": rows}, fh, separators=(",", ":"))
                os.replace(tmp, path)
            finally:
                try:
                    os.unlink(tmp)
                except OSError:
                    pass
        except Exception as exc:
            self._log("databento_cache_error", error=type(exc).__name__, detail=str(exc)[:160])

    def ensure_universe(self, symbol: str, wait: bool = False) -> bool:
        with self._lock:
            if self._defs.get(symbol):
                return True
            if self._auth_error:
                # The key was rejected. Do not ask again: there is nothing to
                # wait for and nothing to retry.
                return False
            if self._universe_retry_at.get(symbol, 0.0) > time.monotonic():
                return False
            if symbol in self._loading:
                loading = True
            else:
                self._loading.add(symbol)
                loading = False
        if loading:
            if not wait:
                return False
            deadline = time.monotonic() + 5.0
            while time.monotonic() < deadline:
                with self._lock:
                    if self._defs.get(symbol):
                        return True
                time.sleep(0.2)
            return False
        try:
            cached = self._load_cache(symbol)
            rows = cached or self._fetch_universe(symbol)
            if rows:
                with self._lock:
                    self._defs[symbol] = rows
                    parents = list(self._parents_by_symbol.get(symbol) or [])
                    self._universe_failures.pop(symbol, None)
                    self._universe_retry_at.pop(symbol, None)
                if not any(row.get("oi", 0) for row in rows):
                    self._start_oi_refresh(
                        symbol, parents or self._parent_candidates(symbol)
                    )
                return True
            with self._lock:
                failures = self._universe_failures.get(symbol, 0) + 1
                self._universe_failures[symbol] = failures
                delay = min(120.0, 5.0 * (2 ** (failures - 1)))
                self._universe_retry_at[symbol] = time.monotonic() + delay
            return False
        finally:
            with self._lock:
                self._loading.discard(symbol)

    def _parent_candidates(self, symbol: str) -> list[str]:
        base = str(symbol or "").strip().upper().lstrip("^")
        names = [base]
        if "." in base:
            names.extend([base.replace(".", " "), base.replace(".", "-")])
        if "-" in base:
            names.extend([base.replace("-", " "), base.replace("-", ".")])
        out = []
        for name in names:
            candidate = f"{name}.OPT"
            if candidate not in out:
                out.append(candidate)
        suffix = PARENT_SUFFIXES.get(base)
        if suffix:
            candidate = f"{base}{suffix}.OPT"
            if candidate not in out:
                out.append(candidate)
        return out

    def _history_range(self, days: int, history=None):
        end = datetime.now(timezone.utc)
        if history is not None:
            try:
                info = history.metadata.get_dataset_range(dataset=DATASET)
                text = str(info.get("end") or "")
                available = datetime.strptime(
                    text[:19], "%Y-%m-%dT%H:%M:%S"
                ).replace(tzinfo=timezone.utc)
                if available < end:
                    end = available - timedelta(seconds=1)
            except Exception:
                pass
        return end - timedelta(days=days), end

    def _fetch_universe(self, symbol: str) -> list[dict]:
        try:
            db = _db()
        except ImportError:
            self._record_error("Databento client is not installed.")
            return []
        try:
            history = db.Historical(self.key)
        except Exception as exc:
            if _is_auth_failure(exc):
                self._set_auth_error(symbol, exc)
                return []
            self._record_error(f"Databento init: {type(exc).__name__}: {exc}")
            return []

        start, end = self._history_range(DEFINITION_LOOKBACK_DAYS, history)
        start = start.replace(hour=0, minute=0, second=0, microsecond=0)
        by_raw: dict[str, dict] = {}
        used_parents: list[str] = []
        parents = self._parent_candidates(symbol)

        def fetch_parent(parent):
            # Ask for one calendar day at a time and merge the pieces.
            #
            # The definition feed repeats an instrument's record on every day
            # it is live, so a multi-day window returns the same instruments
            # over and over. A two-day window is already too much for the
            # gateway and comes back as "504 The remote gateway timed out" -
            # which is what lost the weekly chains: SPXW never arrived, so the
            # weekly expiries never reached the dropdown and the customer only
            # ever saw the far-away monthly expiries.
            #
            # One day per request stays well under the limit. Measured on this
            # account: two days in one request fails, one day per request
            # returns the full list.
            frames = []
            chunks = _definition_chunks(start, end)

            def fetch_chunk(bounds):
                chunk_start, chunk_end = bounds
                last = None
                for attempt in range(3):
                    try:
                        client = db.Historical(self.key)
                        store = client.timeseries.get_range(
                            dataset=DATASET,
                            schema="definition",
                            symbols=[parent],
                            stype_in="parent",
                            stype_out="instrument_id",
                            start=chunk_start,
                            end=chunk_end,
                        )
                        return store.to_df(map_symbols=True)
                    except Exception as exc:
                        last = exc
                        if _is_auth_failure(exc):
                            # A rejected key will still be rejected in three
                            # seconds. Stop the retries here.
                            raise
                        if attempt < 2:
                            time.sleep(1.5 * (attempt + 1))
                raise last

            if len(chunks) == 1:
                frames.append(fetch_chunk(chunks[0]))
            else:
                with ThreadPoolExecutor(max_workers=len(chunks)) as pool:
                    for frame in pool.map(fetch_chunk, chunks):
                        frames.append(frame)
            if not frames:
                return None
            if len(frames) == 1:
                return frames[0]
            import pandas as pd
            return pd.concat(frames, ignore_index=True)

        with ThreadPoolExecutor(max_workers=max(1, len(parents))) as pool:
            futures = [(parent, pool.submit(fetch_parent, parent)) for parent in parents]
            for parent, future in futures:
                try:
                    frame = future.result()
                except Exception as exc:
                    if _is_auth_failure(exc):
                        # The key was rejected. Latch the feed off instead of
                        # logging the same failure once per parent.
                        self._set_auth_error(symbol, exc)
                        return []
                    self._log(
                        "databento_definition_failed",
                        symbol=symbol,
                        parent=parent,
                        error=type(exc).__name__,
                        detail=str(exc)[:180],
                    )
                    continue
                found = self._definitions_from_frame(frame, symbol)
                if not found:
                    continue
                used_parents.append(parent)
                for row in found:
                    by_raw.setdefault(row["r"], row)
        rows = list(by_raw.values())
        if not rows:
            self._log("databento_definition_empty", symbol=symbol,
                      parents=",".join(self._parent_candidates(symbol)))
            return []

        for row in rows:
            row["oi"] = 0
        with self._lock:
            self._parents_by_symbol[symbol] = list(used_parents)
        self._save_cache(symbol, rows)
        self._log(
            "databento_universe",
            symbol=symbol,
            parents=",".join(used_parents),
            contracts=len(rows),
            expiries=len({row["e"] for row in rows}),
        )
        return rows

    def _definitions_from_frame(self, frame, symbol: str) -> list[dict]:
        if frame is None or getattr(frame, "empty", True):
            return []
        rows = []
        seen = set()
        for record in frame.to_dict("records"):
            raw = record.get("symbol")
            if not isinstance(raw, str) or not raw.strip():
                raw = record.get("raw_symbol")
            if not isinstance(raw, str) or not raw.strip():
                continue
            raw = raw.strip()
            root = _root_of(raw) or str(record.get("asset") or "").strip().upper()
            if not _root_matches(symbol, root):
                continue
            option_type = _option_type(record.get("instrument_class"))
            if option_type is None:
                continue
            expiry = _expiry_iso(record.get("expiration"))
            strike = _strike(record.get("strike_price"))
            if not expiry or strike <= 0:
                continue
            if expiry < _et_today():
                continue
            if raw in seen:
                continue
            seen.add(raw)
            rows.append({"r": raw, "root": root, "e": expiry, "k": strike, "t": option_type})
        return rows

    def _start_oi_refresh(self, symbol: str, parents: list[str]) -> None:
        with self._lock:
            if symbol in self._oi_loading:
                return
            self._oi_loading.add(symbol)
        try:
            threading.Thread(
                target=self._refresh_open_interest,
                args=(symbol, parents),
                name=f"databento-oi-{symbol}",
                daemon=True,
            ).start()
        except Exception:
            with self._lock:
                self._oi_loading.discard(symbol)

    def _refresh_open_interest(self, symbol: str, parents: list[str]) -> None:
        try:
            try:
                db = _db()
                history = db.Historical(self.key)
            except Exception as exc:
                self._record_error(
                    f"Databento OI init: {type(exc).__name__}: {exc}"
                )
                return
            oi = self._fetch_open_interest(history, symbol, parents)
            if not oi:
                return
            with self._lock:
                rows = list(self._defs.get(symbol) or [])
                updated = 0
                for row in rows:
                    value = oi.get(row["r"])
                    if value is not None:
                        row["oi"] = int(value)
                        updated += 1
                cached = [dict(row) for row in rows]
            if updated:
                self._save_cache(symbol, cached)
                self._log(
                    "databento_oi",
                    symbol=symbol,
                    parents=",".join(parents),
                    updated=updated,
                    contracts=len(cached),
                )
        except Exception as exc:
            self._record_error(f"Databento OI: {type(exc).__name__}: {exc}")
        finally:
            with self._lock:
                self._oi_loading.discard(symbol)

    def _fetch_open_interest(
        self, history, symbol: str, parents: list[str]
    ) -> dict[str, int]:
        start, end = self._history_range(STATISTICS_LOOKBACK_DAYS, history)
        latest: dict[str, tuple[float, int]] = {}
        for parent in parents:
            try:
                store = history.timeseries.get_range(
                    dataset=DATASET,
                    schema="statistics",
                    symbols=[parent],
                    stype_in="parent",
                    stype_out="instrument_id",
                    start=start,
                    end=end,
                )
                frame = store.to_df(map_symbols=True)
            except Exception as exc:
                self._log(
                    "databento_oi_failed",
                    symbol=symbol,
                    parent=parent,
                    error=type(exc).__name__,
                    detail=str(exc)[:180],
                )
                continue
            if frame is None or getattr(frame, "empty", True):
                continue
            symbol_col = "symbol" if "symbol" in frame.columns else "raw_symbol"
            if symbol_col not in frame.columns:
                continue
            stat_text = frame["stat_type"].astype(str).str.upper()
            mask = stat_text.str.contains("OPEN_INTEREST", na=False) | (
                stat_text == "9"
            )
            if not mask.any():
                continue
            view = frame.loc[mask, [symbol_col, "quantity", "ts_ref"]].copy()
            view = view[view[symbol_col].astype(str).str.len() > 0]
            if view.empty:
                continue
            if "ts_ref" in view.columns:
                view = view.sort_values("ts_ref")
            view = view.drop_duplicates(subset=[symbol_col], keep="last")
            for record in view.to_dict("records"):
                raw = record.get(symbol_col)
                if not isinstance(raw, str) or not raw:
                    continue
                quantity = int(_to_float(record.get("quantity")))
                stamp = _timestamp(record.get("ts_ref"))
                previous = latest.get(raw)
                if previous is None or stamp >= previous[0]:
                    latest[raw] = (stamp, quantity)
        return {raw: value[1] for raw, value in latest.items()}

    def _default_expiries(self, symbol: str, rows: list[dict]) -> list[str]:
        expiries = sorted({row["e"] for row in rows})
        if not expiries:
            return []
        base = _norm(symbol)
        roots_by_expiry: dict[str, set[str]] = {}
        with_oi: set[str] = set()
        for row in rows:
            expiry = row["e"]
            roots_by_expiry.setdefault(expiry, set()).add(_norm(row["root"]))
            if row.get("oi", 0) > 0:
                with_oi.add(expiry)
        standard = [e for e in expiries if base in roots_by_expiry.get(e, ())]
        weekly = [e for e in expiries if e not in set(standard)]
        standard.sort(key=lambda expiry: (0 if expiry in with_oi else 1, expiry))
        return (standard + weekly)[:3]

    def _target_symbols(self, symbol: str, spot: float | None, active_expiries) -> set[str]:
        rows = self._defs.get(symbol) or []
        chosen = [expiry for expiry in (active_expiries or []) if expiry]
        if not chosen:
            chosen = self._default_expiries(symbol, rows)
        by_expiry: dict[str, list[dict]] = {}
        for row in rows:
            by_expiry.setdefault(row["e"], []).append(row)
        targets: set[str] = set()
        for expiry in chosen:
            defs = [
                row
                for row in by_expiry.get(expiry, ())
                if _root_matches(symbol, row["root"])
            ]
            if not defs:
                continue
            strikes = sorted({row["k"] for row in defs})
            if spot and spot > 0:
                center = min(range(len(strikes)), key=lambda i: abs(strikes[i] - spot))
            else:
                center = len(strikes) // 2
            # Subscribe wider than the sheet shows.
            #
            # The gateway throttles subscription requests to ten a second, and
            # every request carries a chunk of symbols. If the subscribed band
            # were exactly the displayed band, every strike the index moved
            # would ask for another symbol. The margin means the index has to
            # travel a long way before a new subscription is needed at all.
            low = max(0, center - self.band - SUBSCRIBE_MARGIN)
            high = min(len(strikes), center + self.band + SUBSCRIBE_MARGIN + 1)
            wanted = set(strikes[low:high])
            targets.update(row["r"] for row in defs if row["k"] in wanted)
        return targets

    def _subscribe(self, live, symbols: set[str]) -> None:
        missing = [raw for raw in sorted(symbols) if raw not in self._subscribed]
        room = MAX_SUBSCRIBED - len(self._subscribed)
        if room <= 0:
            self._log("databento_subscription_cap", total=len(self._subscribed))
            return
        if len(missing) > room:
            self._log(
                "databento_subscription_truncated",
                requested=len(missing),
                room=room,
            )
            missing = missing[:room]
        if not missing:
            return
        for start in range(0, len(missing), SUBSCRIBE_CHUNK):
            chunk = missing[start:start + SUBSCRIBE_CHUNK]
            if not chunk:
                continue
            live.subscribe(
                dataset=DATASET,
                schema=LIVE_SCHEMA,
                symbols=chunk,
                stype_in="raw_symbol",
            )
            self._subscribed.update(chunk)
            time.sleep(SUBSCRIBE_PAUSE)
        self._log(
            "databento_subscribe",
            symbols=self._session_symbols,
            added=len(missing),
            total=len(self._subscribed),
        )

    def _on_record(self, record) -> None:
        try:
            rtype = int(getattr(record, "rtype", 0))
            if rtype == RTYPE_CBBO:
                instrument_id = int(getattr(record, "instrument_id", 0))
                raw = None
                live = self._live
                if live is not None:
                    try:
                        mapped = live.symbology_map.get(instrument_id)
                        if isinstance(mapped, str):
                            raw = mapped
                    except Exception:
                        raw = None
                bid = _price(getattr(record, "bid_px_00", 0))
                ask = _price(getattr(record, "ask_px_00", 0))
                if bid > 0 and ask > 0 and ask < bid:
                    bid = ask = 0.0
                price = _price(getattr(record, "price", 0))
                size = int(_to_float(getattr(record, "size", 0)))
                stamp = _timestamp(getattr(record, "ts_recv", 0))
                with self._lock:
                    self._last_message = time.monotonic()
                    if raw is None:
                        raw = self._id_map.get(instrument_id)
                    if raw:
                        self._quotes[raw] = {
                            "bid": bid,
                            "ask": ask,
                            "ltp": price,
                            "size": size,
                            "ts": stamp,
                        }
                        self._last_record = time.monotonic()
                    else:
                        self._unmapped += 1
                return
            if rtype == RTYPE_SYMBOL_MAPPING:
                instrument_id = int(getattr(record, "instrument_id", 0))
                raw = getattr(record, "stype_in_symbol", None) or getattr(record, "stype_out_symbol", None)
                if isinstance(raw, str) and raw:
                    with self._lock:
                        self._id_map[instrument_id] = raw
                return
            if rtype == RTYPE_ERROR:
                code = str(getattr(record, "code", ""))
                detail = str(getattr(record, "err", ""))[:200]
                self._record_error(f"{code}: {detail}")
                self._log("databento_error", code=code, detail=detail)
                return
            if rtype == RTYPE_SYSTEM:
                with self._lock:
                    # A heartbeat is proof the session is alive even when no
                    # data is flowing.
                    self._last_connect = time.monotonic()
                    self._last_message = time.monotonic()
                return
        except Exception as exc:
            self._log("databento_record_error", error=type(exc).__name__, detail=str(exc)[:160])

    def _on_exception(self, exc: Exception) -> None:
        self._record_error(f"{type(exc).__name__}: {exc}")
        self._log("databento_callback_error", error=type(exc).__name__, detail=str(exc)[:180])

    def _on_reconnect(self, last_ts, new_start) -> None:
        self._last_reconnect = time.monotonic()
        self._log("databento_reconnect", last_ts=str(last_ts), new_start=str(new_start))
        # The gateway reconnects the session but the subscription does not
        # survive it, so the symbols have to be asked for again. Without this
        # the feed goes silent after any network blip, and stays silent until
        # the whole session is torn down and rebuilt.
        live = self._live
        if live is None:
            return
        with self._lock:
            target = set(self._target)
        if not target:
            return
        threading.Thread(
            target=self._resubscribe_after_reconnect,
            args=(live, target),
            daemon=True,
        ).start()

    def _resubscribe_after_reconnect(self, live, target) -> None:
        # The reconnect callback cannot hold _sub_lock while it steps
        # through the subscribe chunks: snapshot polling also takes that
        # lock, and one long resubscribe would stall every poller until it
        # finished.
        with self._sub_lock:
            self._subscribed.clear()
            try:
                self._subscribe(live, target)
            except Exception as exc:
                self._record_error(f"{type(exc).__name__}: {exc}")
                self._log(
                    "databento_resubscribe_failed",
                    error=type(exc).__name__,
                    detail=str(exc)[:160],
                )

    def _watch_hung_connection(self, live) -> None:
        """Drop a session the gateway has gone quiet on.

        The live API sends a heartbeat when nothing else is flowing. If nothing
        at all arrives for one heartbeat interval plus ten seconds the session
        is hung and has to be dropped, so the session loop can build a fresh
        one. Without this a network change that leaves the TCP session
        half-open leaves the chain frozen on its last quote for good.
        """
        while not self._stop.is_set():
            if self._stop.wait(3.0):
                return
            if self._live is not live:
                return
            age = self.last_message_age()
            if age == float("inf") or age > HUNG_AFTER_SECONDS:
                self._log(
                    "databento_hung",
                    age_s=(None if age == float("inf") else int(age)),
                )
                try:
                    live.terminate()
                except Exception:
                    pass
                return

    def _open_session(self, target: set[str]) -> None:
        db = _db()
        live = db.Live(
            self.key,
            ts_out=True,
            reconnect_policy="reconnect",
            compression=db.Compression.ZSTD,
            heartbeat_interval_s=HEARTBEAT_INTERVAL_S,
        )
        live.add_callback(self._on_record, self._on_exception)
        live.add_reconnect_callback(self._on_reconnect, self._on_exception)
        self._live = live
        with self._lock:
            self._id_map.clear()
        self._session_symbols = len(target)
        with self._sub_lock:
            self._subscribe(live, target)
        live.start()
        self._last_connect = time.monotonic()
        self._started.set()
        # Watch for a session that has gone quiet and drop it so a fresh one
        # is built. The gateway heartbeats every five seconds, so fifteen
        # seconds of total silence means the session is hung.
        threading.Thread(
            target=self._watch_hung_connection,
            args=(live,),
            daemon=True,
        ).start()
        self._log("databento_connected", symbols=len(self._subscribed))
        try:
            live.block_for_close()
        finally:
            self._started.clear()
            self._live = None
            with self._sub_lock:
                self._subscribed.clear()
            self._log("databento_disconnected", symbols=len(self._session_symbols))

    def _run(self) -> None:
        backoff = 2.0
        while not self._stop.is_set():
            with self._lock:
                target = set(self._target)
            if not target:
                time.sleep(0.5)
                continue
            try:
                self._open_session(target)
                backoff = 2.0
            except Exception as exc:
                if not self._stop.is_set():
                    self._record_error(f"{type(exc).__name__}: {exc}")
                    self._log(
                        "databento_session_failed",
                        error=type(exc).__name__,
                        detail=traceback.format_exc()[-600:],
                    )
                self._started.clear()
                self._live = None
                with self._sub_lock:
                    self._subscribed.clear()
            if self._stop.is_set():
                break
            self._stop.wait(backoff)
            backoff = min(30.0, backoff * 2.0)

    def ensure_started(self) -> None:
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return
            self._stop.clear()
            self._thread = threading.Thread(
                target=self._run, name="databento-live", daemon=True
            )
            self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        live = self._live
        if live is not None:
            try:
                live.terminate()
            except Exception:
                pass
        thread = self._thread
        if thread is not None:
            thread.join(timeout=5.0)
        self._thread = None
        self._started.clear()
        self._live = None

    def snapshot(
        self,
        symbol: str,
        spot: float | None,
        active_expiries=None,
        wait: bool = False,
    ):
        if self._stop.is_set():
            return None
        if not self.ensure_universe(symbol, wait=wait):
            return None
        rows = self._defs.get(symbol) or []
        if not rows:
            return None
        expiries = sorted({row["e"] for row in rows})
        active = [expiry for expiry in (active_expiries or []) if expiry]
        if not active:
            active = self._default_expiries(symbol, rows)
        active = active[:3]
        targets = self._target_symbols(symbol, spot, active)
        if targets:
            with self._lock:
                self._target.update(targets)
                stale = [raw for raw in self._quotes if raw not in self._target]
                for raw in stale:
                    self._quotes.pop(raw, None)
            with self._sub_lock:
                self._started.wait(timeout=0.2)
                live = self._live
                if live is not None and self._started.is_set():
                    try:
                        self._subscribe(live, targets)
                    except Exception as exc:
                        self._record_error(f"subscribe: {type(exc).__name__}: {exc}")
                        self._log("databento_subscribe_failed", error=type(exc).__name__,
                                  detail=str(exc)[:160])
            self.ensure_started()

        built = self._build_rows(symbol, rows, spot, active)
        return spot, built, expiries, "databento"

    def _build_rows(self, symbol: str, rows: list[dict], spot: float | None,
                    active: list[str]) -> list[dict]:
        with self._lock:
            quotes = self._quotes.copy()
        by_expiry: dict[str, list[dict]] = {}
        roots_by_expiry: dict[str, set[str]] = {}
        oi_roots: set[tuple[str, str]] = set()
        for row in rows:
            expiry = row["e"]
            root = row["root"]
            by_expiry.setdefault(expiry, []).append(row)
            roots_by_expiry.setdefault(expiry, set()).add(root)
            if row.get("oi", 0) > 0:
                oi_roots.add((expiry, root))

        out = []
        active_set = set(active)
        for expiry in active:
            defs = [
                row
                for row in by_expiry.get(expiry, ())
                if _root_matches(symbol, row["root"])
            ]
            if not defs:
                continue
            strikes = sorted({row["k"] for row in defs})
            if spot and spot > 0:
                center = min(range(len(strikes)), key=lambda i: abs(strikes[i] - spot))
            else:
                center = len(strikes) // 2
            low = max(0, center - self.band)
            high = min(len(strikes), center + self.band + 1)
            wanted = set(strikes[low:high])
            for row in defs:
                if row["k"] not in wanted:
                    continue
                quote = quotes.get(row["r"]) or {}
                bid = _to_float(quote.get("bid"))
                ask = _to_float(quote.get("ask"))
                ltp = _to_float(quote.get("ltp"))
                if ltp <= 0:
                    if bid > 0 and ask > 0:
                        ltp = (bid + ask) / 2.0
                    elif bid > 0:
                        ltp = bid
                out.append({
                    "contract": row["r"],
                    "root": row["root"],
                    "strike": row["k"],
                    "type": row["t"],
                    "ltp": ltp,
                    "bid": bid,
                    "ask": ask,
                    "oi": row.get("oi", 0),
                    "volume": 0,
                    "iv": 0,
                    "delta": 0,
                    "expiry": expiry,
                })
        for expiry in sorted(set(by_expiry) - active_set):
            for root in sorted(roots_by_expiry.get(expiry, ())):
                has_oi = (expiry, root) in oi_roots
                out.append({
                    "contract": f"__meta__{expiry}_{root}",
                    "root": root,
                    "strike": 0.0,
                    "type": "call",
                    "ltp": 0.0,
                    "bid": 0.0,
                    "ask": 0.0,
                    "oi": 1 if has_oi else 0,
                    "volume": 0,
                    "iv": 0,
                    "delta": 0,
                    "expiry": expiry,
                })
        return out
