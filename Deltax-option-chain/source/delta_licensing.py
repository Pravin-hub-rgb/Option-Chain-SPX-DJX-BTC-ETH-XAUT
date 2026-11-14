"""Product key, activation and the licence file.

The licence lives next to the executable, so it is per-installation and never
travels with the product. The check is deliberately self-contained: it reads
the file, verifies the signature, and says so in plain words.
"""

import hashlib
import hmac
import json
import os
import platform
import time
from datetime import date, timedelta

import delta_state
from delta_state import MONTHS, MONTHS_REV

LICENSE_PATH = os.path.join(delta_state.BASE_DIR, "license.json")


_SK = [0x5A, 0xA7, 0x13, 0x66, 0xF0, 0x2B, 0x9C, 0x41]


_SE = [30, 255, 50, 81, 129, 123, 234, 115, 121, 245, 126, 95, 188, 83,
       168, 1, 13, 221, 43, 66, 164, 69, 169]


_LICENSE_SECRET = bytes(b ^ _SK[i % 8] for i, b in enumerate(_SE))


def _key_mask():
    return int(hashlib.sha256(_LICENSE_SECRET + b"|mask").hexdigest()[:16], 16) % 100000000


def _vmask(variant):
    h = hashlib.sha256(f"{variant}|vm|".encode() + _LICENSE_SECRET).hexdigest()
    return int(h[:16], 16) % 100000000


def _key_check(token, variant):
    h = hashlib.sha256(f"{token}|{variant}|".encode() + _LICENSE_SECRET).hexdigest()
    return int(h[:4], 16) % 100


def key_to_date(key):
    """Decode hashed product key -> valid-until date, or None if invalid."""
    try:
        key = int(str(key).strip())
    except (TypeError, ValueError):
        return None
    if key <= 0:
        return None
    token_p, rest = divmod(key, 10000)
    variant, chk = divmod(rest, 100)
    token = token_p ^ _vmask(variant)
    if token <= 0 or not hmac.compare_digest(
        str(chk).zfill(2), str(_key_check(token, variant)).zfill(2)
    ):
        return None  # SHA256 check digits do not match (variant counted in)
    raw = token ^ _key_mask()
    s = f"{raw:08d}"
    if len(s) != 8:
        return None
    try:
        dd, mm, yyyy = int(s[0:2]), int(s[2:4]), int(s[4:8])
        if yyyy < 2020 or yyyy > 2100:
            return None
        return date(yyyy, mm, dd)  # calendar-safe: 30 Feb -> ValueError
    except ValueError:
        return None


def key_accepted(valid_until, today=None):
    """Key valid if its date is a future cycle end, max 45 days out."""
    today = today or date.today()
    return valid_until is not None and today < valid_until <= today + timedelta(days=45)


def _license_sig(data):
    canon = "|".join(str(data.get(k, "")) for k in ("valid_through", "issued", "key_month"))
    return hmac.new(_LICENSE_SECRET, canon.encode(), hashlib.sha256).hexdigest()


def _verify_sig(data):
    sig = data.get("sig")
    if not isinstance(sig, str):
        return False
    return hmac.compare_digest(sig, _license_sig(data))


def license_is_valid(today=None):
    today = today or date.today()
    try:
        with open(LICENSE_PATH) as f:
            data = json.load(f)
        if not _verify_sig(data):
            return False, None
        vt = date.fromisoformat(str(data.get("valid_through", "")))
        try:
            ls = date.fromisoformat(str(data.get("last_seen", "")))
        except ValueError:
            ls = None
        if ls is not None and ls > today:
            return False, None  # system clock rolled back -> reject
        if vt >= today:
            if ls != today:
                data["last_seen"] = today.isoformat()
                with open(LICENSE_PATH, "w") as f:
                    json.dump(data, f, indent=2)
            return True, vt
        return False, vt
    except (OSError, ValueError, KeyError, TypeError):
        # no license yet -> NOT pre-activated, ask for key on first run
        return False, None


def save_license(valid_until, today=None):
    today = today or date.today()
    payload = {
        "valid_through": valid_until.isoformat(),
        "issued": today.isoformat(),
        "key_month": f"{valid_until.year:04d}-{valid_until.month:02d}",
        "last_seen": today.isoformat(),
    }
    payload["sig"] = _license_sig(payload)
    with open(LICENSE_PATH, "w") as f:
        json.dump(payload, f, indent=2)


def ask_product_key():
    """Wide product-key popup. Returns stripped str, '' or None if cancelled."""
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        result = {"key": None}

        win = tk.Toplevel(root)
        win.title("Product Key")
        win.geometry("480x200")
        win.resizable(False, False)
        win.attributes("-topmost", True)

        tk.Label(
            win, text="License expired.",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", padx=24, pady=(22, 4))
        tk.Label(
            win, text="Enter product key (from owner):",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=24, pady=(0, 8))

        entry = tk.Entry(
            win, font=("Consolas", 13), width=32, justify="center",
            relief="solid", bd=1,
        )
        entry.pack(fill="x", padx=24)
        entry.focus_set()

        def on_ok():
            result["key"] = entry.get()
            win.destroy()

        def on_cancel():
            result["key"] = None
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=18)
        tk.Button(btns, text="OK", width=14, font=("Segoe UI", 10), command=on_ok).pack(side="left", padx=10)
        tk.Button(btns, text="Cancel", width=14, font=("Segoe UI", 10), command=on_cancel).pack(side="left", padx=10)

        win.bind("<Return>", lambda e: on_ok())
        win.bind("<Escape>", lambda e: on_cancel())
        win.protocol("WM_DELETE_WINDOW", on_cancel)

        # center on screen
        win.update_idletasks()
        x = (win.winfo_screenwidth() - 480) // 2
        y = (win.winfo_screenheight() - 200) // 2
        win.geometry(f"+{x}+{y}")

        root.wait_window(win)
        try:
            root.destroy()
        except Exception:
            pass
        k = result["key"]
        return k.strip() if isinstance(k, str) else k
    except Exception as e:
        print(f"Key dialog failed: {e}", flush=True)
        return None


def ensure_license():
    """Return True if app may run; False -> user must fix / we exit."""
    ok, vt = license_is_valid()
    if ok:
        return True

    # expired — require product key; invalid key -> ask again
    while True:
        key = ask_product_key()
        if key is None or str(key).strip() == "":
            show_error(
                "Product key required.\n\n"
                "Contact the owner for the key.\n"
                "App will close."
            )
            return False
        me = key_to_date(key)
        if key_accepted(me):
            save_license(me, date.today())
            show_info(
                f"License activated through {me.strftime('%d/%m/%Y')}.\n"
                "Continuing..."
            )
            return True
        show_error("Invalid product key. Please try again.")


def show_error(msg):
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Option Chain Tool", 0x10)
    else:
        print(msg, flush=True)


def show_info(msg):
    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "Option Chain Tool", 0x40)
    else:
        print(msg, flush=True)
