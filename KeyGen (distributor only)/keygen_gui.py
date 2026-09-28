"""Distributor-facing Key Generator (builds to KeyGen.exe).

Ye program distributor ke paas jaata hai - END USERS KO NAHI.
Program kholo -> screen pe EK key + Copy + close. Bas. Date ki sochne ki zaroorat nahi:
- jo key dikh rahi hai wo aaj ke liye sabse lambi valid key hai (auto)
- har launch pe number apne aap badalta hai (pichhle se alag, round-robin)
- cycle change -> program khud nayi date ki key dikhata hai
"""
import hashlib
import os
import random
import sys
import tkinter as tk
from datetime import date, timedelta

_SK = [0x5A, 0xA7, 0x13, 0x66, 0xF0, 0x2B, 0x9C, 0x41]
_SE = [30, 255, 50, 81, 129, 123, 234, 115, 121, 245, 126, 95, 188, 83,
       168, 1, 13, 221, 43, 66, 164, 69, 169]
_SECRET = bytes(b ^ _SK[i % 8] for i, b in enumerate(_SE))


def _mask8():
    return int(hashlib.sha256(_SECRET + b"|mask").hexdigest()[:16], 16) % 100000000


def _vmask(variant):
    h = hashlib.sha256(f"{variant}|vm|".encode() + _SECRET).hexdigest()
    return int(h[:16], 16) % 100000000


def _check(token, variant):
    h = hashlib.sha256(f"{token}|{variant}|".encode() + _SECRET).hexdigest()
    return int(h[:4], 16) % 100


def key_for(d, variant=0):
    variant %= 100
    raw = int(f"{d.day:02d}{d.month:02d}{d.year:04d}")
    token = raw ^ _mask8()
    scrambled = token ^ _vmask(variant)
    return scrambled * 10000 + variant * 100 + _check(token, variant)


def last_friday(y, m):
    nxt = date(y + 1, 1, 1) if m == 12 else date(y, m + 1, 1)
    d = nxt - timedelta(days=1)
    while d.weekday() != 4:
        d -= timedelta(days=1)
    return d


def next_cycle(today):
    d = last_friday(today.year, today.month)
    if d > today:
        return d
    if today.month == 12:
        y, m = today.year + 1, 1
    else:
        y, m = today.year, today.month + 1
    return last_friday(y, m)


def best_cycle(today=None):
    """Sabse lambi chalne wali cycle: today < d <= today+45 (app rule)."""
    today = today or date.today()
    best = None
    d = next_cycle(today)
    for _ in range(4):
        if today < d <= today + timedelta(days=45):
            best = d
        d = next_cycle(d + timedelta(days=1))
    return best or next_cycle(today)


def _counter_paths():
    paths = []
    try:
        exe_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        paths.append(os.path.join(exe_dir, "kg.dat"))
    except Exception:
        pass
    appdata = os.environ.get("LOCALAPPDATA")
    if appdata:
        paths.append(os.path.join(appdata, "dxc_kg.dat"))
    return paths


def next_variant():
    """Round-robin 0-99 (counter file), file trouble -> random."""
    for p in _counter_paths():
        try:
            n = 0
            if os.path.exists(p):
                with open(p) as f:
                    n = int(f.read().strip() or 0)
            v = n % 100
            with open(p, "w") as f:
                f.write(str(n + 1))
            return v
        except Exception:
            continue
    return random.randrange(100)


def _flat_btn(parent, text, command, accent=False):
    """Modern flat button — hover highlight, no old-school 3D border."""
    base = "#2f7cf6" if accent else "#232b36"
    hover = "#5a97ff" if accent else "#2e3947"
    pressed = "#2563d0" if accent else "#1b222b"
    btn = tk.Button(parent, text=text, font=("Segoe UI", 10, "bold"),
                    width=16, padx=18, pady=8,
                    bg=base, fg="#ffffff",
                    activebackground=hover, activeforeground="#ffffff",
                    relief="flat", borderwidth=0, highlightthickness=0,
                    overrelief="flat", cursor="hand2", command=command)
    btn.bind("<Enter>", lambda e: btn.config(bg=hover))
    btn.bind("<Leave>", lambda e: btn.config(bg=base))
    btn.bind("<ButtonPress-1>", lambda e: btn.config(bg=pressed))
    btn.bind("<ButtonRelease-1>", lambda e: btn.config(bg=hover))
    return btn


class App:
    def __init__(self, root):
        self.root = root
        root.title("Key Generator")
        root.resizable(False, False)
        root.geometry("560x236")
        root.configure(bg="#14181f")
        self.cycle = best_cycle()
        self.key_var = tk.StringVar()

        tk.Label(root, text="PRODUCT KEY", font=("Segoe UI", 11, "bold"),
                 fg="#8aa0bf", bg="#14181f").pack(pady=(22, 4))
        tk.Label(root, textvariable=self.key_var, font=("Consolas", 27, "bold"),
                 fg="#ffffff", bg="#14181f", padx=24).pack(pady=(2, 6))
        self.info = tk.Label(root, text="", font=("Segoe UI", 10),
                             fg="#6fcf97", bg="#14181f")
        self.info.pack(pady=(0, 5))
        self.status = tk.Label(root, text="", font=("Segoe UI", 9),
                               fg="#e0b15c", bg="#14181f")
        self.status.pack()

        btns = tk.Frame(root, bg="#14181f")
        btns.pack(pady=(12, 22))
        self.copy_btn = _flat_btn(btns, "Copy", self.do_copy, accent=True)
        self.copy_btn.pack()

        self.new_key()
        root.bind("<Escape>", lambda e: root.destroy())

    def new_key(self):
        # zfill(13): scrambled token can be short (XOR), so a raw key may show as
        # only 7-9 digits and look "truncated" to the user. Decoder does int()
        # so leading zeros are harmless. Max width over 10y = 13 (2**27*10000).
        self.key_var.set(str(key_for(self.cycle, next_variant())).zfill(13))
        self.info.config(
            text=f"Valid till: {self.cycle.strftime('%A, %d %B %Y')} (last Friday)"
        )
        self.status.config(text="")

    def do_copy(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.key_var.get())
        self.root.update()
        self.status.config(text="Copied - send this number to everyone")


def main():
    root = tk.Tk()
    App(root)
    root.eval('tk::PlaceWindow . center')
    root.mainloop()


if __name__ == "__main__":
    main()
