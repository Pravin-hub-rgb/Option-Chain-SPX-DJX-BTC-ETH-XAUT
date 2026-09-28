"""
OWNER ONLY - generate product keys for clients.
Do NOT put this file in the client zip/EXE package.

Usage:
  python keygen.py             # keys for next two cycles (next last Fridays)
  python keygen.py 30 10 2026  # key for an explicit date (dd mm yyyy)
  python keygen.py 30 10 2026 7 # same date, variant 0-99 (any variant works)

Key format (opaque, hashed, multi-variant):
  token   = int(DDMMYYYY) ^ mask           # date hidden
  vmask   = SHA256(secret, variant)        # per-variant scramble (0-99)
  key     = (token ^ vmask)*10000 + variant*100 + SHA256 check
  -> har variant ka PURA number alag dikhta hai; app sab accept karti hai (date same)
"""
import hashlib
import sys
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


def main():
    today = date.today()
    if len(sys.argv) in (4, 5):
        try:
            d = date(int(sys.argv[3]), int(sys.argv[2]), int(sys.argv[1]))
            v = int(sys.argv[4]) if len(sys.argv) == 5 else 0
            if not 0 <= v <= 99:
                raise ValueError("variant must be 0-99")
        except ValueError as e:
            print(f"Bad input: {e}")
            sys.exit(1)
        print(f"Valid until : {d.strftime('%d/%m/%Y')}")
        print(f"PRODUCT KEY: {str(key_for(d, v)).zfill(13)}")
        return
    d1 = next_cycle(today)
    d2 = next_cycle(d1 + timedelta(days=1))
    print(f"Today: {today.strftime('%d/%m/%Y')}")
    print("-" * 44)
    print(f"[CURRENT cycle] ends {d1.strftime('%d/%m/%Y')}  ->  KEY: {str(key_for(d1)).zfill(13)}")
    print(f"[NEXT cycle    ] ends {d2.strftime('%d/%m/%Y')}  ->  KEY: {str(key_for(d2)).zfill(13)}")
    print(f"(variants: key_for(d, 0..99) sab valid - alag number, same date)")


if __name__ == "__main__":
    main()
