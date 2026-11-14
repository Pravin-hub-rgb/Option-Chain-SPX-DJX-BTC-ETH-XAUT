import os, sys, shutil
from datetime import date, timedelta
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, ".")
import main as M

def mint(target: date) -> int:
    """Solve for a token whose decoded date == target, then wrap it as a key."""
    raw = int(f"{target.day:02d}{target.month:02d}{target.year:04d}")
    token = raw ^ M._key_mask()
    if token <= 0:
        return 0
    for variant in range(100):
        chk = M._key_check(token, variant)
        key = (token ^ M._vmask(variant)) * 10000 + variant * 100 + chk
        if M.key_to_date(key) == target:
            return key
    return 0

target = date.today() + timedelta(days=29)
key = mint(target)
print("target:", target, " key:", key)
print("decoded:", M.key_to_date(key), " accepted:", M.key_accepted(M.key_to_date(key) or target))
if key:
    M.save_license(target, date.today())
    shutil.copy("license.json", r"C:\UsAudit\license.json")
    M.LICENSE_PATH = r"C:\UsAudit\license.json"
    print("license_is_valid ->", M.license_is_valid())
