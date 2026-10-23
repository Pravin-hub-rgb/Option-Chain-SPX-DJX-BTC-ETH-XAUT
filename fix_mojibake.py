"""Find and repair double-encoded text in the source files.

Some source files were once written by a tool that decoded UTF-8 as cp1252
and saved the result as UTF-8. Every non-ASCII character in them is then
wrong: an em dash becomes "â€”", a plus-minus becomes "Â±", and a rupee
sign becomes "â‚¹".

The rupee one matters most: the tool strips a rupee sign out of incoming
values, and it was stripping the mojibake instead, so a value that really
was in rupees kept its sign and the customer saw it in the sheet.
"""

import io
import os
import sys

# The mojibake forms, longest first so nothing is half-replaced.
MOJIBAKE = {
    "â‚¹": "₹",  # rupee sign
    "â€œ": "“",  # left double quote
    "â€\x9d": "”",  # right double quote
    "â€™": "’",  # right single quote
    "â€˜": "‘",  # left single quote
    "â€”": "—",  # em dash
    "â€“": "–",  # en dash
    "Â±": "±",  # plus-minus
    "Â°": "°",  # degree
    "Â·": "·",  # middle dot
    "Â ": " ",  # non-breaking space
    "â€¦": "…",  # ellipsis
}


def scan(path):
    text = io.open(path, encoding="utf-8").read()
    hits = {}
    for bad, good in MOJIBAKE.items():
        count = text.count(bad)
        if count:
            hits[bad] = (good, count)
    return text, hits


def main(roots):
    total = 0
    for root in roots:
        for base, _dirs, files in os.walk(root):
            if "tools" in base.split(os.sep) or "__pycache__" in base:
                continue
            for name in sorted(files):
                if not name.endswith(".py"):
                    continue
                path = os.path.join(base, name)
                text, hits = scan(path)
                if not hits:
                    continue
                for bad, (good, count) in hits.items():
                    print(f"  {path}: {count} x {bad!r} -> {good!r}".encode(
                        "ascii", "backslashreplace").decode("ascii"))
                    total += count
                fixed = text
                for bad, (good, _count) in hits.items():
                    fixed = fixed.replace(bad, good)
                io.open(path, "w", encoding="utf-8", newline="\n").write(fixed)
    print(f"repaired {total} mojibake sequence(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
