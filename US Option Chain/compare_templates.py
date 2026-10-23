"""The parts of an xlsx that can make Excel hang on open:
data validation, defined names, conditional formatting, and the
workbook-level parts.
"""

import os
import re
import zipfile

US = os.path.join("source", "spx_option_chain.xlsx")
DELTA = os.path.join("..", "Delta Option Chain", "source", "btc_chain.xlsx")

WANTED = re.compile(
    r"<(dataValidation|definedName|conditionalFormatting|autoFilter|"
    r"mergeCell|hyperlink|legacyDrawing|extLst|tablePart)\b.*?(/>|</\1>)",
    re.S,
)

for label, path in (("US", US), ("DELTA", DELTA)):
    print(f"=== {label}")
    with zipfile.ZipFile(path) as zf:
        for member in zf.namelist():
            if not member.endswith(".xml"):
                continue
            text = zf.read(member).decode("utf-8", "replace")
            hits = WANTED.findall(text)
            found = WANTED.finditer(text)
            matches = [m.group(0) for m in found]
            if matches:
                for m in matches[:6]:
                    print(f"   [{member}] {m[:400]}")
    print()
