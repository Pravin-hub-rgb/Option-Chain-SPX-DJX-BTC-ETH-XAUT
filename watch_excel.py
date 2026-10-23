"""Watch Excel's life while the US tool runs.

Prints every change: a new EXCEL.EXE, a window title, and the moment it
goes away. Run this alongside the tool.
"""

import subprocess
import time


def excel_processes():
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
    ).stdout
    rows = []
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[0].upper().startswith("EXCEL"):
            rows.append((parts[1], parts[-1] if len(parts) > 2 else ""))
    return rows


def main():
    last = None
    for _ in range(200):
        rows = excel_processes()
        state = tuple(sorted(rows))
        if state != last:
            stamp = time.strftime("%H:%M:%S")
            if state:
                print(f"{stamp} EXCEL running: {state}", flush=True)
            else:
                print(f"{stamp} EXCEL GONE", flush=True)
            last = state
        time.sleep(3)


if __name__ == "__main__":
    main()
