"""Can the long path be opened using the Windows 8.3 short name?

C:\\Users\\Pravin\\Downloads\\Telegram Desktop\\... is 132 characters and
Excel's Workbooks.Open hangs on it. The 8.3 short name is far shorter.
"""

import ctypes
import os
import subprocess
import time

import pythoncom
import xlwings as xw

US = os.path.abspath(os.path.join("source", "spx_option_chain.xlsx"))
print("long path:", len(US), US)

buffer = ctypes.create_unicode_buffer(1024)
rc = ctypes.windll.kernel32.GetShortPathNameW(US, buffer, 1024)
short = buffer.value if rc else ""
print("short path:", len(short) if short else "(none)", short)

subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True)
time.sleep(4)

pythoncom.CoInitialize()
try:
    app = xw.App(visible=True)
    print("Excel started", flush=True)
    for label, path in (("short name", short or US), ("long path", US)):
        t = time.perf_counter()
        try:
            app.api.Workbooks.Open(path)
            print(f"  {label}: opened in {time.perf_counter() - t:.1f}s", flush=True)
        except Exception as e:
            print(f"  {label}: FAILED after {time.perf_counter() - t:.1f}s: {e}", flush=True)
            break
finally:
    pythoncom.CoUninitialize()
