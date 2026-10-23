"""Why does Excel refuse to open the US workbooks?

os.startfile on the Delta workbook opens Excel. On the US workbook it
returns and nothing happens: no EXCEL.EXE process ever appears.

Check the file itself: is it a valid zip, does openpyxl read it, is there a
mark-of-the-web, is it locked.
"""

import glob
import io
import os
import zipfile

folder = os.path.join("source")
for name in sorted(glob.glob(os.path.join(folder, "*.xlsx"))):
    print("==", os.path.basename(name), os.path.getsize(name), "bytes")
    print("   zip ok:", zipfile.is_zipfile(name))
    try:
        with zipfile.ZipFile(name) as zf:
            print("   members:", len(zf.namelist()), "bad:", zf.testzip())
    except Exception as e:
        print("   zip error:", type(e).__name__, e)
    try:
        import openpyxl

        wb = openpyxl.load_workbook(name)
        ws = wb.active
        print("   openpyxl ok:", ws.title, ws.dimensions,
              "A4:", ws.cell(row=4, column=1).value)
    except Exception as e:
        print("   openpyxl error:", type(e).__name__, str(e)[:120])
    for extra in (name + ":Zone.Identifier",):
        if os.path.exists(extra):
            print("   MOTW present:", extra)
