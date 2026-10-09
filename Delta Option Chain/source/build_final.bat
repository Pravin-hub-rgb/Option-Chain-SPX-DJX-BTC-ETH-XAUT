@echo off
REM Delta Option Chain onedir build. Distribute the complete dist\OptionChain folder.
REM --noupx is required: UPX re-compresses every archive on top of PyInstaller's
REM own zlib pass and was killing the build at the "Building PKG (CArchive)" step.
REM matplotlib/scipy/PyQt5/PIL are excluded because the warn log already reports
REM them as missing-or-optional; they are not imported by main.py.
REM Do NOT exclude pandas, numpy, tkinter, xlwings, openpyxl or lxml - all required.
cd /d "%~dp0"
python -m PyInstaller --noconfirm --clean --onedir --noconsole --noupx --name OptionChain --collect-all xlwings --collect-all openpyxl --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32timezone --exclude-module matplotlib --exclude-module scipy --exclude-module PyQt5 --exclude-module PIL main.py > "%temp%\b_delta_final.log" 2>&1
if errorlevel 1 (
  echo EXITCODE=1 >> "%temp%\b_delta_final.log"
  exit /b 1
)
echo EXITCODE=0 >> "%temp%\b_delta_final.log"
copy /Y "*.xlsx" "dist\OptionChain\" >nul
if errorlevel 1 exit /b 1
copy /Y "config.json" "dist\OptionChain\" >nul
if errorlevel 1 exit /b 1
REM license.json is deliberately NOT copied. It is the owner's activation file,
REM not a client asset: shipping it would let anyone run the tool unlocked.
REM On first run the app creates it itself once a valid product key is entered.
if exist "dist\OptionChain\license.json" del /Q "dist\OptionChain\license.json"
copy /Y "..\client package\README.txt" "dist\OptionChain\" >nul
if errorlevel 1 exit /b 1
copy /Y "..\client package\diagnose_excel_addins.ps1" "dist\OptionChain\" >nul
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File "package_release.ps1"
if errorlevel 1 exit /b 1
exit /b 0
