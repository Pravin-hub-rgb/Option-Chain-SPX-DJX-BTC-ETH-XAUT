@echo off
cd /d "%~dp0"
echo Building OptionChain.exe (onedir, no console)...

rem Build the complete onedir client folder. Do not create an archive here.

python -m PyInstaller --noconfirm --clean --onedir --noconsole --noupx ^
  --name OptionChain ^
  --collect-all xlwings ^
  --collect-all openpyxl ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32timezone ^
  main.py
if errorlevel 1 (
  echo BUILD FAILED
  exit /b 1
)

set "PKG=..\client package"
if not exist "%PKG%" mkdir "%PKG%"

rem Excel templates (also auto-created at runtime, but ship pre-built)
python -c "import main; main.ensure_excel_file('spx'); main.ensure_excel_file('djx'); main.ensure_excel_file('stock')"
if exist "spx_option_chain.xlsx" copy /y "spx_option_chain.xlsx" "%PKG%\" >nul
if exist "djx_option_chain.xlsx" copy /y "djx_option_chain.xlsx" "%PKG%\" >nul
if exist "us_stock_option_chain.xlsx" copy /y "us_stock_option_chain.xlsx" "%PKG%\" >nul
copy /y "config.json" "%PKG%\" >nul

rem Copy the complete onedir runtime and templates. Do not create a ZIP here.
xcopy /e /i /y "dist\OptionChain\*" "%PKG%\" >nul
if errorlevel 2 (
  echo COPY ONEDIR PACKAGE FAILED
  exit /b 1
)

rem never ship an activated licence
if exist "%PKG%\license.json" del /f /q "%PKG%\license.json"
del /f /q "%PKG%\~$*.xlsx" >nul 2>&1

echo.
echo Build OK:
echo   dist\OptionChain\
echo   ..\client package\  (complete onedir folder)
echo   No ZIP was created.
echo.
