@echo off
cd /d "%~dp0"
echo Building OptionChain.exe (onefile, no console)...

rem NOTE: client package lives in the sibling folder "..\client package"
rem (README.txt / instruction.txt already live there - they are NOT copied in.)

python -m PyInstaller --noconfirm --clean --onefile --noconsole ^
  --name OptionChain ^
  --collect-all xlwings ^
  --collect-all openpyxl ^
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

rem Replace the exe only AFTER templates are in place, so the client package
rem is never left without its .xlsx files.
copy /y "dist\OptionChain.exe" "%PKG%\" >nul
if errorlevel 1 (
  echo COPY EXE FAILED
  exit /b 1
)

rem never ship an activated licence
if exist "%PKG%\license.json" del /f /q "%PKG%\license.json"
del /f /q "%PKG%\~$*.xlsx" >nul 2>&1

if exist "..\us-option-chain.zip" del /f /q "..\us-option-chain.zip"
powershell -NoProfile -Command "Compress-Archive -Path '%PKG%' -DestinationPath '..\us-option-chain.zip'"
if errorlevel 1 (
  echo ZIP FAILED
  exit /b 1
)

echo.
echo Build OK:
echo   dist\OptionChain.exe
echo   ..\client package\
echo   ..\us-option-chain.zip
echo.
