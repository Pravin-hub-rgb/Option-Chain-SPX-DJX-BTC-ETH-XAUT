@echo off
cd /d "%~dp0"
set "SETUPTOOLS_USE_DISTUTILS=stdlib"
set "PRODUCT_EXE=mcx_option_chain"
echo Building %PRODUCT_EXE%.exe (onedir, no console)...
python -m PyInstaller --noconfirm --clean --onedir --noconsole --noupx ^
  --name %PRODUCT_EXE% ^
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
echo Build OK: dist\%PRODUCT_EXE%\
