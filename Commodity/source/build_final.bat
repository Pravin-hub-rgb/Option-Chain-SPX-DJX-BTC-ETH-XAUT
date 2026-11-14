@echo off
REM MCX Commodity Option Chain onedir build. Distribute the complete dist\mcx_option_chain folder.
REM The executable is named for the product so the customer can tell the tools
REM apart in Task Manager, and so the instance guard never has to guess which
REM product a running process belongs to.
REM --noupx is required: UPX re-compresses every archive on top of PyInstaller's
REM own zlib pass and was killing the build at the "Building PKG (CArchive)" step.
REM matplotlib/scipy/PyQt5/PIL are excluded because they are not imported by
REM main.py. Do NOT exclude pandas, numpy, tkinter, xlwings or openpyxl.
cd /d "%~dp0"
set "SETUPTOOLS_USE_DISTUTILS=stdlib"
set "PRODUCT_EXE=mcx_option_chain"
python -m PyInstaller --noconfirm --clean --onedir --noconsole --noupx --name %PRODUCT_EXE% --collect-all xlwings --collect-all openpyxl --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32timezone --exclude-module matplotlib --exclude-module scipy --exclude-module PyQt5 --exclude-module PIL --exclude-module pyarrow --exclude-module pyarrow.lib main.py > "%temp%\b_mcx_final.log" 2>&1
if errorlevel 1 (
  echo EXITCODE=1 >> "%temp%\b_mcx_final.log"
  exit /b 1
)
echo EXITCODE=0 >> "%temp%\b_mcx_final.log"
REM Trim the developer-only payload the runtime never reads.
if exist "dist\%PRODUCT_EXE%\_internal\pyarrow\include" rd /s /q "dist\%PRODUCT_EXE%\_internal\pyarrow\include"
del /s /q "dist\%PRODUCT_EXE%\*.lib" >nul 2>&1
del /s /q "dist\%PRODUCT_EXE%\*.pyi" >nul 2>&1
del /s /q "dist\%PRODUCT_EXE%\*.pxd" >nul 2>&1
if exist "dist\%PRODUCT_EXE%\_internal\pyarrow" rd /s /q "dist\%PRODUCT_EXE%\_internal\pyarrow"
> "dist\%PRODUCT_EXE%\optionchain.product" echo mcx
REM Stamp the build so the tool can report which build it is.
set "BUILD_SHA=source-local"
for /f "delims=" %%i in ('python -c "import hashlib,pathlib; r=pathlib.Path('.'); EX=('license.json',); fs=sorted(p for p in r.iterdir() if p.is_file() and p.suffix.lower() in {'.py','.json','.spec'} and not p.name.startswith(('_','test_')) and p.name not in EX); h=hashlib.sha256(); [h.update(p.name.encode()+b'\0'+p.read_bytes()) for p in fs]; print('src-'+h.hexdigest()[:12])"') do set "BUILD_SHA=%%i"
> "dist\%PRODUCT_EXE%\build_info.txt" echo product=mcx
>> "dist\%PRODUCT_EXE%\build_info.txt" echo build_sha=%BUILD_SHA%
>> "dist\%PRODUCT_EXE%\build_info.txt" echo build_date=%DATE% %TIME%
copy /Y "*.xlsx" "dist\%PRODUCT_EXE%\" >nul
if errorlevel 1 exit /b 1
copy /Y "config.json" "dist\%PRODUCT_EXE%\" >nul
if errorlevel 1 exit /b 1
REM license.json is deliberately NOT copied. It is the owner's activation file.
if exist "dist\%PRODUCT_EXE%\license.json" del /Q "dist\%PRODUCT_EXE%\license.json"
if exist "HOW-TO-USE.txt" copy /Y "HOW-TO-USE.txt" "dist\%PRODUCT_EXE%\" >nul
if exist "CHECK-INSTALL.bat" copy /Y "CHECK-INSTALL.bat" "dist\%PRODUCT_EXE%\" >nul
REM The old wrongly-named executable must never ship.
if exist "dist\%PRODUCT_EXE%\OptionChain.exe" del /Q "dist\%PRODUCT_EXE%\OptionChain.exe"
if exist "dist\%PRODUCT_EXE%\OptionChain.pyd" del /Q "dist\%PRODUCT_EXE%\OptionChain.pyd"
REM Refresh the client package folder from this build.
xcopy /E /I /Y /Q "dist\%PRODUCT_EXE%" "..\client package" >nul
if exist "..\client package\OptionChain.exe" del /Q "..\client package\OptionChain.exe"
if exist "..\client package\OptionChain.pyd" del /Q "..\client package\OptionChain.pyd"
if exist "..\client package\license.json" del /Q "..\client package\license.json"
if exist "..\client package\instance.lock" del /Q "..\client package\instance.lock"
if exist "..\client package\instance_guard.log" del /Q "..\client package\instance_guard.log"
if exist "..\client package\doctor-report.txt" del /Q "..\client package\doctor-report.txt"
if exist "dist\OptionChain" rd /S /Q "dist\OptionChain"
exit /b 0
