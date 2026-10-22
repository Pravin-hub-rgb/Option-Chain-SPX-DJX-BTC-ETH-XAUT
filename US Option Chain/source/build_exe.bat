@echo off
cd /d "%~dp0"
set "SETUPTOOLS_USE_DISTUTILS=stdlib"
set "PRODUCT_EXE=us_option_chain"
echo Building %PRODUCT_EXE%.exe (onedir, no console)...

rem Build the complete onedir client folder. Do not create an archive here.

rem databento_dbn ships a native extension (_lib*.pyd). PyInstaller's
rem --collect-all does not always register it as an importable submodule of
rem the frozen package, which produced "No module named 'databento_dbn._lib'"
rem on the client. Add the binary explicitly.
set "DBN_LIB="
for /f "delims=" %%p in ('python -c "import databento_dbn._lib as l; print(l.__file__)" 2^>nul') do set "DBN_LIB=%%p"
if not defined DBN_LIB (
  echo Could not locate databento_dbn._lib - build cannot continue.
  exit /b 1
)
echo Using native module: %DBN_LIB%

python -m PyInstaller --noconfirm --clean --onedir --noconsole --noupx ^
  --name %PRODUCT_EXE% ^
  --collect-all xlwings ^
  --collect-all openpyxl ^
  --collect-all databento ^
  --collect-all databento_dbn ^
  --copy-metadata databento ^
  --add-binary "%DBN_LIB%;databento_dbn" ^
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

rem Trim developer-only payload that the runtime never reads.
rem  - pyarrow\include  : C++ headers for building against Arrow (~4.5 MB)
rem  - *.lib            : MSVC import libraries; only the .dll is loaded at run time
rem  - *.pyi            : type stubs for IDEs; the interpreter never reads them
if exist "dist\%PRODUCT_EXE%\_internal\pyarrow\include" rd /s /q "dist\%PRODUCT_EXE%\_internal\pyarrow\include"
if exist "dist\%PRODUCT_EXE%\_internal\pyarrow\lib" rd /s /q "dist\%PRODUCT_EXE%\_internal\pyarrow\lib"
del /s /q "dist\%PRODUCT_EXE%\*.lib" >nul 2>&1
del /s /q "dist\%PRODUCT_EXE%\*.pyi" >nul 2>&1
del /s /q "dist\%PRODUCT_EXE%\*.pxd" >nul 2>&1
echo Trimmed dev-only payload.

set "PKG=..\client package"
if not exist "%PKG%" mkdir "%PKG%"

> "dist\%PRODUCT_EXE%\optionchain.product" echo us

rem Stamp the build so the tool can report which build it is, replacing the
rem separate .sha256 file that used to be shipped beside the ZIP. Support now
rem reads build_id from startup_diag.log instead.
rem Only files that actually ship count towards the hash. Tests and scratch
rem scripts are not part of the delivered product, so editing a test must not
remake the build id.
set "BUILD_SHA=source-local"
for /f "delims=" %%i in ('python -c "import hashlib,pathlib; r=pathlib.Path('.'); EX=('license.json',); fs=sorted(p for p in r.iterdir() if p.is_file() and p.suffix.lower() in {'.py','.json','.spec'} and not p.name.startswith(('_','test_')) and p.name not in EX); h=hashlib.sha256(); [h.update(p.name.encode()+b'\0'+p.read_bytes()) for p in fs]; print('src-'+h.hexdigest()[:12])"') do set "BUILD_SHA=%%i"
> "dist\%PRODUCT_EXE%\build_info.txt" echo product=us
>> "dist\%PRODUCT_EXE%\build_info.txt" echo build_sha=%BUILD_SHA%
>> "dist\%PRODUCT_EXE%\build_info.txt" echo build_date=%DATE% %TIME%

rem Copy the complete onedir runtime. Do not create a ZIP here.
rem Use robocopy, not xcopy: xcopy is not resolving on this machine and was
rem silently leaving a stale _internal folder in the package, which made the
rem client fail with "No module named 'databento_dbn._lib'".
rem /MIR makes the package an exact mirror, so files removed from the build
rem (trimmed headers, stale modules) do not linger in the delivered folder.
rem It must run BEFORE the templates are copied, because /MIR deletes anything
rem in the target that is not in the source.
robocopy "dist\%PRODUCT_EXE%" "%PKG%" /MIR /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 (
  echo COPY ONEDIR PACKAGE FAILED
  exit /b 1
)

rem Excel templates (also auto-created at runtime, but ship pre-built)
python -c "import main; main.ensure_excel_file('spx'); main.ensure_excel_file('djx'); main.ensure_excel_file('stock')"
if exist "spx_option_chain.xlsx" copy /y "spx_option_chain.xlsx" "%PKG%\" >nul
if exist "djx_option_chain.xlsx" copy /y "djx_option_chain.xlsx" "%PKG%\" >nul
if exist "us_stock_option_chain.xlsx" copy /y "us_stock_option_chain.xlsx" "%PKG%\" >nul
copy /y "config.json" "%PKG%\" >nul
if exist "HOW-TO-USE.txt" copy /y "HOW-TO-USE.txt" "%PKG%\" >nul

rem never ship an activated licence
if exist "%PKG%\license.json" del /f /q "%PKG%\license.json"
del /f /q "%PKG%\~$*.xlsx" >nul 2>&1

rem The product marker tells the instance guard which product a folder holds.
rem Without it the guard has to guess from the workbook names, and a customer
rem who extracts both packages into one folder ends up with one product
rem overwriting the other. Fail the build rather than ship without it.
if not exist "%PKG%\optionchain.product" (
  echo MISSING PRODUCT MARKER in %PKG%
  exit /b 1
)
findstr /b /c:"us" "%PKG%\optionchain.product" >nul
if errorlevel 1 (
  echo PRODUCT MARKER IS NOT 'us'
  exit /b 1
)
if not exist "%PKG%\%PRODUCT_EXE%.exe" (
  echo MISSING %PRODUCT_EXE%.exe in %PKG%
  exit /b 1
)
echo Verified: product marker and %PRODUCT_EXE%.exe are in the package.

echo.
echo Build OK:
echo   dist\%PRODUCT_EXE%\
echo   ..\client package\  (complete onedir folder)
echo   No ZIP was created.
echo.
