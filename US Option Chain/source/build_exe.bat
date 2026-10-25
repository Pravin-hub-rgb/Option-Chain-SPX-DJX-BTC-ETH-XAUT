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

rem pyarrow CANNOT be excluded wholesale (unlike the Delta build, which does
rem not ship databento): databento/__init__.py imports databento.common.dbnstore
rem unconditionally, and dbnstore.py does "import pyarrow" + "import pyarrow.parquet"
rem at module top. The verified pull chain (import databento at runtime) needs only
rem pyarrow.lib/_compute/ipc/types/util, pyarrow.parquet, pyarrow.fs and the cloud
rem filesystem backends pyarrow.fs/__init__ imports eagerly (_s3fs/_gcsfs/_azurefs/_hdfs).
rem The modules below never appear in that chain (verified against sys.modules on
rem 2026-10-25); excluding them trims ~23 MB of dead arrow DLLs/pyds
rem (flight, substrait, dataset, acero, orc, csv, json, feather, encryption).

python -m PyInstaller --noconfirm --clean --onedir --noconsole --noupx ^
  --name %PRODUCT_EXE% ^
  --collect-all xlwings ^
  --collect-all openpyxl ^
  --collect-all databento ^
  --collect-all databento_dbn ^
  --copy-metadata databento ^
  --add-binary "%DBN_LIB%;databento_dbn" ^
  --exclude-module pyarrow.flight ^
  --exclude-module pyarrow.substrait ^
  --exclude-module pyarrow.dataset ^
  --exclude-module pyarrow.orc ^
  --exclude-module pyarrow.csv ^
  --exclude-module pyarrow.json ^
  --exclude-module pyarrow.feather ^
  --exclude-module pyarrow.acero ^
  --exclude-module pyarrow.parquet.encryption ^
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

rem The --exclude-module flags above keep the excluded pyarrow submodules out
rem of the PYZ, but PyInstaller's pyarrow contrib hook still ships their
rem compiled .pyd/.dll plus cython sources as data. Only the modules that the
rem verified databento pull chain imports must survive
rem (lib/_compute/ipc/types/util + parquet/_parquet + fs and the cloud
rem filesystem backends pyarrow.fs imports eagerly: _s3fs/_gcsfs/_azurefs/_hdfs).
rem The dot in each mask keeps _parquet._/_dataset._ from matching
rem _parquet_encryption/_dataset_orc etc.; those variants are listed explicitly.
for %%m in (_acero _csv _dataset _dataset_orc _dataset_parquet _dataset_parquet_encryption _feather _flight _json _orc _parquet_encryption _substrait) do del /q "dist\%PRODUCT_EXE%\_internal\pyarrow\%%m.*.pyd" >nul 2>&1
for %%d in (arrow_acero arrow_dataset arrow_flight arrow_python_flight arrow_python_parquet_encryption arrow_substrait) do del /q "dist\%PRODUCT_EXE%\_internal\pyarrow\%%d.dll" >nul 2>&1
rem cython sources, C++ source tree, and pyarrow's own test suite: never runtime
del /q "dist\%PRODUCT_EXE%\_internal\pyarrow\*.pyx" >nul 2>&1
del /q "dist\%PRODUCT_EXE%\_internal\pyarrow\*.pxi" >nul 2>&1
for %%d in (includes src tests) do if exist "dist\%PRODUCT_EXE%\_internal\pyarrow\%%d" rd /s /q "dist\%PRODUCT_EXE%\_internal\pyarrow\%%d"
echo Trimmed dev-only payload.

set "PKG=..\client package"

rem BASE_DIR resolves to the client package whenever optionchain.product is
rem present there, which made the template step below write into the package
rem and the source-dir label check fail with FileNotFoundError on every build
rem after the first. Clear the marker so the generation resolves to this
rem source folder; the marker is stamped into dist below and robocopy restores
rem it into the package.
if exist "%PKG%\optionchain.product" del /q "%PKG%\optionchain.product"

rem Excel templates (generated before robocopy so BASE_DIR resolves to source/,
rem not client package/). also auto-created at runtime, but ship pre-built.
rem
rem Delete them first. ensure_excel_file only rewrites the labels it owns, and
rem it cannot save a file Excel has open, so a template that went bad in a
rem previous session survived every build and shipped broken. A fresh file
rem has no such history.
if exist "spx_option_chain.xlsx" del /q "spx_option_chain.xlsx"
if exist "djx_option_chain.xlsx" del /q "djx_option_chain.xlsx"
if exist "us_stock_option_chain.xlsx" del /q "us_stock_option_chain.xlsx"
if exist "spx_option_chain_new.xlsx" del /q "spx_option_chain_new.xlsx"
if exist "djx_option_chain_new.xlsx" del /q "djx_option_chain_new.xlsx"
if exist "us_stock_option_chain_new.xlsx" del /q "us_stock_option_chain_new.xlsx"
python -c "import us_workbook; us_workbook.ensure_excel_file('spx'); us_workbook.ensure_excel_file('djx'); us_workbook.ensure_excel_file('stock')"
rem And check what came out. A template with a mangled label in it is worse
rem than no template at all, because the customer sees it in the header.
python -c "import openpyxl,sys; want='\u00b1 strikes \u2014 edit A2'; bad=[]; [bad.append(f) for f in ('spx_option_chain.xlsx','djx_option_chain.xlsx','us_stock_option_chain.xlsx') if openpyxl.load_workbook(f).active.cell(row=2,column=2).value != want]; sys.exit('Template labels are wrong: '+','.join(bad)) if bad else None"
if errorlevel 1 (
  echo TEMPLATE LABEL CHECK FAILED - not shipping a package with a mangled header
  exit /b 1
)

mkdir "%PKG%" >nul 2>&1

> "dist\%PRODUCT_EXE%\optionchain.product" echo us

rem Stamp the build so the tool can report which build it is, replacing the
rem separate .sha256 file that used to be shipped beside the ZIP. Support now
rem reads build_id from startup_diag.log instead.
rem Only files that actually ship count towards the hash. Tests and scratch
rem scripts are not part of the delivered product, so editing a test must not
rem remake the build id.
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
rem /R:1 /W:1 keeps a file locked by an open workbook from stalling the build
rem for minutes on Excel's default 30x30-second retry loop.
robocopy "dist\%PRODUCT_EXE%" "%PKG%" /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 (
  echo COPY ONEDIR PACKAGE FAILED
  exit /b 1
)

if exist "spx_option_chain.xlsx" copy /y "spx_option_chain.xlsx" "%PKG%\" >nul
if exist "djx_option_chain.xlsx" copy /y "djx_option_chain.xlsx" "%PKG%\" >nul
if exist "us_stock_option_chain.xlsx" copy /y "us_stock_option_chain.xlsx" "%PKG%\" >nul
copy /y "config.json" "%PKG%\" >nul
if exist "HOW-TO-USE.txt" copy /y "HOW-TO-USE.txt" "%PKG%\" >nul
if exist "CHECK-INSTALL.bat" copy /y "CHECK-INSTALL.bat" "%PKG%\" >nul

rem never ship an activated licence
if exist "%PKG%\license.json" del /f /q "%PKG%\license.json"
del /f /q "%PKG%\~$*.xlsx" >nul 2>&1

rem runtime artefacts are not part of the product
if exist "%PKG%\instance.lock" del /f /q "%PKG%\instance.lock"
if exist "%PKG%\instance_guard.log" del /f /q "%PKG%\instance_guard.log"
if exist "%PKG%\doctor-report.txt" del /f /q "%PKG%\doctor-report.txt"

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
