@echo off
REM Delta Option Chain onedir build. Distribute the complete dist\delta_option_chain folder.
REM The executable is named for the product so the customer can tell the two
REM tools apart in Task Manager, and so the instance guard never has to guess
REM which product a running process belongs to.
REM --noupx is required: UPX re-compresses every archive on top of PyInstaller's
REM own zlib pass and was killing the build at the "Building PKG (CArchive)" step.
REM matplotlib/scipy/PyQt5/PIL are excluded because the warn log already reports
REM them as missing-or-optional; they are not imported by main.py.
REM Do NOT exclude pandas, numpy, tkinter, xlwings, openpyxl or lxml - all required.
cd /d "%~dp0"
set "SETUPTOOLS_USE_DISTUTILS=stdlib"
set "PRODUCT_EXE=delta_option_chain"
python -m PyInstaller --noconfirm --clean --onedir --noconsole --noupx --name %PRODUCT_EXE% --collect-all xlwings --collect-all openpyxl --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32timezone --exclude-module matplotlib --exclude-module scipy --exclude-module PyQt5 --exclude-module PIL --exclude-module pyarrow --exclude-module pyarrow.lib main.py > "%temp%\b_delta_final.log" 2>&1
if errorlevel 1 (
  echo EXITCODE=1 >> "%temp%\b_delta_final.log"
  exit /b 1
)
echo EXITCODE=0 >> "%temp%\b_delta_final.log"
REM pyarrow is 78 MB of the package and nothing here uses it: the chain is
REM built with plain pandas frames, and pandas only reaches for pyarrow when
REM you ask for arrow-backed dtypes. Shipping it doubled the download for
REM nothing, and a download that big is its own kind of "it does not work"
REM when it does not complete.
REM Trim the developer-only payload the runtime never reads.
if exist "dist\%PRODUCT_EXE%\_internal\pyarrow\include" rd /s /q "dist\%PRODUCT_EXE%\_internal\pyarrow\include"
del /s /q "dist\%PRODUCT_EXE%\*.lib" >nul 2>&1
del /s /q "dist\%PRODUCT_EXE%\*.pyi" >nul 2>&1
del /s /q "dist\%PRODUCT_EXE%\*.pxd" >nul 2>&1
if exist "dist\%PRODUCT_EXE%\_internal\pyarrow" rd /s /q "dist\%PRODUCT_EXE%\_internal\pyarrow"
> "dist\%PRODUCT_EXE%\optionchain.product" echo delta
REM Stamp the build so the tool can report which build it is. This replaces the
REM separate .sha256 file that used to be shipped beside the ZIP: the customer
REM could not download that from Drive, and could not act on it even if they
REM could. Now they send startup_diag.log and support reads build_id from it.
REM
REM The id is a hash of the source, not a git commit. A commit hash does not
REM change when the working tree changes, so a build made from uncommitted
REM work would carry the id of some older commit and support could never
REM tell which code the customer is actually running.
REM Only files that actually ship count towards the hash. Tests and scratch
REM scripts are not part of the delivered product, so editing a test must not
REM remake the build id.
set "BUILD_SHA=source-local"
for /f "delims=" %%i in ('python -c "import hashlib,pathlib; r=pathlib.Path('.'); EX=('license.json',); fs=sorted(p for p in r.iterdir() if p.is_file() and p.suffix.lower() in {'.py','.json','.spec'} and not p.name.startswith(('_','test_')) and p.name not in EX); h=hashlib.sha256(); [h.update(p.name.encode()+b'\0'+p.read_bytes()) for p in fs]; print('src-'+h.hexdigest()[:12])"') do set "BUILD_SHA=%%i"
> "dist\%PRODUCT_EXE%\build_info.txt" echo product=delta
>> "dist\%PRODUCT_EXE%\build_info.txt" echo build_sha=%BUILD_SHA%
>> "dist\%PRODUCT_EXE%\build_info.txt" echo build_date=%DATE% %TIME%
copy /Y "*.xlsx" "dist\%PRODUCT_EXE%\" >nul
if errorlevel 1 exit /b 1
copy /Y "config.json" "dist\%PRODUCT_EXE%\" >nul
if errorlevel 1 exit /b 1
REM license.json is deliberately NOT copied. It is the owner's activation file,
REM not a client asset: shipping it would let anyone run the tool unlocked.
REM On first run the app creates it itself once a valid product key is entered.
if exist "dist\%PRODUCT_EXE%\license.json" del /Q "dist\%PRODUCT_EXE%\license.json"
copy /Y "..\client package\README.txt" "dist\%PRODUCT_EXE%\" >nul
if errorlevel 1 exit /b 1
copy /Y "..\client package\diagnose_excel_addins.ps1" "dist\%PRODUCT_EXE%\" >nul
if errorlevel 1 exit /b 1
copy /Y "..\client package\instruction.txt" "dist\%PRODUCT_EXE%\" >nul
if errorlevel 1 exit /b 1
if exist "HOW-TO-USE.txt" copy /Y "HOW-TO-USE.txt" "dist\%PRODUCT_EXE%\" >nul
if exist "CHECK-INSTALL.bat" copy /Y "CHECK-INSTALL.bat" "dist\%PRODUCT_EXE%\" >nul
REM The old wrongly-named executable must never ship. Two executables in one
REM folder, one of them the wrong product, is how a customer ends up running
REM the wrong tool and blaming this one.
if exist "dist\%PRODUCT_EXE%\OptionChain.exe" del /Q "dist\%PRODUCT_EXE%\OptionChain.exe"
if exist "dist\%PRODUCT_EXE%\OptionChain.pyd" del /Q "dist\%PRODUCT_EXE%\OptionChain.pyd"
REM Refresh the client package folder from this build. Without this step the
REM client package kept whatever EXE was there from an earlier build, so the
REM client ZIP and the folder on disk could be different programs entirely.
xcopy /E /I /Y /Q "dist\%PRODUCT_EXE%" "..\client package" >nul
REM xcopy does not delete, so a file that was in the client package from an
REM earlier build and is not in this one would still be there. The old
REM wrongly-named executable is exactly that kind of file, and shipping it
REM means the customer can run the wrong tool and blame this one.
if exist "..\client package\OptionChain.exe" del /Q "..\client package\OptionChain.exe"
if exist "..\client package\OptionChain.pyd" del /Q "..\client package\OptionChain.pyd"
if exist "..\client package\license.json" del /Q "..\client package\license.json"
if exist "..\client package\instance.lock" del /Q "..\client package\instance.lock"
if exist "..\client package\instance_guard.log" del /Q "..\client package\instance_guard.log"
if exist "..\client package\doctor-report.txt" del /Q "..\client package\doctor-report.txt"
REM The old release folder is a month-old build that nothing writes to. It is
REM what package_release.ps1 used to zip, so the customer was sent that.
if exist "dist\OptionChain" rd /S /Q "dist\OptionChain"
powershell -NoProfile -ExecutionPolicy Bypass -File "package_release.ps1"
if errorlevel 1 exit /b 1
exit /b 0
