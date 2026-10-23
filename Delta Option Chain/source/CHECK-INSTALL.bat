@echo off
REM ---------------------------------------------------------------------
REM Check that this installation can work.
REM
REM It writes a report to doctor-report.txt next to this file and opens
REM it. Run it first if the tool is not behaving, and send the report to
REM support. It needs no Python and no internet beyond the data source.
REM ---------------------------------------------------------------------
"%~dp0delta_option_chain.exe" --doctor
if errorlevel 1 (
  echo.
  echo The self-test found a problem. See doctor-report.txt in this folder.
  pause
)
