@echo off
REM Purpose: read-only speed test for one receipt header semantic lookup.
REM This file is ASCII-only because cmd.exe breaks easily on UTF-8 batch text.
setlocal
set PYTHONIOENCODING=utf-8
cd /d "%~dp0.."

set "LABEL_KEY=customer"
set "PROBE_TIMEOUT=0.35"
set "PROBE_REPEAT=5"

echo.
echo Receipt header semantic lookup speed test - READ ONLY
echo Current directory: %CD%
echo.
echo Field keys:
echo   customer    = customer field
echo   date        = document date
echo   currency    = currency
echo   settlement  = settlement method
echo   finance     = finance organization
echo.
set /p "LABEL_INPUT=Field key, default customer: "
if not "%LABEL_INPUT%"=="" set "LABEL_KEY=%LABEL_INPUT%"
set /p "TIMEOUT_INPUT=Timeout seconds, default 0.35: "
if not "%TIMEOUT_INPUT%"=="" set "PROBE_TIMEOUT=%TIMEOUT_INPUT%"
set /p "REPEAT_INPUT=Repeat count, default 5: "
if not "%REPEAT_INPUT%"=="" set "PROBE_REPEAT=%REPEAT_INPUT%"

echo.
echo Read-only probe: field=%LABEL_KEY% timeout=%PROBE_TIMEOUT%s repeat=%PROBE_REPEAT%
echo Make sure NC is already on the receipt self-made entry page.
echo This script will not type, save, or open a receipt.
echo.
py -3.11 tools\receipt_self_made_flow.py --probe-header-semantic-field "%LABEL_KEY%" --probe-timeout %PROBE_TIMEOUT% --probe-repeat %PROBE_REPEAT% --json
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Exit code: %EXITCODE%
pause
exit /b %EXITCODE%
