@echo off
setlocal
set PYTHONIOENCODING=utf-8
cd /d "%~dp0.."
py -3.11 tools\receipt_self_made_flow.py --probe-customer-name-readback --json
echo.
echo Exit code: %ERRORLEVEL%
pause
