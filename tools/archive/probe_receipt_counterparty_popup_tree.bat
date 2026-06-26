@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0..\.."
set "PYTHONIOENCODING=utf-8"
set "PYTHONPATH=%CD%"
py -3.11 tools\archive\probe_receipt_counterparty_popup_tree.py --json
set "EXITCODE=%ERRORLEVEL%"
echo.
echo Exit code: %EXITCODE%
pause
exit /b %EXITCODE%
