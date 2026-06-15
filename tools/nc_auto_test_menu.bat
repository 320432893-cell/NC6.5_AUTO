@echo off
rem ASCII-only launcher. The Chinese menu lives in nc_auto_test_menu.py.
setlocal EnableExtensions

chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set "SCRIPT_DIR=%~dp0"
set "PROJECT_DIR=%SCRIPT_DIR%.."
set "PYTHON_EXE=%PROJECT_DIR%\.venv-local\Scripts\python.exe"

if exist "%PYTHON_EXE%" (
    "%PYTHON_EXE%" "%SCRIPT_DIR%nc_auto_test_menu.py"
) else (
    py -3.11 "%SCRIPT_DIR%nc_auto_test_menu.py"
)

set "RC=%ERRORLEVEL%"
echo.
echo Exit code: %RC%
pause
exit /b %RC%
