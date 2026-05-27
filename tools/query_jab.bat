@echo off
setlocal

set PYTHONIOENCODING=utf-8
set JRE_BIN=%LOCALAPPDATA%\UClient\share\java1.7.0_51-x64\bin
set JAB_DLL=%JRE_BIN%\WindowsAccessBridge-64.dll
set PROBE=%~dp0jab_probe.py

if "%~1"=="" (
    echo Usage: query_jab.bat keyword
    exit /b 2
)

py -3.11 "%PROBE%" --all --children --depth 25 --max-children 1000 --query "%~1" --startup-wait 5 --dll "%JAB_DLL%"
