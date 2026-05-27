@echo off
setlocal

set JRE_BIN=%LOCALAPPDATA%\UClient\share\java1.7.0_51-x64\bin
set JAB_DLL=%JRE_BIN%\WindowsAccessBridge-64.dll
set PROBE=%~dp0jab_probe.py
set OUT=%~dp0jab_probe_output.txt

"%JRE_BIN%\jabswitch.exe" -enable
py -3.11 "%PROBE%" --all --children --depth 4 --startup-wait 5 --dll "%JAB_DLL%" > "%OUT%" 2>&1
type "%OUT%"

echo.
echo Output saved to: "%OUT%"

pause
