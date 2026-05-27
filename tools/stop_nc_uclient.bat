@echo off
setlocal

echo Stopping NC/UClient Java processes...
wmic process where "name='javaw.exe' and commandline like '%%UClient%%'" call terminate

echo Stopping UClient shell...
taskkill /f /im Uclient.exe

echo.
echo Remaining UClient related processes:
wmic process where "commandline like '%%UClient%%'" get Name,ProcessId,ExecutablePath,CommandLine

pause
