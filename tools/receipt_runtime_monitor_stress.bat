@echo off
setlocal

set "ROOT=%~dp0.."
set "OUT_FILE=%TEMP%\receipt_monitor_stress.jsonl"

cd /d "%ROOT%"

py -3.11 tools\receipt_runtime_monitor.py --interval 1 --stress --stall-seconds 20 --out "%OUT_FILE%" %*
if errorlevel 1 python tools\receipt_runtime_monitor.py --interval 1 --stress --stall-seconds 20 --out "%OUT_FILE%" %*
