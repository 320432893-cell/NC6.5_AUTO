@echo off
setlocal

chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set SCRIPT_DIR=%~dp0
set PROJECT_DIR=%SCRIPT_DIR%..
set PYTHON_EXE=%PROJECT_DIR%\.venv-local\Scripts\python.exe
set ENTRY=%SCRIPT_DIR%receipt_detail_entry.py

if not exist "%PYTHON_EXE%" (
    set PYTHON_EXE=py
    set PYTHON_ARGS=-3.11
) else (
    set PYTHON_ARGS=
)

:menu
cls
echo 收款单明细测试入口
echo.
echo 前置条件：
echo  1. NC 已停在【收款单自制录入】界面
echo  2. 当前没有参照窗口或提示框
echo  3. 本入口不会保存、不会暂存、不会关闭收款单
echo.
echo 请选择测试功能：
echo  1. 写明细主行
echo  2. 写手续费行（默认手续费 10）
echo  3. 只清理第 1 行以外的多余行
echo  4. 显示命令帮助
echo  0. 退出
echo.
set /p CHOICE=请输入编号后回车：

if "%CHOICE%"=="1" goto main_line
if "%CHOICE%"=="2" goto fee_line
if "%CHOICE%"=="3" goto cleanup_rows
if "%CHOICE%"=="4" goto help
if "%CHOICE%"=="0" goto end

echo.
echo 输入无效，请重新选择。
pause
goto menu

:main_line
call :run "%PYTHON_EXE%" %PYTHON_ARGS% "%ENTRY%"
goto menu

:fee_line
set FEE_AMOUNT=10
set /p FEE_AMOUNT=请输入手续费金额，直接回车默认 10：
if "%FEE_AMOUNT%"=="" set FEE_AMOUNT=10
call :run "%PYTHON_EXE%" %PYTHON_ARGS% "%ENTRY%" --fee-only --fee-amount "%FEE_AMOUNT%"
goto menu

:cleanup_rows
call :run "%PYTHON_EXE%" %PYTHON_ARGS% "%ENTRY%" --cleanup-extra-rows-only
goto menu

:help
call :run "%PYTHON_EXE%" %PYTHON_ARGS% "%ENTRY%" --help
goto menu

:run
echo.
echo 即将运行：
echo %*
echo.
echo 请确认 NC 界面已经准备好，然后按任意键开始；按 Ctrl+C 可取消。
pause >nul
echo 开始时间：%DATE% %TIME%
call %*
set "RC=%ERRORLEVEL%"
echo.
echo 结束时间：%DATE% %TIME%
echo 退出码：%RC%
pause
exit /b %RC%

:end
endlocal
exit /b 0
