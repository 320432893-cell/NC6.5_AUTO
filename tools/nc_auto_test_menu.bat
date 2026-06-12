@echo off
rem 职责：项目级测试菜单，按风险分组启动现有正式/诊断入口。
rem 不做什么：不实现业务规则、不绕过 Python 入口二次确认、不恢复已清理的收款单真实保存 T0 脚本。
rem 允许依赖层：tools 下 CLI/bat 入口。
rem 谁不应该 import：本文件是 Windows 批处理入口，不作为 Python 模块依赖。
setlocal EnableExtensions

chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set "SCRIPT_DIR=%~dp0"
set "PROJECT_DIR=%SCRIPT_DIR%.."
set "PYTHON_EXE=%PROJECT_DIR%\.venv-local\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    set "PYTHON_EXE=py"
    set "PYTHON_ARGS=-3.11"
) else (
    set "PYTHON_ARGS="
)

:menu
cls
echo NC 自动化项目级测试入口
echo.
echo Python: %PYTHON_EXE% %PYTHON_ARGS%
echo 项目目录: %PROJECT_DIR%
echo.
echo 风险分组：
echo  A. 只读/本地检查：不写 Excel，不点击保存
echo  B. 写 Excel：会写 Sheet2 或 Sheet1 状态列
echo  C. 真实 NC 不保存：会点击/填写当前 NC 页面，但不保存业务单据
echo  D. 真实保存：会保存凭证或收款单，必须二次确认
echo.
echo 请选择功能：
echo  1. 工程检查 tools/check.py changed
echo  2. 凭证计划预览（不保存）
echo  3. 凭证真实生成并保存
echo  4. 凭证回填凭证号
echo  5. 凭证切到已生成/正式单据列表（可记录 perf）
echo  6. 收款单本地预检（不写 Excel）
echo  7. 收款单本地预检并写 Sheet2
echo  8. 收款单完整流程测试（默认不保存，消费 ReceiptPlanRow）
echo  9. 收款单完整流程真实保存（高风险）
echo 10. 收款单单行无保存试填（旧分阶段入口）
echo 11. 收款单明细测试子菜单（主行/手续费/删多余行，不保存）
echo 12. 收款单查询读取结果（不写 Excel）
echo 13. 收款单历史查重写回 Sheet1 状态列
echo  0. 退出
echo.
set /p CHOICE=请输入编号后回车：

if "%CHOICE%"=="1" goto check_changed
if "%CHOICE%"=="2" goto voucher_plan
if "%CHOICE%"=="3" goto voucher_generate
if "%CHOICE%"=="4" goto voucher_backfill
if "%CHOICE%"=="5" goto voucher_switch_generated
if "%CHOICE%"=="6" goto receipt_local_plan
if "%CHOICE%"=="7" goto receipt_write_sheet2
if "%CHOICE%"=="8" goto receipt_full_no_save
if "%CHOICE%"=="9" goto receipt_full_save
if "%CHOICE%"=="10" goto receipt_no_save_trial
if "%CHOICE%"=="11" goto receipt_detail_menu
if "%CHOICE%"=="12" goto receipt_query_read
if "%CHOICE%"=="13" goto receipt_query_write_back
if "%CHOICE%"=="0" goto end

echo.
echo 输入无效，请重新选择。
pause
goto menu

:check_changed
set "RUN_TITLE=工程检查 changed"
set "RUN_LEVEL=A 只读/本地检查"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\check.py" changed
goto menu

:voucher_plan
call :prompt_range
set "RUN_TITLE=凭证计划预览"
set "RUN_LEVEL=A 只读/不保存"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\jab_batch.py" plan %RANGE_ARGS%
goto menu

:voucher_generate
echo.
echo [高风险] 该功能会真实点击 NC 生成并保存凭证，也会写 Excel 状态。
echo Python 入口仍会要求输入 yes；这里先做菜单层确认。
set /p CONFIRM_SAVE=确认继续请输入 SAVE：
if /I not "%CONFIRM_SAVE%"=="SAVE" (
    echo 已取消。
    pause
    goto menu
)
call :prompt_range
set "MAX_BATCH_ARG="
set /p MAX_BATCHES=最多执行几个批次，直接回车不限制：
if not "%MAX_BATCHES%"=="" set "MAX_BATCH_ARG=--max-batches %MAX_BATCHES%"
set "PERF_ARG="
set /p PERF_ON=是否记录 perf JSONL？输入 y 开启：
if /I "%PERF_ON%"=="y" set "PERF_ARG=--perf --perf-label voucher-generate-menu"
set "RUN_TITLE=凭证真实生成并保存"
set "RUN_LEVEL=D 真实保存"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\jab_batch.py" generate %RANGE_ARGS% %MAX_BATCH_ARG% %PERF_ARG%
goto menu

:voucher_backfill
call :prompt_range
set "BACKFILL_PAGE_ARG="
set /p BACKFILL_MANUAL_PAGE=不自动从待生成切已生成？输入 y 开启：
if /I "%BACKFILL_MANUAL_PAGE%"=="y" set "BACKFILL_PAGE_ARG=--no-backfill-"
if /I "%BACKFILL_MANUAL_PAGE%"=="y" set "BACKFILL_PAGE_ARG=%BACKFILL_PAGE_ARG%auto-switch"
set "RUN_TITLE=凭证回填凭证号"
set "RUN_LEVEL=B 写 Excel + 读取 NC"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\jab_batch.py" backfill %RANGE_ARGS% %BACKFILL_PAGE_ARG%
goto menu

:voucher_switch_generated
set "GENERATED_DATE_ARG="
set /p GENERATED_DATE=目的业务日期 YYYY-MM-DD，直接回车用配置或当天：
if not "%GENERATED_DATE%"=="" set "GENERATED_DATE_ARG=--generated-date %GENERATED_DATE%"
set "PERF_ARG="
set /p PERF_ON=是否记录 perf JSONL？输入 y 开启：
if /I "%PERF_ON%"=="y" set "PERF_ARG=--perf --perf-label switch-generated-menu"
set "RUN_TITLE=凭证切到已生成/正式单据列表"
set "RUN_LEVEL=C 真实 NC 不保存"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\jab_batch.py" switch-generated %GENERATED_DATE_ARG% %PERF_ARG%
goto menu

:receipt_local_plan
set "RUN_TITLE=收款单本地预检"
set "RUN_LEVEL=A 只读/不写 Excel/不碰 NC"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\receipt_entry_check.py"
goto menu

:receipt_write_sheet2
echo.
echo [写 Excel] 该功能会追加/维护 Sheet2：收款单自动化结果。
set /p CONFIRM_SHEET2=确认写 Sheet2 请输入 WRITE：
if /I not "%CONFIRM_SHEET2%"=="WRITE" (
    echo 已取消。
    pause
    goto menu
)
set "VALIDATION_ARG="
set /p SKIP_INVALID=是否跳过异常行继续生成可运行计划？输入 y 开启：
if /I "%SKIP_INVALID%"=="y" set "VALIDATION_ARG=--validation-mode skip_invalid_rows"
set "RUN_TITLE=收款单本地预检并写 Sheet2"
set "RUN_LEVEL=B 写 Excel"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\receipt_entry_check.py" %VALIDATION_ARG% --write
goto menu

:receipt_full_no_save
call :prompt_receipt_full_flow
set "PLAN_SHEET_ARG="
set /p WRITE_PLAN=运行前是否写 Sheet2 本地预检结果？输入 y 开启：
if /I "%WRITE_PLAN%"=="y" set "PLAN_SHEET_ARG=--write-plan-sheet"
set "RUN_TITLE=收款单完整流程测试（不保存）"
set "RUN_LEVEL=C 真实 NC 不保存"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\receipt_full_flow_entry.py" %RECEIPT_FULL_ARGS% %PLAN_SHEET_ARG%
goto menu

:receipt_full_save
echo.
echo [高风险] 该功能会消费 ReceiptPlanRow，开自制收款单，填写表头/明细/手续费，并真实保存。
echo 保存前请确认测试单据可清理、NC 已在收款单录入页、Excel/WPS 未占用。
set /p CONFIRM_RECEIPT_SAVE=确认真实保存请输入 SAVE：
if /I not "%CONFIRM_RECEIPT_SAVE%"=="SAVE" (
    echo 已取消。
    pause
    goto menu
)
call :prompt_receipt_full_flow
set "RUN_TITLE=收款单完整流程真实保存"
set "RUN_LEVEL=D 真实保存"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\receipt_full_flow_entry.py" %RECEIPT_FULL_ARGS% --save --yes-i-understand
goto menu

:receipt_no_save_trial
echo.
echo [真实 NC 不保存] 要求 NC 在收款单录入页，脚本会尝试开自制单并填写字段，但不会保存/暂存。
set /p RECEIPT_ROW=请输入 Excel 行号：
if "%RECEIPT_ROW%"=="" (
    echo 未输入行号，已取消。
    pause
    goto menu
)
set "TRIAL_ARGS=%RECEIPT_ROW% --open-self-made"
set /p FILL_DETAIL=是否继续填明细？输入 y 开启：
if /I "%FILL_DETAIL%"=="y" set "TRIAL_ARGS=%TRIAL_ARGS% --fill-detail"
set "RUN_TITLE=收款单单行无保存试填"
set "RUN_LEVEL=C 真实 NC 不保存"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\receipt_self_made_fill_trial.py" %TRIAL_ARGS%
goto menu

:receipt_detail_menu
set "RUN_TITLE=收款单明细测试子菜单"
set "RUN_LEVEL=C 真实 NC 不保存"
call :run_logged "%SCRIPT_DIR%receipt_detail_test_menu.bat"
goto menu

:receipt_query_read
call :prompt_receipt_query
set "RUN_TITLE=收款单查询读取结果"
set "RUN_LEVEL=C 读取 NC 不写 Excel"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\receipt_query_fill.py" --org-code %ORG_CODE% --date-from %DATE_FROM% --date-to %DATE_TO% --confirm --read-results --max-rows %MAX_ROWS% --max-cols %MAX_COLS%
goto menu

:receipt_query_write_back
echo.
echo [写 Excel] 这是历史查重/诊断入口，会写 Sheet1 是否NC已做过等状态列，不是新批量录入主线。
set /p CONFIRM_WRITE_BACK=确认写回 Sheet1 请输入 WRITEBACK：
if /I not "%CONFIRM_WRITE_BACK%"=="WRITEBACK" (
    echo 已取消。
    pause
    goto menu
)
call :prompt_receipt_query
set "INCLUDE_FILLED_ARG="
set /p INCLUDE_FILLED=是否覆盖已有状态列？输入 y 开启：
if /I "%INCLUDE_FILLED%"=="y" set "INCLUDE_FILLED_ARG=--include-filled-status"
set "RUN_TITLE=收款单历史查重写回 Sheet1"
set "RUN_LEVEL=B 写 Excel + 读取 NC"
call :run_logged "%PYTHON_EXE%" %PYTHON_ARGS% "%PROJECT_DIR%\tools\receipt_query_fill.py" --org-code %ORG_CODE% --date-from %DATE_FROM% --date-to %DATE_TO% --confirm --dry-run-match --write-back %INCLUDE_FILLED_ARG% --max-rows %MAX_ROWS% --max-cols %MAX_COLS%
goto menu

:prompt_range
set "RANGE_ARGS="
set /p LIMIT=仅处理前 N 条，直接回车不限制：
if not "%LIMIT%"=="" set "RANGE_ARGS=%RANGE_ARGS% --limit %LIMIT%"
set /p START_ROW=Excel 起始行，直接回车不限制：
if not "%START_ROW%"=="" set "RANGE_ARGS=%RANGE_ARGS% --start-row %START_ROW%"
set /p END_ROW=Excel 结束行，直接回车不限制：
if not "%END_ROW%"=="" set "RANGE_ARGS=%RANGE_ARGS% --end-row %END_ROW%"
exit /b 0

:prompt_receipt_full_flow
set "RECEIPT_FULL_ARGS="
set /p RECEIPT_FULL_ROW=指定 Sheet1 行号，直接回车自动取第一条通过预检行：
if not "%RECEIPT_FULL_ROW%"=="" set "RECEIPT_FULL_ARGS=%RECEIPT_FULL_ARGS% --excel-row %RECEIPT_FULL_ROW%"
set "RECEIPT_FULL_LIMIT=1"
set /p RECEIPT_FULL_LIMIT=最多测试几行，直接回车默认 1：
if "%RECEIPT_FULL_LIMIT%"=="" set "RECEIPT_FULL_LIMIT=1"
set "RECEIPT_FULL_ARGS=%RECEIPT_FULL_ARGS% --limit %RECEIPT_FULL_LIMIT%"
set "RECEIPT_START_DELAY=2"
set /p RECEIPT_START_DELAY=启动前等待秒数，直接回车默认 2：
if "%RECEIPT_START_DELAY%"=="" set "RECEIPT_START_DELAY=2"
set "RECEIPT_FULL_ARGS=%RECEIPT_FULL_ARGS% --start-delay %RECEIPT_START_DELAY%"
exit /b 0

:prompt_receipt_query
set "ORG_CODE=A001"
set /p ORG_CODE=主体编码，直接回车默认 A001：
if "%ORG_CODE%"=="" set "ORG_CODE=A001"
set "DATE_FROM=2026-05-01"
set /p DATE_FROM=开始日期 YYYY-MM-DD，直接回车默认 2026-05-01：
if "%DATE_FROM%"=="" set "DATE_FROM=2026-05-01"
set "DATE_TO=2026-06-02"
set /p DATE_TO=结束日期 YYYY-MM-DD，直接回车默认 2026-06-02：
if "%DATE_TO%"=="" set "DATE_TO=2026-06-02"
set "MAX_ROWS=600"
set /p MAX_ROWS=最多读取行数，直接回车默认 600：
if "%MAX_ROWS%"=="" set "MAX_ROWS=600"
set "MAX_COLS=140"
set /p MAX_COLS=最多读取列数，直接回车默认 140：
if "%MAX_COLS%"=="" set "MAX_COLS=140"
exit /b 0

:run_logged
echo.
echo 功能：%RUN_TITLE%
echo 风险：%RUN_LEVEL%
echo 开始时间：%DATE% %TIME%
echo 命令：%*
echo.
echo 请确认当前 NC/Excel 状态满足该功能前置条件；按任意键开始，按 Ctrl+C 可取消。
pause >nul
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
