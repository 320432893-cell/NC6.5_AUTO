#!/usr/bin/env bash
# 职责：从 WSL/Linux 终端启动 Windows Python，只读测速收款单表头单字段语义定位耗时
# 不做什么：不使用 Linux Python 控制 NC；不写字段、不保存/暂存、不打开收款单
# 允许依赖层：cmd.exe、Windows py launcher、tools/receipt_self_made_flow.py
# 谁不应该 import：不适用，shell 仅作为现场人工启动入口
set -euo pipefail

LABEL="${1:-customer}"
TIMEOUT="${2:-0.35}"
REPEAT="${3:-5}"
WIN_ROOT='H:\python脚本\.venv\nc_auto_v2'

echo "Receipt header semantic lookup speed test - READ ONLY"
echo "field=${LABEL} timeout=${TIMEOUT}s repeat=${REPEAT}"
echo "JAB must use Windows Python; this script calls cmd.exe from WSL."
cmd.exe /c "set PYTHONIOENCODING=utf-8&& cd /d ${WIN_ROOT}&& py -3.11 tools\receipt_self_made_flow.py --probe-header-semantic-field \"${LABEL}\" --probe-timeout ${TIMEOUT} --probe-repeat ${REPEAT} --json"
