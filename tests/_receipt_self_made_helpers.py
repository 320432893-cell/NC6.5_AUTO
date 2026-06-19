# 生命周期：持久维护
# 覆盖的业务场景：收款单自制单流程测试的共享 import（receipt_self_made_flow / receipt_new_probe）
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_self_made_*.py

import json
import subprocess
import ctypes

from tools import receipt_self_made_flow as trial
from tools import receipt_new_probe
from tools.receipt_new_probe import (
    detect_self_made_entry_state,
    is_current_visible_control,
)


__all__ = [
    'ctypes',
    'detect_self_made_entry_state',
    'is_current_visible_control',
    'json',
    'receipt_new_probe',
    'subprocess',
    'trial',
]
