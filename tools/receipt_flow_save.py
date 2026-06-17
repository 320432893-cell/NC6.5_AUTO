# 职责：用键盘 Ctrl+S 触发收款单保存,并以父页【新增】恢复为成功 oracle
# 不做什么：不写入表头/明细字段,不做报告渲染,不做整行编排,不默认保存
# 允许依赖层：tools.receipt_new_probe、tools.receipt_flow_entry_state
# 谁不应该 import：core 层模块不应 import；本模块不应反向 import row_runner

import sys
import time

from tools.receipt_flow_entry_state import extract_entry_scope_hwnd
from tools.receipt_new_probe import (
    annotate_foreground_root_for_targets,
    filter_usable_new_buttons,
    find_named_controls_in_windows,
    foreground_info,
)


class _FlowNamespace:
    # 按调用时从已加载的入口模块取属性：让测试对 tools.receipt_full_flow_entry 上
    # probe_receipt_entry_page / collect_receipt_new_windows /
    # detect_self_made_entry_state / detect_receipt_parent_new_ready /
    # foreground_matches_window 的 monkeypatch 与拆分前一致地生效，
    # 且不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_full_flow_entry"], name)


_flow = _FlowNamespace()


def probe_receipt_entry_page(jab):
    windows = _flow.collect_receipt_new_windows(jab)
    state = _flow.detect_self_made_entry_state(windows)
    scope_hwnd = extract_entry_scope_hwnd({"entry_state": state, "windows": windows})
    return {
        "ok": bool(state.get("ok")),
        "method": "entry-state",
        "scope_hwnd": scope_hwnd,
        "windows": windows,
        "entry_state": state,
    }


def save_receipt_by_ctrl_s(jab, scope_hwnd=None, timeout=1.0):
    page = None
    if not scope_hwnd:
        page = _flow.probe_receipt_entry_page(jab)
        if not page.get("ok"):
            return {
                "ok": False,
                "triggered": False,
                "reason": "Ctrl+S 保存前未确认当前是收款单自制录入页",
                "page": page,
            }
        scope = page.get("scope") or {}
        scope_hwnd = scope.get("scope_hwnd") or page.get("scope_hwnd")
    if not scope_hwnd:
        return {
            "ok": False,
            "triggered": False,
            "reason": "Ctrl+S 保存前未取得收款单窗口句柄",
            "page": page,
        }
    guard = _flow.foreground_matches_window({"hwnd": scope_hwnd})
    if not guard.get("ok"):
        return {
            "ok": False,
            "triggered": False,
            "reason": guard.get("reason") or "当前前台窗口不是目标 NC 窗口",
            "guard": guard,
            "page": page,
        }
    try:
        jab.press_hotkey("ctrl", "s", wait=0)
    except Exception as exc:
        return {
            "ok": False,
            "triggered": False,
            "reason": f"Ctrl+S 键盘热键触发失败：{type(exc).__name__}: {exc}",
            "guard": guard,
            "page": page,
        }
    started = time.perf_counter()
    last_state = None
    while time.perf_counter() - started < timeout:
        windows = _flow.collect_receipt_new_windows(jab)
        state = _flow.detect_self_made_entry_state(windows)
        last_state = state
        parent_new_state = _flow.detect_receipt_parent_new_ready(windows)
        if parent_new_state.get("ok") and not state.get("ok"):
            return {
                "ok": True,
                "triggered": True,
                "hotkey": {"ok": True, "mode": "jab.press_hotkey", "key": "Ctrl+S"},
                "precondition": {
                    "page": page,
                    "foreground_guard": guard,
                },
                "seconds": round(time.perf_counter() - started, 3),
                "oracle": {
                    "name": "receipt_parent_new_ready_after_save",
                    "ok": True,
                    "evidence": "保存后重新检测到收款单录入父页前台【新增】按钮，且保存/暂存/取消三按钮不再同时存在",
                    "parent_new_state": parent_new_state,
                    "self_made_entry_state": state,
                },
                "entry_state": state,
                "parent_new_state": parent_new_state,
            }
        time.sleep(0.2)
    final_windows = _flow.collect_receipt_new_windows(jab)
    parent_new_state = _flow.detect_receipt_parent_new_ready(final_windows)
    return {
        "ok": False,
        "triggered": True,
        "hotkey": {"ok": True, "mode": "jab.press_hotkey", "key": "Ctrl+S"},
        "precondition": {
            "page": page,
            "foreground_guard": guard,
        },
        "seconds": round(time.perf_counter() - started, 3),
        "reason": "保存后未确认收款单父页【新增】已恢复，不能证明保存成功",
        "oracle": {
            "name": "receipt_parent_new_ready_after_save",
            "ok": False,
            "evidence": "需要同时满足：前台收款单父页【新增】按钮可用，且保存/暂存/取消三按钮不再同时存在",
            "parent_new_state": parent_new_state,
            "self_made_entry_state": last_state,
        },
        "entry_state": last_state,
        "parent_new_state": parent_new_state,
    }


def detect_receipt_parent_new_ready(windows):
    foreground = foreground_info()
    buttons = find_named_controls_in_windows(
        windows,
        "新增",
        role=None,
        class_name="SunAwtFrame",
        require_action=True,
    )
    annotate_foreground_root_for_targets(buttons, foreground)
    usable = filter_usable_new_buttons(buttons, foreground)
    return {
        "ok": bool(usable),
        "foreground": foreground,
        "usable_new_button_count": len(usable),
        "usable_new_buttons": [
            {
                "window": item.get("window"),
                "control": {
                    key: (item.get("control") or {}).get(key)
                    for key in ("name", "description", "role", "states", "path")
                },
            }
            for item in usable[:3]
        ],
        "candidate_count": len(buttons),
    }
