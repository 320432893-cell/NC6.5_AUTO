# 职责：收款单保存(Ctrl+S)、取消重开熔断、确认弹窗判定、父页恢复就绪等待
# 不做什么：不做整单编排、不写明细字段
# 允许依赖层：core JAB、tools 收款定位缓存/模态守卫;谁不应 import：core 层

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from tools.receipt_keyboard_utils import (  # noqa: E402
    foreground_matches_window,
    send_hotkey_alt_y,
    send_hotkey_ctrl_q,
    send_hotkey_ctrl_s,
)
from tools.receipt_modal_guard import (  # noqa: E402
    collect_visible_java_dialogs,
    focus_window,
)
from tools.receipt_new_probe import (  # noqa: E402
    annotate_foreground_root_for_targets,
    collect_receipt_new_windows,
    detect_self_made_entry_state,
    filter_usable_new_buttons,
    find_named_controls_in_windows,
    foreground_info,
    root_hwnd,
)


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

CIRCUIT_BREAKER_RETRY_STEPS = {
    "detail-extra-text-verify",
    "detail-pipeline-verify",
    "detail-exchange-rate-guard",
}

def save_receipt_by_ctrl_s(
    jab,
    scope_hwnd=None,
    timeout=3.0,
    min_samples=3,
    interval=0.1,
    success_samples=1,
    min_observe_seconds=0.0,
):
    page = None
    if not scope_hwnd:
        return {
            "ok": False,
            "triggered": False,
            "reason": "Ctrl+S 保存前未取得收款单窗口句柄",
            "page": page,
        }
    target_hwnd = root_hwnd(scope_hwnd) or scope_hwnd
    maximize = jab.maximize_window_by_handle(target_hwnd)
    guard = foreground_matches_window({"hwnd": target_hwnd})
    if not guard.get("ok"):
        return {
            "ok": False,
            "triggered": False,
            "reason": guard.get("reason") or "当前前台窗口不是目标 NC 窗口",
            "maximize": maximize,
            "guard": guard,
            "page": page,
            "scope_hwnd": scope_hwnd,
            "target_hwnd": target_hwnd,
        }
    try:
        send_hotkey_ctrl_s()
    except Exception as exc:
        return {
            "ok": False,
            "triggered": False,
            "reason": f"Ctrl+S SendInput 触发失败：{type(exc).__name__}: {exc}",
            "guard": guard,
            "page": page,
            "scope_hwnd": scope_hwnd,
            "target_hwnd": target_hwnd,
        }
    wait_parent = wait_receipt_parent_new_ready_after_entry_exit(
        jab,
        timeout=timeout,
        interval=interval,
        success_samples=success_samples,
    )
    wait_parent.setdefault("oracle", {})["name"] = "receipt_parent_new_ready_after_save"
    if wait_parent.get("ok"):
        return {
            "ok": True,
            "triggered": True,
            "hotkey": {"ok": True, "mode": "send_input", "key": "Ctrl+S"},
            "precondition": {
                "page": page,
                "maximize": maximize,
                "foreground_guard": guard,
                "scope_hwnd": scope_hwnd,
                "target_hwnd": target_hwnd,
            },
            "seconds": wait_parent.get("seconds"),
            "samples": wait_parent.get("samples") or [],
            "min_samples": int(min_samples),
            "success_samples": int(success_samples),
            "timeout": float(timeout),
            "min_observe_seconds": float(min_observe_seconds),
            "oracle": wait_parent.get("oracle"),
            "entry_state": wait_parent.get("entry_state"),
            "parent_new_state": wait_parent.get("parent_new_state"),
        }
    return {
        "ok": False,
        "triggered": True,
        "hotkey": {"ok": True, "mode": "send_input", "key": "Ctrl+S"},
        "precondition": {
            "page": page,
            "maximize": maximize,
            "foreground_guard": guard,
            "scope_hwnd": scope_hwnd,
            "target_hwnd": target_hwnd,
        },
        "seconds": wait_parent.get("seconds"),
        "samples": wait_parent.get("samples") or [],
        "min_samples": int(min_samples),
        "success_samples": int(success_samples),
        "timeout": float(timeout),
        "min_observe_seconds": float(min_observe_seconds),
        "reason": "保存后未确认收款单父页【新增】已恢复，不能证明保存成功",
        "oracle": wait_parent.get("oracle"),
        "entry_state": wait_parent.get("entry_state"),
        "parent_new_state": wait_parent.get("parent_new_state"),
    }

def cancel_current_receipt_entry(
    config,
    timeout=3.0,
    interval=0.1,
    confirm_wait=0.8,
):
    jab = JABOperator(config)
    jab.hide_blank_awt_windows_enabled = False
    report = {
        "ok": False,
        "method": "ctrl-q-confirm-alt-y",
        "triggered": False,
        "confirmed": False,
    }
    try:
        jab.ensure_started()
        before_windows = collect_receipt_new_windows(jab)
        before_entry_state = detect_self_made_entry_state(before_windows)
        report["entry_state_before"] = before_entry_state
        if not before_entry_state.get("ok"):
            report["reason"] = "取消前未检测到保存/暂存/取消录入态按钮"
            return report
        target_hwnd = current_receipt_root_from_entry_state(before_entry_state)
        if not target_hwnd:
            report["reason"] = "取消前未取得当前收款单窗口句柄"
            return report
        report["target_hwnd"] = target_hwnd
        maximize = jab.maximize_window_by_handle(target_hwnd)
        guard = foreground_matches_window({"hwnd": target_hwnd})
        report["precondition"] = {
            "maximize": maximize,
            "foreground_guard": guard,
        }
        if not guard.get("ok"):
            report["reason"] = guard.get("reason") or "当前前台窗口不是目标 NC 窗口"
            return report
        dialogs_before = collect_visible_java_dialogs(jab)
        report["dialogs_before_count"] = len(dialogs_before)
        try:
            send_hotkey_ctrl_q()
        except Exception as exc:
            report["reason"] = f"Ctrl+Q SendInput 触发失败：{type(exc).__name__}: {exc}"
            return report
        report["triggered"] = True
        dialog_wait = wait_confirm_cancel_dialog(
            jab,
            dialogs_before,
            timeout=float(confirm_wait or 0.8),
            interval=0.08,
        )
        report["confirm_dialog"] = dialog_wait
        dialog = dialog_wait.get("dialog")
        if not dialog_wait.get("ok") or not dialog:
            report["reason"] = dialog_wait.get("reason") or "未检测到确认取消弹窗"
            return report
        focus = focus_window(dialog.get("hwnd"))
        report["confirm_focus"] = focus
        try:
            send_hotkey_alt_y()
        except Exception as exc:
            report["reason"] = f"Alt+Y SendInput 确认失败：{type(exc).__name__}: {exc}"
            return report
        report["confirmed"] = True
        wait_parent = wait_receipt_parent_new_ready_after_entry_exit(
            jab,
            timeout=timeout,
            interval=interval,
        )
        report["parent_ready_after_cancel"] = wait_parent
        report["ok"] = bool(wait_parent.get("ok"))
        if not report["ok"]:
            report["reason"] = wait_parent.get("reason") or "取消后未确认父页新增可用"
        return report
    finally:
        jab.close()

def should_retry_row_by_cancel_reopen(row_report):
    if not row_report or row_report.get("ok"):
        return False
    if (row_report.get("circuit_breaker") or {}).get("triggered"):
        return False
    if row_report.get("save_attempted"):
        return False
    failed_step = str(row_report.get("failed_step") or "")
    if failed_step not in CIRCUIT_BREAKER_RETRY_STEPS:
        return False
    save_report = row_report.get("save") or {}
    if save_report and not save_report.get("skipped"):
        return False
    return True

def summarize_retry_attempt(row_report):
    return {
        "ok": bool((row_report or {}).get("ok")),
        "excel_row": (row_report or {}).get("excel_row"),
        "failed_step": (row_report or {}).get("failed_step"),
        "reason": (row_report or {}).get("reason"),
        "slow_steps": (row_report or {}).get("slow_steps") or [],
        "detail_exchange_rate_guard": (row_report or {}).get(
            "detail_exchange_rate_guard"
        ),
        "extra_text_verify_failures": (row_report or {}).get(
            "extra_text_verify_failures"
        ),
        "detail_pipeline_verify": (row_report or {}).get("detail_pipeline_verify"),
        "detail_pipeline_verify_after_repair": (row_report or {}).get(
            "detail_pipeline_verify_after_repair"
        ),
    }

def current_receipt_root_from_entry_state(entry_state):
    hits = (entry_state or {}).get("hits") or []
    for hit in hits:
        window = hit.get("window") or {}
        hwnd = window.get("hwnd")
        if hwnd:
            return root_hwnd(hwnd) or hwnd
    return None

def wait_confirm_cancel_dialog(jab, before_dialogs, timeout=0.8, interval=0.08):
    started = time.perf_counter()
    before_keys = {dialog_key(item) for item in before_dialogs or []}
    attempts = []
    while True:
        dialogs = collect_visible_java_dialogs(jab)
        matching = [
            item
            for item in dialogs
            if is_confirm_cancel_dialog(item)
            and (
                dialog_key(item) not in before_keys
                or not before_keys
            )
        ]
        attempts.append(
            {
                "t": round(time.perf_counter() - started, 3),
                "dialog_count": len(dialogs),
                "matching_count": len(matching),
            }
        )
        if matching:
            return {
                "ok": True,
                "seconds": round(time.perf_counter() - started, 3),
                "attempts": attempts,
                "dialog": summarize_dialog_for_report(matching[0]),
            }
        if time.perf_counter() - started >= float(timeout or 0):
            return {
                "ok": False,
                "seconds": round(time.perf_counter() - started, 3),
                "attempts": attempts,
                "reason": "未发现标题为【确认取消】且包含【是(Y)/否(N)】的确认弹窗",
                "last_dialogs": [summarize_dialog_for_report(item) for item in dialogs],
            }
        time.sleep(float(interval or 0.08))

def dialog_key(dialog):
    return (
        (dialog or {}).get("hwnd"),
        (dialog or {}).get("title"),
        (dialog or {}).get("class_name"),
    )

def is_confirm_cancel_dialog(dialog):
    if (dialog or {}).get("class_name") != "SunAwtDialog":
        return False
    if (dialog or {}).get("title") != "确认取消":
        return False
    names = {button.get("name") for button in (dialog or {}).get("buttons") or []}
    return {"是(Y)", "否(N)"} <= names

def summarize_dialog_for_report(dialog):
    if not dialog:
        return None
    return {
        "hwnd": dialog.get("hwnd"),
        "title": dialog.get("title"),
        "class_name": dialog.get("class_name"),
        "pid": dialog.get("pid"),
        "visible": dialog.get("visible"),
        "root_hwnd": dialog.get("root_hwnd"),
        "buttons": [
            {
                "path": button.get("path"),
                "name": button.get("name"),
                "description": button.get("description"),
                "bounds": button.get("bounds"),
            }
            for button in (dialog.get("buttons") or [])
        ],
    }

def wait_receipt_parent_new_ready_after_entry_exit(
    jab,
    timeout=3.0,
    interval=0.1,
    success_samples=1,
):
    started = time.perf_counter()
    samples = []
    strong_success_streak = 0
    last_state = None
    last_parent_new_state = None
    while True:
        sample_started = time.perf_counter()
        windows = collect_receipt_new_windows(jab)
        state = detect_self_made_entry_state(windows)
        parent_new_state = detect_receipt_parent_new_ready(windows)
        last_state = state
        last_parent_new_state = parent_new_state
        strong_success = bool(parent_new_state.get("ok")) and not state.get("ok")
        strong_success_streak = strong_success_streak + 1 if strong_success else 0
        samples.append(
            {
                "sample_index": len(samples) + 1,
                "t": round(time.perf_counter() - started, 3),
                "collect_seconds": round(time.perf_counter() - sample_started, 3),
                "new_candidate_count": parent_new_state.get("candidate_count"),
                "new_usable_count": parent_new_state.get("usable_new_button_count"),
                "entry_ok": bool(state.get("ok")),
                "entry_partial_ok": bool(state.get("partial_ok")),
                "strong_success": strong_success,
                "strong_success_streak": strong_success_streak,
            }
        )
        if strong_success_streak >= int(success_samples or 1):
            return {
                "ok": True,
                "seconds": round(time.perf_counter() - started, 3),
                "samples": samples,
                "oracle": {
                    "name": "receipt_parent_new_ready_after_entry_exit",
                    "ok": True,
                    "evidence": "检测到收款单父页【新增】按钮可用，且保存/暂存/取消三按钮不再同时存在",
                    "parent_new_state": parent_new_state,
                    "self_made_entry_state": state,
                },
                "entry_state": state,
                "parent_new_state": parent_new_state,
            }
        if time.perf_counter() - started >= float(timeout or 0):
            return {
                "ok": False,
                "seconds": round(time.perf_counter() - started, 3),
                "samples": samples,
                "reason": "未确认收款单父页【新增】已恢复",
                "oracle": {
                    "name": "receipt_parent_new_ready_after_entry_exit",
                    "ok": False,
                    "evidence": "需要同时满足：前台收款单父页【新增】按钮可用，且保存/暂存/取消三按钮不再同时存在",
                    "parent_new_state": last_parent_new_state,
                    "self_made_entry_state": last_state,
                },
                "entry_state": last_state,
                "parent_new_state": last_parent_new_state,
            }
        time.sleep(float(interval or 0.1))

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
