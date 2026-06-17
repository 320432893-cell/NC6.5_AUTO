# 职责：定位/点击"新增"按钮、打开新增菜单
# 不做什么：不做 CLI 解析/不起 JABOperator(那是 receipt_new_probe 主入口)
# 允许依赖层：标准库、core JAB(经 jab 参数)、tools.jab_probe、tools.receipt_new_* 同层
# 谁不应该 import：core 层模块不应 import

import ctypes
import os
import sys
import threading
import time




class _ProbeNamespace:
    # 调用时从已加载的 receipt_new_probe 读顶层函数,使测试对
    # tools.receipt_new_probe.<name> 的 monkeypatch 与拆分前一致生效,且不在加载期 import 成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_new_probe"], name)


_probe = _ProbeNamespace()


def open_new_menu_with_known_buttons(
    jab, args, buttons, all_buttons=None, foreground=None
):
    if args.method == "button" and buttons:
        target = buttons[0]
        action_report = _probe.trigger_button_async(
            jab,
            target["window"]["hwnd"],
            target["control"]["path"],
            action_name=args.action,
            return_timeout=args.return_timeout,
            target=target,
        )
        if action_report.get("ok"):
            return action_report
        return {
            "ok": False,
            "method": "button",
            "reason": "新增按钮 action 失败；正式收款流程不回退 Ctrl+N",
            "button_action": action_report,
        }
    if args.method == "button" and all_buttons:
        return {
            "ok": False,
            "method": "button",
            "reason": "找到新增候选，但没有前台收款单窗口内可用的新增按钮；正式流程不回退 Ctrl+N",
            "button_reason": "new button candidates were found, but none were usable in the foreground NC window",
            "rejected_count": len(all_buttons),
            "rejected": _probe.summarize_candidates(all_buttons[:20]),
        }
    if args.method == "button":
        return {
            "ok": False,
            "method": "button",
            "reason": "未找到新增按钮；正式收款流程不回退 Ctrl+N",
            "button_reason": "new button not found",
            "rejected_count": 0,
            "rejected": [],
        }
    return _probe.open_new_menu(jab, args)



def open_new_menu_with_ctrl_n(foreground):
    guard = _probe.foreground_nc_guard(foreground)
    if not guard.get("ok"):
        return {
            "ok": False,
            "method": "ctrl+n-fallback",
            "reason": "foreground is not NC; Ctrl+N not sent",
            "foreground_guard": guard,
        }
    try:
        _probe.send_ctrl_n()
    except Exception as exc:
        return {
            "ok": False,
            "method": "ctrl+n-fallback",
            "reason": f"Ctrl+N send failed: {exc!r}",
            "foreground_guard": guard,
        }
    return {
        "ok": True,
        "method": "ctrl+n-fallback",
        "foreground_guard": guard,
    }



def foreground_nc_guard(foreground):
    if os.name != "nt":
        return {"ok": False, "reason": "Windows only", "foreground": foreground}
    if not foreground:
        return {
            "ok": False,
            "reason": "missing foreground window",
            "foreground": foreground,
        }
    ok = bool(
        foreground.get("class_name") == "YonyouUWnd"
        or foreground.get("root_class_name") == "YonyouUWnd"
    )
    return {
        "ok": ok,
        "reason": None if ok else "foreground root is not YonyouUWnd",
        "foreground": foreground,
    }



def send_ctrl_n():
    user32 = ctypes.windll.user32
    vk_control = 0x11
    vk_n = 0x4E
    try:
        user32.keybd_event(vk_control, 0, 0, 0)
        time.sleep(0.02)
        user32.keybd_event(vk_n, 0, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(vk_n, 0, 2, 0)
    finally:
        user32.keybd_event(vk_control, 0, 2, 0)



def open_new_menu(jab, args):
    if args.method == "probe-button":
        return {"ok": True, "method": "probe-button"}
    if args.method == "button":
        foreground = _probe.foreground_info()
        buttons = _probe.find_new_buttons(jab, args.name, args.role, args.class_name)
        if not buttons:
            fallback = _probe.open_new_menu_with_ctrl_n(foreground)
            fallback.update({"button_reason": "new button not found"})
            return fallback
        return _probe.open_new_menu_with_known_buttons(
            jab, args, buttons, all_buttons=buttons, foreground=foreground
        )

    ok = jab.do_action_by_path(
        args.path,
        title=args.title,
        class_name=args.class_name,
        name=args.name,
        role=args.role,
        action_name=args.action,
        wait=0,
        timeout=2.0,
        require_showing=False,
        require_valid_bounds=False,
        cleanup_blank_awt=False,
    )
    return {"ok": bool(ok), "method": "action-path", "path": args.path}



def wait_for_self_made_popup(jab, before, timeout=0.8, interval=0.08):
    start = time.perf_counter()
    deadline = time.perf_counter() + max(float(timeout or 0), 0)
    attempts = 0
    last_windows = []
    while True:
        attempts += 1
        last_windows = _probe.collect_new_visible_popup_windows(jab, before)
        popup = _probe.find_new_visible_popup(before, last_windows)
        if popup:
            return {
                "ok": True,
                "attempts": attempts,
                "wait_seconds": _probe.elapsed(start),
                "popup": popup,
                "windows": last_windows,
            }
        remaining = deadline - time.perf_counter()
        if remaining <= 0:
            break
        time.sleep(min(interval, remaining))
    full_windows = _probe.collect_receipt_new_windows(jab)
    popup = _probe.find_new_visible_popup(before, full_windows)
    return {
        "ok": bool(popup),
        "attempts": attempts,
        "wait_seconds": max(float(timeout or 0), 0),
        "popup": popup,
        "windows": full_windows if popup else last_windows,
    }



def find_new_buttons(jab, name_query="新增", role=None, class_name=None):
    buttons = _probe.find_named_controls(
        jab,
        name_query=name_query,
        role=role,
        class_name=class_name,
        require_action=True,
    )
    foreground = _probe.foreground_info()
    _probe.annotate_foreground_root_for_targets(buttons, foreground)
    buttons = _probe.filter_usable_new_buttons(buttons, foreground)
    buttons.sort(key=new_button_priority)
    return buttons



def new_button_priority(item):
    control = item.get("control") or {}
    states = control.get("states", "")
    bounds = control.get("bounds") or []
    is_foreground_root = bool((item.get("window") or {}).get("is_foreground_root"))
    has_valid_size = _probe.has_valid_bounds(bounds)
    is_showing = "showing" in states.lower()
    desc = control.get("description") or ""
    is_plain_new = desc == "新增(Ctrl+N)"
    return (
        not is_foreground_root,
        not is_showing,
        not has_valid_size,
        not is_plain_new,
    )



def filter_usable_new_buttons(buttons, foreground=None):
    buttons = [
        item
        for item in buttons or []
        if _probe.is_current_visible_control(item.get("control") or {})
    ]
    if not buttons:
        return []
    if foreground and foreground.get("root"):
        foreground_buttons = [
            item
            for item in buttons
            if (item.get("window") or {}).get("is_foreground_root")
        ]
        if foreground_buttons:
            return foreground_buttons
    return buttons



def trigger_button_async(
    jab, hwnd, path, action_name=None, return_timeout=0.2, target=None
):
    result = jab.find_context_by_path_once(
        path,
        scope_hwnd=hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    context, vm_id, owned, _window_info = result
    if not context:
        return {
            "ok": False,
            "method": "button",
            "reason": "button path not found",
            "target": target,
        }

    status = {"returned": False, "ok": None, "exception": None}

    def run_action():
        try:
            status["ok"] = jab.do_action(
                vm_id,
                context,
                action_name=action_name,
                cleanup_blank_awt=False,
            )
        except Exception as exc:
            status["exception"] = repr(exc)
        finally:
            status["returned"] = True

    thread = threading.Thread(target=run_action, daemon=True)
    thread.start()
    thread.join(return_timeout)
    returned = not thread.is_alive()
    if returned:
        jab.release_contexts(vm_id, owned)
    return {
        "ok": True if not returned else bool(status["ok"]),
        "method": "button",
        "path": path,
        "target": target,
        "action_returned_within_timeout": returned,
        "action_status": status,
    }
