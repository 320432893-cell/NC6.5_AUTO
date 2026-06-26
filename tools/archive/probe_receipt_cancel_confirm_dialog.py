import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_keyboard_utils import (  # noqa: E402
    VK_CONTROL,
    VK_MENU,
    send_virtual_key,
)
from tools.receipt_modal_guard import (  # noqa: E402
    collect_visible_java_dialogs,
    focus_window,
)
from tools.receipt_full_flow_entry import detect_receipt_parent_new_ready  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    annotate_foreground_root,
    annotate_foreground_root_for_targets,
    collect_receipt_new_windows,
    detect_self_made_entry_state,
    filter_usable_new_buttons,
    find_named_controls_in_windows,
    foreground_info,
    is_current_visible_control,
    trigger_button_async,
)

VK_Q = 0x51
VK_Y = 0x59
VK_N = 0x4E


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Probe the confirmation dialog opened by the NC receipt entry "
            "Cancel(Ctrl+Q) button. By default this is read-only."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument(
        "--trigger",
        action="store_true",
        help="Click the entry-page Cancel(Ctrl+Q) button and inspect the dialog.",
    )
    parser.add_argument(
        "--trigger-method",
        choices=("ctrl-q", "button-action"),
        default="ctrl-q",
        help=(
            "ctrl-q sends the hotkey after locating the Cancel button; "
            "button-action uses JAB action and may block on modal dialogs."
        ),
    )
    parser.add_argument(
        "--confirm",
        choices=("none", "yes", "no"),
        default="none",
        help=(
            "After a safe confirmation dialog match, optionally send Alt+Y or Alt+N."
        ),
    )
    parser.add_argument("--wait", type=float, default=0.8)
    parser.add_argument("--return-timeout", type=float, default=0.2)
    parser.add_argument("--max-depth", type=int, default=25)
    parser.add_argument("--max-children", type=int, default=1000)
    parser.add_argument("--json", action="store_true")
    return parser.parse_args(argv)


def main(argv=None):
    stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
    if callable(stdout_reconfigure):
        stdout_reconfigure(encoding="utf-8", errors="replace")
    args = parse_args(argv)
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    report = {
        "ok": False,
        "read_only": not bool(args.trigger),
        "trigger": bool(args.trigger),
        "requested": {
            "wait": float(args.wait),
            "return_timeout": float(args.return_timeout),
            "max_depth": int(args.max_depth),
            "max_children": int(args.max_children),
            "trigger_method": args.trigger_method,
            "confirm": args.confirm,
        },
        "foreground": None,
        "entry_state": None,
        "cancel_buttons": [],
        "chosen_cancel_button": None,
        "dialogs_before": [],
        "trigger_action": None,
        "dialogs_after": [],
        "new_or_changed_dialogs": [],
        "confirm_action": None,
        "dialogs_after_confirm": [],
        "entry_state_after_confirm": None,
        "parent_new_state_after_confirm": None,
        "diagnosis": None,
    }
    try:
        jab.ensure_started()
        report["foreground"] = foreground_info()
        windows = collect_receipt_new_windows(
            jab,
            max_depth=int(args.max_depth),
            max_children=int(args.max_children),
        )
        annotate_foreground_root(windows, report["foreground"])
        report["entry_state"] = detect_self_made_entry_state(windows)
        buttons = find_cancel_buttons(windows, report["foreground"])
        report["cancel_buttons"] = summarize_targets(buttons)
        chosen = choose_cancel_button(buttons)
        report["chosen_cancel_button"] = summarize_target(chosen) if chosen else None
        report["dialogs_before"] = collect_visible_java_dialogs(jab)

        if not chosen:
            report["diagnosis"] = {
                "ok": False,
                "reason": "未找到当前前台收款单录入页可用的取消(Ctrl+Q)按钮",
            }
            return finish(report, args)

        if not args.trigger:
            report["ok"] = True
            report["diagnosis"] = {
                "ok": True,
                "reason": "已找到取消按钮；未传 --trigger，因此没有打开确认窗口",
            }
            return finish(report, args)

        report["trigger_action"] = trigger_cancel(jab, chosen, args)
        time.sleep(max(float(args.wait or 0), 0.0))
        report["dialogs_after"] = collect_visible_java_dialogs(jab)
        report["new_or_changed_dialogs"] = diff_dialogs(
            report["dialogs_before"], report["dialogs_after"]
        )
        report["ok"] = bool(report["new_or_changed_dialogs"] or report["dialogs_after"])
        report["diagnosis"] = diagnose_dialog_probe(report)
        if report["ok"] and args.confirm != "none":
            report["confirm_action"] = confirm_cancel_dialog(jab, report, args.confirm)
            time.sleep(max(float(args.wait or 0), 0.0))
            report["dialogs_after_confirm"] = collect_visible_java_dialogs(jab)
            after_windows = collect_receipt_new_windows(
                jab,
                max_depth=int(args.max_depth),
                max_children=int(args.max_children),
            )
            annotate_foreground_root(after_windows, foreground_info())
            report["entry_state_after_confirm"] = detect_self_made_entry_state(
                after_windows
            )
            report["parent_new_state_after_confirm"] = detect_receipt_parent_new_ready(
                after_windows
            )
            report["diagnosis"] = diagnose_confirm_probe(report, args.confirm)
            report["ok"] = bool(report["diagnosis"].get("ok"))
    finally:
        jab.close()
    return finish(report, args)


def find_cancel_buttons(windows, foreground):
    candidates = find_named_controls_in_windows(
        windows,
        name_query="取消",
        role="push button",
        require_action=True,
    )
    annotate_foreground_root_for_targets(candidates, foreground)
    candidates = [
        item
        for item in candidates
        if is_current_visible_control(item.get("control") or {})
    ]
    foreground_items = [
        item
        for item in candidates
        if (item.get("window") or {}).get("is_foreground_root")
    ]
    if foreground_items:
        candidates = foreground_items
    return sorted(candidates, key=cancel_button_priority)


def cancel_button_priority(item):
    control = item.get("control") or {}
    text = f"{control.get('name', '')} {control.get('description', '')}"
    exact_entry = "取消(Ctrl+Q)" in text
    return (
        not exact_entry,
        not bool((item.get("window") or {}).get("is_foreground_root")),
        control.get("path") or "",
    )


def choose_cancel_button(buttons):
    for item in buttons or []:
        control = item.get("control") or {}
        text = f"{control.get('name', '')} {control.get('description', '')}"
        if "取消(Ctrl+Q)" in text:
            return item
    return buttons[0] if buttons else None


def choose_click_action(actions):
    actions = actions or []
    for preferred in ("单击", "click"):
        if preferred in actions:
            return preferred
    return actions[0] if actions else None


def trigger_cancel(jab, chosen, args):
    if args.trigger_method == "button-action":
        return trigger_button_async(
            jab,
            chosen["window"]["hwnd"],
            chosen["control"]["path"],
            action_name=choose_click_action(chosen["control"].get("actions", [])),
            return_timeout=float(args.return_timeout),
            target=chosen,
        )
    focus = focus_window(chosen["window"]["hwnd"])
    try:
        send_hotkey_ctrl_q()
    except Exception as exc:
        return {
            "ok": False,
            "method": "ctrl-q",
            "focus": focus,
            "target": summarize_target(chosen),
            "reason": f"Ctrl+Q SendInput 失败：{type(exc).__name__}: {exc}",
        }
    return {
        "ok": True,
        "method": "ctrl-q",
        "focus": focus,
        "target": summarize_target(chosen),
    }


def send_hotkey_ctrl_q():
    send_virtual_key(VK_CONTROL, key_up=False)
    send_virtual_key(VK_Q, key_up=False)
    send_virtual_key(VK_Q, key_up=True)
    send_virtual_key(VK_CONTROL, key_up=True)


def send_hotkey_alt_y():
    send_virtual_key(VK_MENU, key_up=False)
    send_virtual_key(VK_Y, key_up=False)
    send_virtual_key(VK_Y, key_up=True)
    send_virtual_key(VK_MENU, key_up=True)


def send_hotkey_alt_n():
    send_virtual_key(VK_MENU, key_up=False)
    send_virtual_key(VK_N, key_up=False)
    send_virtual_key(VK_N, key_up=True)
    send_virtual_key(VK_MENU, key_up=True)


def confirm_cancel_dialog(jab, report, choice):
    dialogs = find_confirm_cancel_dialogs(
        report.get("new_or_changed_dialogs") or report.get("dialogs_after") or []
    )
    if not dialogs:
        return {
            "ok": False,
            "choice": choice,
            "reason": "未找到结构匹配的【确认取消】弹窗，拒绝发送确认键",
        }
    dialog = dialogs[0]
    focus = focus_window(dialog.get("hwnd"))
    try:
        if choice == "yes":
            send_hotkey_alt_y()
            method = "Alt+Y"
        elif choice == "no":
            send_hotkey_alt_n()
            method = "Alt+N"
        else:
            return {"ok": False, "choice": choice, "reason": "unsupported choice"}
    except Exception as exc:
        return {
            "ok": False,
            "choice": choice,
            "dialog": summarize_dialog(dialog),
            "focus": focus,
            "reason": f"确认快捷键发送失败：{type(exc).__name__}: {exc}",
        }
    return {
        "ok": True,
        "choice": choice,
        "method": method,
        "dialog": summarize_dialog(dialog),
        "focus": focus,
    }


def find_confirm_cancel_dialogs(dialogs):
    result = []
    for dialog in dialogs or []:
        if dialog.get("class_name") != "SunAwtDialog":
            continue
        if dialog.get("title") != "确认取消":
            continue
        names = {button.get("name") for button in dialog.get("buttons") or []}
        if {"是(Y)", "否(N)"} <= names:
            result.append(dialog)
    return result


def summarize_dialog(dialog):
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


def summarize_targets(targets, limit=20):
    return [summarize_target(item) for item in (targets or [])[:limit]]


def summarize_target(item):
    if not item:
        return None
    return {
        "window": {
            key: (item.get("window") or {}).get(key)
            for key in (
                "hwnd",
                "title",
                "class_name",
                "visible",
                "root_hwnd",
                "is_foreground_root",
            )
        },
        "control": summarize_control(item.get("control") or {}),
    }


def summarize_control(control):
    return {
        "path": control.get("path"),
        "role": control.get("role"),
        "name": control.get("name"),
        "description": control.get("description"),
        "states": control.get("states"),
        "bounds": control.get("bounds"),
        "accessibleAction": control.get("accessibleAction"),
        "actions": control.get("actions") or [],
    }


def dialog_signature(item):
    buttons = tuple(
        (
            button.get("path"),
            button.get("name"),
            button.get("description"),
            tuple(button.get("bounds") or []),
        )
        for button in item.get("buttons", [])
    )
    return (
        item.get("hwnd"),
        item.get("title"),
        item.get("class_name"),
        item.get("visible"),
        item.get("root_hwnd"),
        buttons,
    )


def diff_dialogs(before, after):
    before_signatures = {item.get("hwnd"): dialog_signature(item) for item in before}
    changed = []
    for item in after or []:
        hwnd = item.get("hwnd")
        if hwnd not in before_signatures or before_signatures[hwnd] != dialog_signature(
            item
        ):
            changed.append(item)
    return changed


def diagnose_dialog_probe(report):
    trigger_action = report.get("trigger_action") or {}
    dialogs = report.get("new_or_changed_dialogs") or report.get("dialogs_after") or []
    if not trigger_action.get("ok"):
        return {
            "ok": False,
            "reason": "录入页取消按钮 action 未成功返回；需要检查 cancel button target",
        }
    if not dialogs:
        return {
            "ok": False,
            "reason": "取消按钮已触发，但未发现新的 SunAwtDialog 确认窗口",
        }
    buttons = []
    for dialog in dialogs:
        for button in dialog.get("buttons", []) or []:
            buttons.append(
                {
                    "dialog_hwnd": dialog.get("hwnd"),
                    "path": button.get("path"),
                    "name": button.get("name"),
                    "description": button.get("description"),
                    "actions": button.get("actions") or [],
                    "bounds": button.get("bounds"),
                }
            )
    return {
        "ok": True,
        "reason": "已打开确认窗口；本探针不会点击确认按钮",
        "dialog_count": len(dialogs),
        "button_count": len(buttons),
        "buttons": buttons[:20],
    }


def diagnose_confirm_probe(report, choice):
    confirm = report.get("confirm_action") or {}
    if not confirm.get("ok"):
        return {
            "ok": False,
            "reason": confirm.get("reason") or "确认动作失败",
            "confirm_action": confirm,
        }
    dialogs_left = find_confirm_cancel_dialogs(report.get("dialogs_after_confirm") or [])
    entry_state = report.get("entry_state_after_confirm") or {}
    parent_new_state = report.get("parent_new_state_after_confirm") or {}
    if choice == "no":
        if dialogs_left:
            return {
                "ok": False,
                "reason": "已发送 Alt+N，但确认取消弹窗仍存在",
                "dialogs_left": len(dialogs_left),
                "confirm_action": confirm,
            }
        return {
            "ok": bool(entry_state.get("ok")),
            "reason": (
                "已发送 Alt+N，弹窗关闭且仍在录入态"
                if entry_state.get("ok")
                else "已发送 Alt+N，弹窗关闭，但未检测到录入态按钮"
            ),
            "dialogs_left": len(dialogs_left),
            "entry_state_ok": bool(entry_state.get("ok")),
            "parent_new_state_ok": bool(parent_new_state.get("ok")),
        }
    if dialogs_left:
        return {
            "ok": False,
            "reason": "已发送 Alt+Y，但确认取消弹窗仍存在",
            "dialogs_left": len(dialogs_left),
            "confirm_action": confirm,
        }
    if entry_state.get("ok"):
        return {
            "ok": False,
            "reason": "已发送 Alt+Y，弹窗关闭，但仍检测到保存/暂存/取消录入态按钮",
            "entry_state": entry_state,
        }
    if not parent_new_state.get("ok"):
        return {
            "ok": False,
            "reason": "已发送 Alt+Y，弹窗关闭，录入态按钮消失，但未检测到父页【新增】可用",
            "parent_new_state": parent_new_state,
        }
    return {
        "ok": True,
        "reason": "已发送 Alt+Y，确认取消弹窗关闭，父页【新增】可用且录入态按钮消失",
        "parent_new_state": parent_new_state,
    }


def finish(report, args):
    output_path = logs_dir() / (
        "receipt_cancel_confirm_probe_" + time.strftime("%Y%m%d_%H%M%S") + ".json"
    )
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), "utf-8")
    report["output_path"] = str(output_path)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_human(report)
    return 0 if report.get("ok") else 1


def print_human(report):
    print(f"ok={report.get('ok')} read_only={report.get('read_only')}")
    diagnosis = report.get("diagnosis") or {}
    print(f"reason={diagnosis.get('reason') or report.get('reason')}")
    print(f"cancel_buttons={len(report.get('cancel_buttons') or [])}")
    print(f"dialogs_after={len(report.get('dialogs_after') or [])}")
    print(f"new_or_changed_dialogs={len(report.get('new_or_changed_dialogs') or [])}")
    print(f"output_path={report.get('output_path')}")


if __name__ == "__main__":
    raise SystemExit(main())
