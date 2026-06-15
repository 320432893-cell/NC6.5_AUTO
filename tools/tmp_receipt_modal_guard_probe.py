# 职责: 只读探测 NC 模态弹窗特征，并可验证 Alt+C 恢复收款单页面
# 不做什么: 不写表头/明细，不保存/暂存，不枚举业务弹窗白名单
# 允许依赖层: core JAB、tools.jab_probe、收款单表头 scope 定位函数
# 谁不应该 import: 正式流程和 core 模块不应 import 本临时探针
# 生命周期: T0 临时探针（删除条件：模态弹窗守卫正式化并完成异常窗口验证）

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import JOBJECT, enum_windows  # noqa: E402
from tools.receipt_self_made_fill_trial import locate_receipt_header_scope  # noqa: E402
from tools.tmp_receipt_cell_probe_run import send_virtual_key  # noqa: E402


def info_to_dict(info):
    states = info.states_en_US.strip() or info.states.strip()
    return {
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": info.role_en_US.strip() or info.role.strip(),
        "states": states,
        "showing": "showing" in states.lower(),
        "bounds": [info.x, info.y, info.width, info.height],
    }


def collect_visible_java_dialogs(jab):
    dialogs = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not visible or class_name != "SunAwtDialog":
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue
        item = {
            "hwnd": int(hwnd),
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "visible": bool(visible),
            "root_hwnd": root_hwnd(hwnd),
            "cancel_controls": [],
            "buttons": [],
        }
        item.update(scan_dialog_controls(jab, hwnd))
        dialogs.append(item)
    return dialogs


def scan_dialog_controls(jab, hwnd):
    vm_id_ref = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd,
        ctypes.byref(vm_id_ref),
        ctypes.byref(root_context),
    ):
        return {"error": "getAccessibleContextFromHWND failed"}
    buttons = []
    owned = [root_context.value]
    try:
        collect_buttons(jab, vm_id_ref.value, root_context.value, [], buttons, owned, 0)
    finally:
        jab.release_contexts(vm_id_ref.value, list(dict.fromkeys(owned)))
    cancel_controls = [
        item
        for item in buttons
        if "取消" in item.get("name", "") or "Alt+C" in item.get("description", "")
    ]
    return {"buttons": buttons[:20], "cancel_controls": cancel_controls}


def collect_buttons(jab, vm_id, context, path, buttons, owned, depth):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    if role == "push button" and "showing" in states:
        item = info_to_dict(info)
        item["path"] = ".".join(map(str, path))
        buttons.append(item)
    if depth >= min(jab.max_depth, 12):
        return
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        owned.append(child)
        collect_buttons(jab, vm_id, child, path + [index], buttons, owned, depth + 1)


def root_hwnd(hwnd):
    if sys.platform != "win32" or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(wintypes.HWND(int(hwnd)), 2) or 0)


def send_alt_c():
    send_virtual_key(0x12, key_up=False)
    send_virtual_key(0x43, key_up=False)
    send_virtual_key(0x43, key_up=True)
    send_virtual_key(0x12, key_up=True)


def main():
    parser = argparse.ArgumentParser(description="Probe NC modal dialog guard.")
    parser.add_argument("--config", default="config.json")
    parser.add_argument(
        "--recover",
        action="store_true",
        help="send Alt+C once if a visible Java dialog with cancel control exists",
    )
    args = parser.parse_args()

    config = load_config(args.config)
    started_at = time.perf_counter()
    report: dict[str, object] = {"recover_enabled": bool(args.recover)}
    jab = JABOperator(config)
    try:
        jab.ensure_started()
        before_scope = locate_receipt_header_scope(jab)
        before_dialogs = collect_visible_java_dialogs(jab)
        report["before_scope"] = before_scope
        report["before_dialogs"] = before_dialogs
        recoverable = [item for item in before_dialogs if item.get("cancel_controls")]
        report["recoverable_dialog_count"] = len(recoverable)
        if args.recover and recoverable:
            send_alt_c()
            time.sleep(0.5)
            report["recover_action"] = {"ok": True, "method": "Alt+C"}
            report["after_dialogs"] = collect_visible_java_dialogs(jab)
            report["after_scope"] = locate_receipt_header_scope(jab)
        else:
            report["recover_action"] = {
                "ok": False,
                "skipped": True,
                "reason": "recover disabled or no cancelable dialog",
            }
        after_scope = report.get("after_scope")
        after_dialogs = report.get("after_dialogs")
        report["ok"] = bool(
            not args.recover
            or (
                isinstance(after_scope, dict)
                and after_scope.get("ok")
                and not after_dialogs
            )
        )
    finally:
        jab.close()
    report["total_seconds"] = round(time.perf_counter() - started_at, 3)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
