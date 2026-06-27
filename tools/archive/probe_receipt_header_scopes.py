# 生命周期：T0 一次性（删除条件：确认多窗口收款单表头 scope 规则后删除）
# 覆盖的业务阶段：收款单自制页多窗口 JAB scope 只读探测
# 依赖的服务/环境：Windows Python、NC 前台、多窗口/多页签现场、Java Access Bridge
# 运行方式：python tools/archive/probe_receipt_header_scopes.py

import ctypes
import json
import sys
from pathlib import Path
from typing import Any, cast

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from core.jab_probe import JOBJECT, enum_windows  # noqa: E402


HEADER_LABELS = {
    "财务组织": ("财务组织", "财务组织(O)"),
    "客户": ("客户",),
    "单据日期": ("单据日期",),
    "币种": ("币种",),
}


def foreground_info():
    if sys.platform != "win32":
        return {}
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return {}
    root = user32.GetAncestor(hwnd, 2)
    return {
        "hwnd": int(hwnd),
        "root": int(root or 0),
        "class_name": window_class_name(hwnd),
        "title": window_text(hwnd),
        "root_class_name": window_class_name(root),
        "root_title": window_text(root),
    }


def window_text(hwnd):
    if sys.platform != "win32" or not hwnd:
        return ""
    user32 = ctypes.windll.user32
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def window_class_name(hwnd):
    if sys.platform != "win32" or not hwnd:
        return ""
    user32 = ctypes.windll.user32
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def root_hwnd(hwnd):
    if sys.platform != "win32" or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(hwnd, 2) or 0)


def info_dict(info):
    return {
        "name": info.name.strip(),
        "description": info.description.strip(),
        "role": (info.role_en_US.strip() or info.role.strip()),
        "states": (info.states_en_US.strip() or info.states.strip()),
        "x": info.x,
        "y": info.y,
        "width": info.width,
        "height": info.height,
        "valid_bounds": info.x >= 0
        and info.y >= 0
        and info.width > 0
        and info.height > 0,
        "showing": "showing"
        in (info.states_en_US.strip() or info.states.strip()).lower(),
        "visible": "visible"
        in (info.states_en_US.strip() or info.states.strip()).lower(),
    }


def describe_window_header(jab, hwnd, title, class_name, pid, visible):
    vm_id_ref = ctypes.c_long()
    root_context = JOBJECT()
    result = {
        "hwnd": int(hwnd),
        "root_hwnd": root_hwnd(hwnd),
        "title": title,
        "class_name": class_name,
        "pid": pid,
        "window_visible": visible,
        "labels": {},
    }
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd,
        ctypes.byref(vm_id_ref),
        ctypes.byref(root_context),
    ):
        result["error"] = "getAccessibleContextFromHWND failed"
        return result
    vm_id = vm_id_ref.value
    owned = [root_context.value]
    try:
        for canonical_label, aliases in HEADER_LABELS.items():
            result["labels"][canonical_label] = {"ok": False, "aliases": list(aliases)}
            for alias in aliases:
                found = jab.find_text_near_label_by_bounds(
                    vm_id,
                    root_context.value,
                    alias,
                    require_showing=True,
                )
                context, owned_contexts, label_info, text_info = found
                owned.extend(owned_contexts or [])
                if not context:
                    continue
                result["labels"][canonical_label] = {
                    "ok": True,
                    "matched_alias": alias,
                    "aliases": list(aliases),
                    "label": info_dict(label_info),
                    "text": info_dict(text_info),
                }
                break
    finally:
        jab.release_contexts(vm_id, list(dict.fromkeys(owned)))
    ok_labels = [label for label, item in result["labels"].items() if item.get("ok")]
    result["ok_label_count"] = len(ok_labels)
    result["ok_labels"] = ok_labels
    result["complete_header"] = len(ok_labels) == len(HEADER_LABELS)
    result["header_y_values"] = [
        result["labels"][label]["label"]["y"]
        for label in ok_labels
        if result["labels"][label].get("label")
    ]
    return result


def main():
    stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
    if callable(stdout_reconfigure):
        cast(Any, stdout_reconfigure)(encoding="utf-8", errors="replace")
    config = load_config(str(ROOT / "config.json"))
    jab = JABOperator(config)
    jab.ensure_started()
    fg = foreground_info()
    windows = []
    try:
        for hwnd, title, class_name, pid, visible in enum_windows(
            include_children=True
        ):
            if class_name != "SunAwtCanvas" or not visible:
                continue
            if not jab.dll.isJavaWindow(hwnd):
                continue
            item = describe_window_header(jab, hwnd, title, class_name, pid, visible)
            item["is_foreground_root"] = bool(
                fg.get("root") and item["root_hwnd"] == fg.get("root")
            )
            windows.append(item)
    finally:
        jab.close()
    windows.sort(
        key=lambda item: (
            not item.get("is_foreground_root"),
            -int(item.get("ok_label_count") or 0),
            item.get("root_hwnd") or 0,
            item.get("hwnd") or 0,
        )
    )
    complete = [item for item in windows if item.get("complete_header")]
    foreground_complete = [item for item in complete if item.get("is_foreground_root")]
    print(
        json.dumps(
            {
                "foreground": fg,
                "window_count": len(windows),
                "complete_header_count": len(complete),
                "foreground_complete_header_count": len(foreground_complete),
                "recommended_scope_hwnd": (
                    foreground_complete[0]["hwnd"]
                    if len(foreground_complete) == 1
                    else None
                ),
                "windows": windows,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
