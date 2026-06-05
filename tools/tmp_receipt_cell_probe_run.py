# 生命周期：T0 一次性（删除条件：明细表单元格写入路径确认后删除）
# 覆盖的业务阶段：收款单自制录入-明细主行单元格写入探测
# 依赖的服务/环境：Windows Python、NC 收款单自制录入界面、Java Access Bridge
# 运行方式：python tools/tmp_receipt_cell_probe_run.py

import argparse
import ctypes
from decimal import Decimal, InvalidOperation
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready, print_jab_health_failure  # noqa: E402
from tools.jab_probe import (  # noqa: E402
    AccessibleActions,
    AccessibleTableCellInfo,
    JOBJECT,
)
from tools.receipt_account_reference_try import STOP_HOTKEY, is_stop_hotkey_pressed  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_self_made_fill_trial import read_body_table  # noqa: E402
from tools.receipt_table_cell_probe import collect_text_controls, diff_text_controls  # noqa: E402

START_DELAY_SECONDS = 2
TARGET_ROW = 0
TARGET_COL = 7
TARGET_NAME = "贷方原币金额"
TARGET_VALUE = "1090"


def build_parser():
    parser = argparse.ArgumentParser(
        description="探测 NC 收款单明细金额单元格的 JAB 写入/编辑器入口。"
    )
    parser.add_argument(
        "--activation",
        choices=("f2-enter", "enter", "f2", "none"),
        default="none",
        help="直接 JAB 写入失败后的编辑器激活探测方式。Enter 可能触发增行，默认不发送。",
    )
    parser.add_argument(
        "--key-target",
        choices=("foreground", "table"),
        default="foreground",
        help="受保护按键发送目标：foreground=NC根窗口，table=明细表SunAwtCanvas。",
    )
    parser.add_argument(
        "--write-candidate-amount",
        action="store_true",
        help="试写【组织本币金额】右侧真实 text 控件，然后读回明细表；不保存、不暂存。",
    )
    parser.add_argument(
        "--screen-write-amount",
        action="store_true",
        help="屏幕双击明细金额格后输入1090并读回；不保存、不暂存。",
    )
    parser.add_argument(
        "--screen-commit",
        choices=("tab", "none", "enter"),
        default="none",
        help="屏幕写入后的提交方式。默认 none，不移动单元格；enter 可能触发增行，只能在明确需要时使用。",
    )
    parser.add_argument(
        "--dump-nearby",
        action="store_true",
        help="输出明细表附近全部可见控件。默认关闭，避免测试输出过杂。",
    )
    return parser


def activation_keys(mode):
    return {
        "f2-enter": ("F2", "Enter"),
        "enter": ("Enter",),
        "f2": ("F2",),
        "none": (),
    }[mode]


def normalize_amount(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def amount_matches(actual, expected):
    actual_amount = normalize_amount(actual)
    expected_amount = normalize_amount(expected)
    return (
        actual_amount is not None
        and expected_amount is not None
        and actual_amount == expected_amount
    )


def current_amount_text(table_text):
    for key in ("after_screen", "after_candidate", "after_write"):
        value = table_text.get(key)
        if value is not None:
            return value
    return None


def print_header(
    activation_mode,
    key_target,
    write_candidate_amount,
    screen_write_amount,
    screen_commit,
    dump_nearby,
):
    print("测试功能：收款单明细表单字段 JAB 后台写入探测")
    print()
    print("测试数据来源：")
    print("1. 试验字段：第 1 行【贷方原币金额】")
    print("2. 目标值：1090（来自当前 row 1424 到账金额试验值）")
    print()
    print("前置条件：")
    print("1. NC 已停在收款单自制录入界面")
    print("2. 当前没有打开参照窗口、提示框、下拉框")
    print("3. 明细表可见，第一行仍是要测试的主行")
    print()
    print("本脚本会做：")
    print("1. 检查 JAB 健康状态")
    print("2. 定位 25 列收款明细表")
    print("3. 选中第 1 行【贷方原币金额】单元格")
    print("4. 读取该单元格的 JAB 接口、动作、文本状态")
    print("5. 尝试一次 JAB 后台 setTextContents 写入 1090")
    print(
        "6. 如果直接写失败，且 NC 是前台窗口，"
        f"则按 {activation_keys(activation_mode)} 探测编辑器"
    )
    print("7. 读取写入/激活前后的明细表和新增/变化的文本控件")
    print()
    if screen_write_amount:
        print("屏幕试写会做：受保护坐标点击当前金额格、输入目标金额、读回校验。")
    print("不会做：保存、暂存、关闭收款单、剪贴板输入")
    if screen_commit == "none":
        print("不会做：按 Enter 增行、按 Tab 移到下一格")
    elif screen_commit != "enter":
        print("不会做：按 Enter 增行")
    print(
        "注意：F2/Enter 只在确认当前前台窗口属于同一个 NC 窗口后发送；不匹配则不发送。"
    )
    print(f"本轮激活探测：{activation_mode}")
    print(f"本轮按键目标：{key_target}")
    print(f"本轮候选金额控件试写：{write_candidate_amount}")
    print(f"本轮屏幕金额格试写：{screen_write_amount}")
    print(f"屏幕写入提交键：{screen_commit}")
    if screen_commit == "none":
        print("说明：默认不按 Enter/Tab，避免增行或移动到下一个单元格。")
    if screen_commit == "enter":
        print("警告：Enter 可能触发增行；只有明确业务需要新增行时才使用。")
    print(f"输出附近控件明细：{dump_nearby}")
    print(f"紧急停止：按 {STOP_HOTKEY}")
    print(f"启动后等待：{START_DELAY_SECONDS} 秒，用来切到 NC 收款单窗口")
    print("=" * 60)


def wait_exit():
    try:
        input("按回车退出...")
    except (KeyboardInterrupt, EOFError):
        print()
        print("已退出。")


def print_key_table(title, snapshot):
    print(title)
    if not snapshot.get("ok"):
        print(f"  失败：{snapshot.get('reason')}")
        return
    print(f"  明细表：{snapshot.get('row_count')} 行 x {snapshot.get('col_count')} 列")
    rows = snapshot.get("rows") or []
    if rows:
        cells = rows[0].get("cells") or {}
        print(
            "  第 1 行关键列："
            f"业务类型={cells.get('1')!r}, "
            f"币种={cells.get('3')!r}, "
            f"账户={cells.get('4')!r}, "
            f"科目={cells.get('5')!r}, "
            f"金额={cells.get('7')!r}, "
            f"结算={cells.get('11')!r}"
        )


def safe_repr(value):
    return ascii(value)


def get_action_names(jab, vm_id, context):
    if not hasattr(jab.dll, "getAccessibleActions"):
        return []
    actions = AccessibleActions()
    if not jab.dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
        return []
    return [
        actions.actionInfo[index].name.strip() for index in range(actions.actionsCount)
    ]


def describe_context(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return {"ok": False, "reason": "getAccessibleContextInfo failed"}
    return {
        "ok": True,
        "role": info.role_en_US.strip() or info.role.strip(),
        "name": info.name.strip(),
        "description": info.description.strip(),
        "states": info.states_en_US.strip() or info.states.strip(),
        "bounds": [info.x, info.y, info.width, info.height],
        "children": info.childrenCount,
        "accessibleText": bool(info.accessibleText),
        "accessibleAction": bool(info.accessibleAction),
        "accessibleSelection": bool(info.accessibleSelection),
        "accessibleInterfaces": bool(info.accessibleInterfaces),
        "textValue": jab.get_text_context_value(vm_id, context),
        "actions": get_action_names(jab, vm_id, context),
    }


def describe_nearby_controls(jab, table_window, table_bounds):
    if not table_window or not table_bounds:
        return []
    hwnd = int(table_window.get("hwnd") or 0)
    if not hwnd:
        return []
    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd,
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return []
    controls = []
    try:
        collect_nearby_controls_in_tree(
            jab,
            vm_id.value,
            root_context.value,
            path=[],
            table_bounds=table_bounds,
            controls=controls,
            depth=0,
        )
    finally:
        jab.release_contexts(vm_id.value, [root_context.value])
    return controls[:80]


def write_candidate_amount_control(jab, table_window, table_bounds, value):
    candidate = find_candidate_amount_control(
        jab,
        table_window,
        table_bounds,
        label_name="组织本币金额",
    )
    if not candidate.get("ok"):
        return candidate

    result = jab.find_context_by_path_once(
        candidate["text_path"],
        class_name=table_window.get("class_name"),
        require_showing=True,
        require_valid_bounds=False,
    )
    context, vm_id, owned, _window_info = result
    if not context:
        return {
            "ok": False,
            "reason": "按候选 path 重新取得 text context 失败",
            "candidate": candidate,
        }
    try:
        before = jab.get_text_context_value(vm_id, context)
        before_info = jab.get_context_info(vm_id, context)
        write_ok = jab.set_text_context(vm_id, context, value)
        time.sleep(0.8)
        after = jab.get_text_context_value(vm_id, context)
        after_info = jab.get_context_info(vm_id, context)
        return {
            "ok": bool(write_ok),
            "candidate": candidate,
            "text_before": before,
            "description_before": before_info.description.strip()
            if before_info
            else None,
            "text_after": after,
            "description_after": after_info.description.strip() if after_info else None,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def find_candidate_amount_control(jab, table_window, table_bounds, label_name):
    if not table_window or not table_bounds:
        return {"ok": False, "reason": "缺少明细表窗口或 bounds"}
    hwnd = int(table_window.get("hwnd") or 0)
    if not hwnd:
        return {"ok": False, "reason": "缺少明细表 hwnd"}
    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd,
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return {"ok": False, "reason": "无法取得明细表窗口 JAB root context"}
    all_controls = []
    try:
        collect_candidate_controls_in_tree(
            jab,
            vm_id.value,
            root_context.value,
            path=[],
            table_bounds=table_bounds,
            controls=all_controls,
            depth=0,
        )
    finally:
        jab.release_contexts(vm_id.value, [root_context.value])

    labels = [
        item
        for item in all_controls
        if item["role"].lower() == "label" and item["name"] == label_name
    ]
    if not labels:
        return {"ok": False, "reason": f"未找到 label：{label_name}"}
    label = sorted(labels, key=lambda item: item["bounds"][1])[0]
    lx, ly, lw, lh = label["bounds"]
    text_candidates = []
    for item in all_controls:
        if item["role"].lower() != "text":
            continue
        x, y, _width, _height = item["bounds"]
        same_row = abs(y - ly) <= 6
        right_side = x > lx + lw
        if same_row and right_side:
            text_candidates.append(item)
    if not text_candidates:
        return {
            "ok": False,
            "reason": f"找到 label 但右侧没有 text 控件：{label_name}",
            "label": label,
        }
    text = sorted(text_candidates, key=lambda item: item["bounds"][0])[0]
    return {
        "ok": True,
        "label": label,
        "text": text,
        "text_path": text["path"],
    }


def collect_candidate_controls_in_tree(
    jab, vm_id, context, path, table_bounds, controls, depth
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    role = info.role_en_US.strip() or info.role.strip()
    states = info.states_en_US.strip() or info.states.strip()
    bounds = [info.x, info.y, info.width, info.height]
    if control_is_near_table(role.lower(), states.lower(), bounds, table_bounds):
        controls.append(
            {
                "path": "0" + "".join(f".{index}" for index in path),
                "role": role,
                "name": info.name.strip(),
                "description": info.description.strip(),
                "states": states,
                "bounds": bounds,
            }
        )
    if depth >= jab.max_depth:
        return
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        try:
            collect_candidate_controls_in_tree(
                jab,
                vm_id,
                child,
                path + [index],
                table_bounds,
                controls,
                depth + 1,
            )
        finally:
            jab.release_contexts(vm_id, [child])


def collect_nearby_controls_in_tree(
    jab, vm_id, context, path, table_bounds, controls, depth
):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return
    role = info.role_en_US.strip() or info.role.strip()
    role_l = role.lower()
    states = info.states_en_US.strip() or info.states.strip()
    states_l = states.lower()
    bounds = [info.x, info.y, info.width, info.height]
    path_text = "0" + "".join(f".{index}" for index in path)
    if control_is_near_table(role_l, states_l, bounds, table_bounds):
        controls.append(
            {
                "path": path_text,
                "role": role,
                "name": info.name.strip(),
                "description": info.description.strip(),
                "states": states,
                "bounds": bounds,
                "accessibleText": bool(info.accessibleText),
                "accessibleAction": bool(info.accessibleAction),
                "accessibleSelection": bool(info.accessibleSelection),
                "textValue": jab.get_text_context_value(vm_id, context),
                "actions": get_action_names(jab, vm_id, context),
            }
        )
    if depth >= jab.max_depth:
        return
    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        try:
            collect_nearby_controls_in_tree(
                jab,
                vm_id,
                child,
                path + [index],
                table_bounds,
                controls,
                depth + 1,
            )
        finally:
            jab.release_contexts(vm_id, [child])


def control_is_near_table(role, states, bounds, table_bounds):
    if "visible" not in states or "showing" not in states:
        return False
    x, y, width, height = bounds
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        return False
    table_x, table_y, table_width, table_height = table_bounds
    vertical_ok = (table_y - 160) <= y <= (table_y + table_height + 160)
    horizontal_ok = x + width >= table_x and x <= table_x + table_width
    interesting_role = role in {
        "text",
        "combo box",
        "push button",
        "label",
        "table",
        "panel",
        "viewport",
    }
    return vertical_ok and horizontal_ok and interesting_role


def read_window_info(hwnd):
    if sys.platform != "win32" or not hwnd:
        return None
    user32 = ctypes.windll.user32
    length = user32.GetWindowTextLengthW(hwnd)
    title = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title, length + 1)
    class_name = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_name, 256)
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    root = user32.GetAncestor(hwnd, 2)
    return {
        "hwnd": int(hwnd),
        "title": title.value,
        "class_name": class_name.value,
        "pid": int(pid.value),
        "visible": bool(user32.IsWindowVisible(hwnd)),
        "root_hwnd": int(root) if root else None,
    }


KEYS = {
    "F2": 0x71,
    "Enter": 0x0D,
    "Delete": 0x2E,
    "Ctrl": 0x11,
    "D": 0x44,
    "I": 0x49,
}


def screen_write_amount_cell(
    table_window,
    table_bounds,
    row,
    col,
    col_count,
    value,
    commit_key,
    row_count=1,
):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    if not table_bounds or len(table_bounds) != 4:
        return {"ok": False, "reason": "缺少有效表格 bounds"}
    x, y, width, height = table_bounds
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        return {"ok": False, "reason": f"表格 bounds 不可见：{table_bounds}"}
    if col_count <= 0:
        return {"ok": False, "reason": f"列数无效：{col_count}"}
    table_info = read_window_info((table_window or {}).get("hwnd"))
    user32 = ctypes.windll.user32
    foreground_info = read_window_info(user32.GetForegroundWindow())
    if not table_info or not foreground_info:
        return {
            "ok": False,
            "reason": "无法读取当前前台窗口或明细表窗口",
            "table_window": table_info,
            "foreground": foreground_info,
        }
    same_root = (
        foreground_info.get("hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("root_hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("hwnd") == table_info.get("hwnd")
    )
    if not same_root:
        return {
            "ok": False,
            "reason": "当前前台窗口不是本次定位到的 NC 收款单窗口，未执行屏幕输入",
            "table_window": table_info,
            "foreground": foreground_info,
        }
    target_x, target_y, cell_width, cell_height = cell_center(
        table_bounds, row, col, row_count, col_count
    )
    try:
        move_mouse(target_x, target_y)
        mouse_click()
        time.sleep(0.08)
        mouse_click()
        time.sleep(0.2)
        send_hotkey_ctrl_a()
        time.sleep(0.1)
        send_text(value)
        time.sleep(0.1)
        if commit_key == "tab":
            send_virtual_key(0x09)
        elif commit_key == "enter":
            send_virtual_key(0x0D)
    except Exception as exc:
        return {
            "ok": False,
            "reason": f"屏幕输入失败：{type(exc).__name__}: {exc}",
            "target": [target_x, target_y],
            "table_bounds": table_bounds,
            "cell_width": cell_width,
            "cell_height": cell_height,
            "commit_key": commit_key,
            "foreground": foreground_info,
            "table_window": table_info,
        }
    time.sleep(0.8)
    return {
        "ok": True,
        "target": [target_x, target_y],
        "table_bounds": table_bounds,
        "cell_width": cell_width,
        "cell_height": cell_height,
        "commit_key": commit_key,
        "foreground": foreground_info,
        "table_window": table_info,
    }


def cell_center(table_bounds, row, col, row_count, col_count):
    x, y, width, height = table_bounds
    safe_row_count = max(int(row_count or 1), int(row) + 1, 1)
    safe_col_count = max(int(col_count or 1), int(col) + 1, 1)
    cell_width = width / safe_col_count
    cell_height = height / safe_row_count
    target_x = int(x + cell_width * col + cell_width / 2)
    target_y = int(y + cell_height * row + cell_height / 2)
    return target_x, target_y, cell_width, cell_height


def move_mouse(x, y):
    ctypes.windll.user32.SetCursorPos(int(x), int(y))


def mouse_click():
    send_mouse_event(0x0002)
    send_mouse_event(0x0004)


def send_mouse_event(flags):
    inp = INPUT()
    inp.type = 0
    inp.mi = MOUSEINPUT(0, 0, 0, flags, 0, None)
    send_input(inp)


def send_hotkey_ctrl_a():
    send_virtual_key(0x11, key_up=False)
    send_virtual_key(0x41, key_up=False)
    send_virtual_key(0x41, key_up=True)
    send_virtual_key(0x11, key_up=True)


def send_hotkey_ctrl_i():
    send_virtual_key(0x11, key_up=False)
    send_virtual_key(0x49, key_up=False)
    send_virtual_key(0x49, key_up=True)
    send_virtual_key(0x11, key_up=True)


def send_hotkey_ctrl_d():
    send_virtual_key(0x11, key_up=False)
    send_virtual_key(0x44, key_up=False)
    send_virtual_key(0x44, key_up=True)
    send_virtual_key(0x11, key_up=True)


def send_text(text):
    for char in str(text):
        send_unicode_char(char)


def send_unicode_char(char):
    code = ord(char)
    inp = INPUT()
    inp.type = 1
    inp.ki = KEYBDINPUT(0, code, 0x0004, 0, None)
    send_input(inp)
    inp_up = INPUT()
    inp_up.type = 1
    inp_up.ki = KEYBDINPUT(0, code, 0x0004 | 0x0002, 0, None)
    send_input(inp_up)


def send_virtual_key(vk, key_up=False):
    inp = INPUT()
    inp.type = 1
    inp.ki = KEYBDINPUT(vk, 0, 0x0002 if key_up else 0, 0, None)
    send_input(inp)


def send_input(inp):
    ctypes.windll.kernel32.SetLastError(0)
    sent = ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    if sent != 1:
        error_code = ctypes.windll.kernel32.GetLastError()
        raise RuntimeError(f"SendInput failed, error={error_code}")


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class INPUT(ctypes.Structure):
    class _INPUT_UNION(ctypes.Union):
        _fields_ = [
            ("mi", MOUSEINPUT),
            ("ki", KEYBDINPUT),
        ]

    _anonymous_ = ("union",)
    _fields_ = [
        ("type", ctypes.c_ulong),
        ("union", _INPUT_UNION),
    ]


def guarded_press_key(table_window, key_name, key_target):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    if key_name not in KEYS:
        return {"ok": False, "reason": f"未知按键：{key_name}"}
    user32 = ctypes.windll.user32
    table_hwnd = int((table_window or {}).get("hwnd") or 0)
    table_info = read_window_info(table_hwnd)
    foreground_hwnd = user32.GetForegroundWindow()
    foreground_info = read_window_info(foreground_hwnd)
    if not table_info or not foreground_info:
        return {
            "ok": False,
            "reason": "无法读取当前前台窗口或明细表窗口",
            "table_window": table_info,
            "foreground": foreground_info,
        }

    same_root = (
        foreground_info.get("hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("root_hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("hwnd") == table_info.get("hwnd")
    )
    if not same_root:
        return {
            "ok": False,
            "reason": f"当前前台窗口不是本次定位到的 NC 收款单窗口，未发送 {key_name}",
            "table_window": table_info,
            "foreground": foreground_info,
        }

    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    if key_target == "table":
        target_hwnd = table_info["hwnd"]
    else:
        target_hwnd = foreground_info["hwnd"]
    vk = KEYS[key_name]
    down_ok = bool(user32.PostMessageW(target_hwnd, WM_KEYDOWN, vk, 0))
    up_ok = bool(user32.PostMessageW(target_hwnd, WM_KEYUP, vk, 0))
    time.sleep(0.6)
    return {
        "ok": bool(down_ok and up_ok),
        "mode": f"PostMessage({key_name})",
        "key": key_name,
        "key_target": key_target,
        "target_hwnd": target_hwnd,
        "table_window": table_info,
        "foreground": foreground_info,
        "down_ok": down_ok,
        "up_ok": up_ok,
    }


def guarded_send_ctrl_i(table_window):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    user32 = ctypes.windll.user32
    table_hwnd = int((table_window or {}).get("hwnd") or 0)
    table_info = read_window_info(table_hwnd)
    foreground_hwnd = user32.GetForegroundWindow()
    foreground_info = read_window_info(foreground_hwnd)
    if not table_info or not foreground_info:
        return {
            "ok": False,
            "reason": "无法读取当前前台窗口或明细表窗口",
            "table_window": table_info,
            "foreground": foreground_info,
        }

    same_root = (
        foreground_info.get("hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("root_hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("hwnd") == table_info.get("hwnd")
    )
    if not same_root:
        return {
            "ok": False,
            "reason": "当前前台窗口不是本次定位到的 NC 收款单窗口，未发送 Ctrl+I",
            "table_window": table_info,
            "foreground": foreground_info,
        }

    try:
        send_hotkey_ctrl_i()
    except Exception as exc:
        return {
            "ok": False,
            "mode": "SendInput(Ctrl+I)",
            "key": "Ctrl+I",
            "table_window": table_info,
            "foreground": foreground_info,
            "reason": f"{type(exc).__name__}: {exc}",
        }
    time.sleep(0.8)
    return {
        "ok": True,
        "mode": "SendInput(Ctrl+I)",
        "key": "Ctrl+I",
        "table_window": table_info,
        "foreground": foreground_info,
    }


def guarded_send_ctrl_d(table_window):
    return guarded_send_table_hotkey(table_window, "Ctrl+D", send_hotkey_ctrl_d)


def guarded_send_delete(table_window):
    return guarded_send_table_hotkey(
        table_window, "Delete", lambda: send_virtual_key(0x2E)
    )


def guarded_send_table_hotkey(table_window, key_name, sender):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    user32 = ctypes.windll.user32
    table_hwnd = int((table_window or {}).get("hwnd") or 0)
    table_info = read_window_info(table_hwnd)
    foreground_hwnd = user32.GetForegroundWindow()
    foreground_info = read_window_info(foreground_hwnd)
    if not table_info or not foreground_info:
        return {
            "ok": False,
            "reason": "无法读取当前前台窗口或明细表窗口",
            "table_window": table_info,
            "foreground": foreground_info,
        }

    same_root = (
        foreground_info.get("hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("root_hwnd") == table_info.get("root_hwnd")
        or foreground_info.get("hwnd") == table_info.get("hwnd")
    )
    if not same_root:
        return {
            "ok": False,
            "reason": f"当前前台窗口不是本次定位到的 NC 收款单窗口，未发送 {key_name}",
            "table_window": table_info,
            "foreground": foreground_info,
        }

    try:
        sender()
    except Exception as exc:
        return {
            "ok": False,
            "mode": f"SendInput({key_name})",
            "key": key_name,
            "table_window": table_info,
            "foreground": foreground_info,
            "reason": f"{type(exc).__name__}: {exc}",
        }
    time.sleep(0.8)
    return {
        "ok": True,
        "mode": f"SendInput({key_name})",
        "key": key_name,
        "table_window": table_info,
        "foreground": foreground_info,
    }


def get_cell_context(jab, vm_id, table_context, row, col):
    if not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        return None, {
            "ok": False,
            "reason": "JAB DLL 不支持 getAccessibleTableCellInfo",
        }
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    report = {
        "ok": bool(ok),
        "index": cell_info.index,
        "row": cell_info.row,
        "col": cell_info.column,
        "rowExtent": cell_info.rowExtent,
        "columnExtent": cell_info.columnExtent,
        "isSelected": bool(cell_info.isSelected),
        "hasContext": bool(cell_info.accessibleContext),
    }
    if not ok or not cell_info.accessibleContext:
        return None, report
    return cell_info.accessibleContext, report


def probe_one_cell(
    jab,
    activation_mode,
    key_target,
    write_candidate_amount,
    screen_write_amount,
    screen_commit,
    dump_nearby,
):
    located = locate_receipt_body_table(jab, max_rows=3)
    best = located.get("best")
    if not best:
        return {
            "ok": False,
            "failed_step": "locate-body-table",
            "candidates": (located.get("candidates") or [])[:5],
        }

    context, vm_id, owned, window_info = jab.find_context_by_path_once(
        best["path"],
        class_name=best["window"].get("class_name"),
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "failed_step": "reopen-body-table",
            "reason": "定位到候选表，但按 path 重新取得 table context 失败",
            "best": best,
        }

    cell_context = None
    before_text_controls = []
    after_text_controls = []
    after_f2_text_controls = []
    after_enter_text_controls = []
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            return {
                "ok": False,
                "failed_step": "read-table-info",
                "reason": "getAccessibleTableInfo 失败",
                "best": best,
                "window": window_info,
            }
        if TARGET_ROW >= table_info.rowCount or TARGET_COL >= table_info.columnCount:
            return {
                "ok": False,
                "failed_step": "target-out-of-range",
                "reason": (
                    f"目标 row={TARGET_ROW} col={TARGET_COL} 超出 "
                    f"{table_info.rowCount}x{table_info.columnCount}"
                ),
            }

        child_index = TARGET_ROW * table_info.columnCount + TARGET_COL
        before_table_text = jab.get_table_cell_text(
            vm_id, context, TARGET_ROW, TARGET_COL
        )
        selected_before = jab.get_selected_child_indexes(
            vm_id, context, table_info.rowCount * table_info.columnCount
        )
        jab.dll.clearAccessibleSelectionFromContext(vm_id, context)
        jab.dll.addAccessibleSelectionFromContext(vm_id, context, child_index)
        time.sleep(0.2)
        selected_after = jab.get_selected_child_indexes(
            vm_id, context, table_info.rowCount * table_info.columnCount
        )
        after_select_table_text = jab.get_table_cell_text(
            vm_id, context, TARGET_ROW, TARGET_COL
        )
        cell_context, cell_info = get_cell_context(
            jab, vm_id, context, TARGET_ROW, TARGET_COL
        )
        table_context_info = describe_context(jab, vm_id, context)
        nearby_controls = []
        if dump_nearby:
            nearby_controls = describe_nearby_controls(
                jab,
                window_info,
                table_context_info.get("bounds"),
            )
        cell_before = (
            describe_context(jab, vm_id, cell_context) if cell_context else None
        )

        table_focus_ok = None
        cell_focus_ok = None
        if hasattr(jab.dll, "requestFocus"):
            table_focus_ok = bool(jab.dll.requestFocus(vm_id, context))
            if cell_context:
                cell_focus_ok = bool(jab.dll.requestFocus(vm_id, cell_context))
        time.sleep(0.2)

        before_text_controls = collect_text_controls(jab, window_class="SunAwtCanvas")

        direct_write = {"ok": False, "reason": "cell context unavailable"}
        if cell_context:
            text_before = jab.get_text_context_value(vm_id, cell_context)
            write_ok = jab.set_text_context(vm_id, cell_context, TARGET_VALUE)
            time.sleep(0.5)
            direct_write = {
                "ok": bool(write_ok),
                "text_before": text_before,
                "text_after": jab.get_text_context_value(vm_id, cell_context),
            }

        after_text_controls = collect_text_controls(jab, window_class="SunAwtCanvas")
        final_table_text = jab.get_table_cell_text(
            vm_id, context, TARGET_ROW, TARGET_COL
        )
        cell_after = (
            describe_context(jab, vm_id, cell_context) if cell_context else None
        )
        editor_probe = None
        enter_probe = None
        candidate_amount_write = None
        screen_amount_write = None
        after_candidate_table_text = None
        after_screen_table_text = None
        after_f2_table_text = None
        after_enter_table_text = None
        if (
            not amount_matches(final_table_text, TARGET_VALUE)
            and write_candidate_amount
        ):
            candidate_amount_write = write_candidate_amount_control(
                jab,
                window_info,
                table_context_info.get("bounds"),
                TARGET_VALUE,
            )
            after_candidate_table_text = jab.get_table_cell_text(
                vm_id, context, TARGET_ROW, TARGET_COL
            )
        current_table_text = after_candidate_table_text or final_table_text
        if not amount_matches(current_table_text, TARGET_VALUE) and screen_write_amount:
            screen_amount_write = screen_write_amount_cell(
                window_info,
                table_context_info.get("bounds"),
                TARGET_ROW,
                TARGET_COL,
                table_info.columnCount,
                TARGET_VALUE,
                screen_commit,
                table_info.rowCount,
            )
            after_screen_table_text = jab.get_table_cell_text(
                vm_id, context, TARGET_ROW, TARGET_COL
            )
        success_text = current_amount_text(
            {
                "after_write": final_table_text,
                "after_candidate": after_candidate_table_text,
                "after_screen": after_screen_table_text,
            }
        )
        keys_to_probe = activation_keys(activation_mode)
        if not amount_matches(success_text, TARGET_VALUE) and "F2" in keys_to_probe:
            editor_probe = guarded_press_key(window_info, "F2", key_target)
            after_f2_text_controls = collect_text_controls(
                jab, window_class="SunAwtCanvas"
            )
            after_f2_table_text = jab.get_table_cell_text(
                vm_id, context, TARGET_ROW, TARGET_COL
            )
            success_text = after_f2_table_text or success_text
        if not amount_matches(success_text, TARGET_VALUE) and "Enter" in keys_to_probe:
            enter_probe = guarded_press_key(window_info, "Enter", key_target)
            after_enter_text_controls = collect_text_controls(
                jab, window_class="SunAwtCanvas"
            )
            after_enter_table_text = jab.get_table_cell_text(
                vm_id, context, TARGET_ROW, TARGET_COL
            )
            success_text = after_enter_table_text or success_text
        success = amount_matches(success_text, TARGET_VALUE)

        return {
            "ok": bool(success),
            "failed_step": None if success else "cell-write",
            "table": {
                "path": best["path"],
                "window": window_info,
                "row_count": table_info.rowCount,
                "col_count": table_info.columnCount,
            },
            "target": {
                "row": TARGET_ROW,
                "col": TARGET_COL,
                "name": TARGET_NAME,
                "child_index": child_index,
                "value": TARGET_VALUE,
            },
            "selection": {
                "selected_before": selected_before,
                "selected_after": selected_after,
                "ok": child_index in selected_after,
            },
            "table_text": {
                "before": before_table_text,
                "after_select": after_select_table_text,
                "after_write": final_table_text,
                "after_candidate": after_candidate_table_text,
                "after_screen": after_screen_table_text,
                "after_f2": after_f2_table_text,
                "after_enter": after_enter_table_text,
                "success_text": success_text,
            },
            "table_context": table_context_info,
            "nearby_controls": nearby_controls,
            "cell_info": cell_info,
            "cell_before": cell_before,
            "focus": {
                "table_requestFocus": table_focus_ok,
                "cell_requestFocus": cell_focus_ok,
            },
            "direct_write": direct_write,
            "candidate_amount_write": candidate_amount_write,
            "screen_amount_write": screen_amount_write,
            "cell_after": cell_after,
            "changed_text_controls": diff_text_controls(
                before_text_controls, after_text_controls
            )[:20],
            "editor_probe": editor_probe,
            "changed_text_controls_after_f2": diff_text_controls(
                before_text_controls, after_f2_text_controls
            )[:30],
            "enter_probe": enter_probe,
            "changed_text_controls_after_enter": diff_text_controls(
                before_text_controls, after_enter_text_controls
            )[:30],
        }
    finally:
        if cell_context:
            jab.release_contexts(vm_id, [cell_context])
        jab.release_contexts(vm_id, owned)


def print_candidates(candidates):
    if not candidates:
        print("  未发现任何 JAB 表格候选。")
        return
    print("  候选表摘要：")
    for item in candidates[:5]:
        print(
            "  - "
            f"{item.get('row_count')} 行 x {item.get('col_count')} 列，"
            f"score={item.get('score')}，"
            f"窗口={((item.get('window') or {}).get('title') or '<无标题>')}"
        )
        print(f"    原因：{item.get('reasons')}")
        rows = item.get("rows") or []
        if rows:
            print(f"    第 1 行关键列：{(rows[0].get('cells') or {})}")


def print_context_block(title, context):
    print(title)
    if not context:
        print("  无单元格 context。")
        return
    if not context.get("ok"):
        print(f"  失败：{context.get('reason')}")
        return
    print(f"  role={safe_repr(context.get('role'))}")
    print(f"  name={safe_repr(context.get('name'))}")
    print(f"  description={safe_repr(context.get('description'))}")
    print(f"  states={safe_repr(context.get('states'))}")
    print(f"  bounds={context.get('bounds')}")
    print(
        "  接口："
        f"text={context.get('accessibleText')} "
        f"action={context.get('accessibleAction')} "
        f"selection={context.get('accessibleSelection')} "
        f"interfaces={context.get('accessibleInterfaces')}"
    )
    print(f"  textValue={safe_repr(context.get('textValue'))}")
    print(f"  actions={safe_repr(context.get('actions'))}")


def print_text_control_changes(title, changed):
    print(title)
    if not changed:
        print("  无。")
        return
    for item in changed[:30]:
        window = item.get("window") or {}
        print(
            "  - "
            f"path={item.get('path')} role={safe_repr(item.get('role'))} "
            f"name={safe_repr(item.get('name'))} "
            f"desc={safe_repr(item.get('description'))} "
            f"value={safe_repr(item.get('value'))} showing={item.get('showing')} "
            f"bounds={item.get('bounds')} "
            f"window={safe_repr(window.get('title') or '<无标题>')}/"
            f"{safe_repr(window.get('class'))}"
        )


def print_nearby_controls(controls):
    print("明细表附近可见控件：")
    if not controls:
        print("  未输出。需要时加 --dump-nearby。")
        return
    for item in controls[:80]:
        print(
            "  - "
            f"path={item.get('path')} role={safe_repr(item.get('role'))} "
            f"name={safe_repr(item.get('name'))} "
            f"desc={safe_repr(item.get('description'))} "
            f"states={safe_repr(item.get('states'))} "
            f"bounds={item.get('bounds')} "
            f"text={item.get('accessibleText')} "
            f"action={item.get('accessibleAction')} "
            f"selection={item.get('accessibleSelection')} "
            f"value={safe_repr(item.get('textValue'))} "
            f"actions={safe_repr(item.get('actions'))}"
        )


def print_summary(report):
    print()
    print("测试结果：")
    if report.get("stopped_by_hotkey"):
        print(f"已停止：检测到紧急停止键 {STOP_HOTKEY}。")
        print(f"停止位置：{report.get('failed_step')}")
        return

    if report.get("exception"):
        print(f"脚本异常：{report.get('exception')}")
        print(f"原因：{report.get('reason')}")
        return

    if report.get("failed_step") == "jab-health-check":
        health = report.get("jab_health") or {}
        if isinstance(health, dict):
            print_jab_health_failure(health)
        print("失败：当前不能读取 NC JAB 控件树，未执行单元格探测。")
        return

    if report.get("failed_step") == "locate-body-table":
        print("明细表定位：失败。")
        print_candidates(report.get("candidates") or [])
        print("本次没有尝试写入。")
        return

    if report.get("failed_step") in (
        "reopen-body-table",
        "read-table-info",
        "target-out-of-range",
    ):
        print("明细表准备：失败。")
        print(f"原因：{report.get('reason')}")
        return

    print_key_table("写入前明细表：", report.get("before_table") or {})
    table = report.get("table") or {}
    target = report.get("target") or {}
    print("定位结果：")
    print(
        f"  明细表：{table.get('row_count')} 行 x {table.get('col_count')} 列，"
        f"path={table.get('path')}"
    )
    print(
        f"  目标单元格：第 {target.get('row', 0) + 1} 行，"
        f"第 {target.get('col', 0) + 1} 列，字段={target.get('name')}，"
        f"childIndex={target.get('child_index')}"
    )

    selection = report.get("selection") or {}
    print("选中结果：")
    print(f"  选中前：{selection.get('selected_before')}")
    print(f"  选中后：{selection.get('selected_after')}")
    print(f"  目标是否被选中：{selection.get('ok')}")

    print("表格读值：")
    table_text = report.get("table_text") or {}
    print(f"  写入前：{table_text.get('before')!r}")
    print(f"  选中后：{table_text.get('after_select')!r}")
    print(f"  写入后：{table_text.get('after_write')!r}")
    print(f"  候选控件写后：{table_text.get('after_candidate')!r}")
    print(f"  屏幕写后：{table_text.get('after_screen')!r}")
    print(f"  F2 后：{table_text.get('after_f2')!r}")
    print(f"  Enter 后：{table_text.get('after_enter')!r}")
    print(f"  本轮判定值：{table_text.get('success_text')!r}")

    print_context_block("明细表父 context：", report.get("table_context"))
    print_nearby_controls(report.get("nearby_controls") or [])
    print("单元格基础信息：")
    cell_info = report.get("cell_info") or {}
    print(
        "  "
        f"getCellInfo={cell_info.get('ok')} "
        f"hasContext={cell_info.get('hasContext')} "
        f"index={cell_info.get('index')} "
        f"isSelected={cell_info.get('isSelected')}"
    )
    print_context_block("写入前单元格 context：", report.get("cell_before"))

    focus = report.get("focus") or {}
    print("requestFocus：")
    print(f"  table={focus.get('table_requestFocus')}")
    print(f"  cell={focus.get('cell_requestFocus')}")

    direct = report.get("direct_write") or {}
    print("JAB 后台写入：")
    print(f"  setTextContents 结果：{direct.get('ok')}")
    if direct.get("reason"):
        print(f"  原因：{direct.get('reason')}")
    print(f"  context text 写入前：{direct.get('text_before')!r}")
    print(f"  context text 写入后：{direct.get('text_after')!r}")
    print_context_block("写入后单元格 context：", report.get("cell_after"))

    candidate_write = report.get("candidate_amount_write")
    print("候选金额控件试写：")
    if not candidate_write:
        print("  未执行。")
    else:
        print(f"  写入结果：{candidate_write.get('ok')}")
        if candidate_write.get("reason"):
            print(f"  原因：{candidate_write.get('reason')}")
        candidate = candidate_write.get("candidate") or {}
        label = candidate.get("label") or {}
        text = candidate.get("text") or {}
        print(
            "  label："
            f"path={label.get('path')} name={safe_repr(label.get('name'))} "
            f"bounds={label.get('bounds')}"
        )
        print(
            "  text："
            f"path={candidate.get('text_path')} "
            f"desc={safe_repr(text.get('description'))} "
            f"bounds={text.get('bounds')}"
        )
        print(
            "  控件值："
            f"text_before={safe_repr(candidate_write.get('text_before'))} "
            f"desc_before={safe_repr(candidate_write.get('description_before'))} "
            f"text_after={safe_repr(candidate_write.get('text_after'))} "
            f"desc_after={safe_repr(candidate_write.get('description_after'))}"
        )

    screen_write = report.get("screen_amount_write")
    print("屏幕金额格试写：")
    if not screen_write:
        print("  未执行。")
    else:
        print(f"  执行结果：{screen_write.get('ok')}")
        if screen_write.get("reason"):
            print(f"  原因：{screen_write.get('reason')}")
        print(f"  点击坐标：{screen_write.get('target')}")
        print(f"  表格 bounds：{screen_write.get('table_bounds')}")
        print(f"  cell_width：{screen_write.get('cell_width')}")
        print(f"  提交键：{screen_write.get('commit_key')}")

    print_text_control_changes(
        "直接写入后新增/变化文本控件：", report.get("changed_text_controls") or []
    )

    editor = report.get("editor_probe")
    print("F2 编辑器激活探测：")
    if not editor:
        print("  未执行：直接 JAB 写入已成功或前置阶段未到达。")
    else:
        print(f"  F2 是否发送：{editor.get('ok')}")
        print(
            f"  发送目标：{editor.get('key_target')} hwnd={editor.get('target_hwnd')}"
        )
        if editor.get("reason"):
            print(f"  原因：{editor.get('reason')}")
        foreground = editor.get("foreground") or {}
        table_window = editor.get("table_window") or {}
        print(
            "  前台窗口："
            f"hwnd={foreground.get('hwnd')} "
            f"title={foreground.get('title')!r} "
            f"class={foreground.get('class_name')!r} "
            f"root={foreground.get('root_hwnd')}"
        )
        print(
            "  明细表窗口："
            f"hwnd={table_window.get('hwnd')} "
            f"title={table_window.get('title')!r} "
            f"class={table_window.get('class_name')!r} "
            f"root={table_window.get('root_hwnd')}"
        )
    print_text_control_changes(
        "F2 后新增/变化文本控件：",
        report.get("changed_text_controls_after_f2") or [],
    )

    enter = report.get("enter_probe")
    print("Enter 编辑器激活探测：")
    if not enter:
        print("  未执行：直接 JAB 写入已成功或前置阶段未到达。")
    else:
        print(f"  Enter 是否发送：{enter.get('ok')}")
        print(f"  发送目标：{enter.get('key_target')} hwnd={enter.get('target_hwnd')}")
        if enter.get("reason"):
            print(f"  原因：{enter.get('reason')}")
        foreground = enter.get("foreground") or {}
        table_window = enter.get("table_window") or {}
        print(
            "  前台窗口："
            f"hwnd={foreground.get('hwnd')} "
            f"title={foreground.get('title')!r} "
            f"class={foreground.get('class_name')!r} "
            f"root={foreground.get('root_hwnd')}"
        )
        print(
            "  明细表窗口："
            f"hwnd={table_window.get('hwnd')} "
            f"title={table_window.get('title')!r} "
            f"class={table_window.get('class_name')!r} "
            f"root={table_window.get('root_hwnd')}"
        )
    print_text_control_changes(
        "Enter 后新增/变化文本控件：",
        report.get("changed_text_controls_after_enter") or [],
    )

    print_key_table("写入后明细表：", report.get("after_table") or {})
    if report.get("ok"):
        print("成功：金额单元格已写入并通过表格读值校验。")
    else:
        print("失败：金额单元格未通过表格读值校验；本次没有保存、没有暂存。")


def main():
    args = build_parser().parse_args()
    print_header(
        args.activation,
        args.key_target,
        args.write_candidate_amount,
        args.screen_write_amount,
        args.screen_commit,
        args.dump_nearby,
    )
    print()
    print(f"请在 {START_DELAY_SECONDS} 秒内切到 NC 收款单窗口...")
    time.sleep(START_DELAY_SECONDS)
    print("开始测试。")

    config = load_config(str(ROOT / "config.json"))
    report = {
        "launcher": "tmp_receipt_cell_probe_run.py",
        "stop_hotkey": STOP_HOTKEY,
        "target": {
            "row": TARGET_ROW,
            "col": TARGET_COL,
            "name": TARGET_NAME,
            "value": TARGET_VALUE,
        },
    }
    try:
        if is_stop_hotkey_pressed():
            report.update(
                {
                    "ok": False,
                    "stopped_by_hotkey": True,
                    "failed_step": "before-start",
                }
            )
            print_summary(report)
            wait_exit()
            return 1

        jab = JABOperator(config)
        try:
            jab.ensure_started()
            health = check_jab_ready(jab)
            report["jab_health"] = health
            if not health.get("ok"):
                report.update(
                    {
                        "ok": False,
                        "failed_step": "jab-health-check",
                        "reason": health.get("reason"),
                    }
                )
                print_summary(report)
                wait_exit()
                return 1

            report["before_table"] = read_body_table(jab, "before_cell_probe")
            if is_stop_hotkey_pressed():
                report.update(
                    {
                        "ok": False,
                        "stopped_by_hotkey": True,
                        "failed_step": "before-cell-probe",
                    }
                )
            else:
                report.update(
                    probe_one_cell(
                        jab,
                        args.activation,
                        args.key_target,
                        args.write_candidate_amount,
                        args.screen_write_amount,
                        args.screen_commit,
                        args.dump_nearby,
                    )
                )
                report["after_table"] = read_body_table(jab, "after_cell_probe")
        finally:
            jab.close()
    except Exception as exc:
        report.update(
            {
                "ok": False,
                "exception": type(exc).__name__,
                "reason": str(exc),
                "traceback": traceback.format_exc(),
            }
        )

    print_summary(report)
    wait_exit()
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
