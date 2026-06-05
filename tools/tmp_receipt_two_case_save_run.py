# 生命周期：T0 一次性（删除条件：两条收款单真实保存循环验证后删除）
# 覆盖的业务阶段：收款单自制录入-两案例保存循环
# 依赖的服务/环境：Windows Python、NC 收款单录入页、Java Access Bridge
# 运行方式：python tools/tmp_receipt_two_case_save_run.py

import ctypes
from dataclasses import dataclass
import subprocess
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import ReceiptEntryConfig  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready, print_jab_health_failure  # noqa: E402
from tools.receipt_account_reference_try import STOP_HOTKEY, is_stop_hotkey_pressed  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    collect_receipt_new_windows,
    detect_self_made_entry_state,
    find_new_buttons,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    find_receipt_header_form_field,
    read_body_table,
    set_receipt_header_form_field,
    set_text_by_control_name,
)
from tools.tmp_receipt_cell_probe_run import (  # noqa: E402
    mouse_click,
    move_mouse,
    send_hotkey_ctrl_a,
    send_text,
)
from tools.tmp_receipt_detail_main_line_run import (  # noqa: E402
    run_fee_only,
    write_detail_line_by_screen,
)


START_DELAY_SECONDS = 2
SAVE_SUCCESS_TIMEOUT = 12.0


@dataclass(frozen=True)
class TestCase:
    name: str
    document_date: str
    customer_code: str
    bank_label: str
    currency: str
    amount: str
    fee: str


TEST_CASES = [
    TestCase(
        name="无手续费",
        document_date="2026-04-02",
        customer_code="YW03200",
        bank_label="招行",
        currency="人民币",
        amount="1",
        fee="0",
    ),
    TestCase(
        name="有手续费",
        document_date="2026-04-02",
        customer_code="YW03200",
        bank_label="招行",
        currency="人民币",
        amount="2",
        fee="10",
    ),
]


def print_header():
    print("测试功能：收款单两案例真实保存循环")
    print()
    print("测试数据：")
    for index, case in enumerate(TEST_CASES, start=1):
        print(
            f"{index}. {case.name}: 日期={case.document_date}, 客户={case.customer_code}, "
            f"银行={case.bank_label}, 币种={case.currency}, 金额={case.amount}, 手续费={case.fee}"
        )
    print()
    print("本脚本会做：")
    print("1. 每条从【新增】入口进入【自制】")
    print("2. 写表头：财务组织、客户、单据日期")
    print("3. 写明细主行：货款、币种、收款银行账户、科目、金额、网银")
    print("4. 手续费非零时：Ctrl+I 增行，写手续费行，清账户，删多余空行")
    print("5. 前台守卫通过后发送 Ctrl+S 保存")
    print("6. 保存后等待【新增】再次出现，作为保存成功")
    print()
    print("不会做：关闭窗口、写 Excel 状态、处理非测试数据")
    print(f"紧急停止：按 {STOP_HOTKEY}")
    print(f"启动后等待：{START_DELAY_SECONDS} 秒，用来切到 NC 窗口")
    print("=" * 60)


def elapsed(start):
    return round(time.perf_counter() - start, 3)


def run_receipt_new_probe():
    cmd = [
        sys.executable,
        str(ROOT / "tools" / "receipt_new_probe.py"),
        "--method",
        "button",
        "--class-name",
        "SunAwtFrame",
        "--choose-self-made",
        "--wait",
        "0.8",
        "--summary",
    ]
    start = time.perf_counter()
    proc = subprocess.run(
        cmd,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    ok = proc.returncode == 0
    return {
        "ok": ok,
        "seconds": elapsed(start),
        "returncode": proc.returncode,
        "stdout": proc.stdout.strip(),
        "stderr": proc.stderr.strip(),
    }


def wait_new_visible(jab, timeout=SAVE_SUCCESS_TIMEOUT):
    start = time.perf_counter()
    deadline = time.time() + timeout
    last_count = 0
    while time.time() < deadline:
        buttons = find_new_buttons(jab, name_query="新增", class_name="SunAwtFrame")
        last_count = len(buttons)
        if buttons:
            return {
                "ok": True,
                "seconds": elapsed(start),
                "count": len(buttons),
                "first": buttons[0],
            }
        time.sleep(0.3)
    return {
        "ok": False,
        "seconds": elapsed(start),
        "count": last_count,
        "reason": "保存后未等到【新增】入口出现",
    }


def ensure_starts_from_new_state(jab):
    buttons = find_new_buttons(jab, name_query="新增", class_name="SunAwtFrame")
    if buttons:
        return {
            "ok": True,
            "state": "new-visible",
            "count": len(buttons),
            "first": buttons[0],
        }
    windows = collect_receipt_new_windows(jab)
    entry_state = detect_self_made_entry_state(windows)
    if entry_state.get("ok"):
        return {
            "ok": True,
            "state": "existing-self-made-entry",
            "entry_state": entry_state,
        }
    entry = read_body_table(jab, "start_state_body_probe")
    return {
        "ok": False,
        "reason": "当前没有发现【新增】入口，也不是自制录入态。为避免误保存旧页面，本轮不继续。",
        "body_table": entry,
    }


def business_for_case(config, case):
    receipt_config = ReceiptEntryConfig(config)
    account = receipt_config.account_for_bank(case.bank_label)
    if not account:
        raise RuntimeError(f"config.json 中找不到银行账户映射：{case.bank_label}")
    org = receipt_config.organizations.get(account.organization_code)
    if not org:
        raise RuntimeError(f"账号 {case.bank_label} 绑定的财务组织不存在")
    candidates = account.nc_candidates(case.currency)
    bank_account = candidates[0] if candidates else account.account_no
    return {
        "finance_org_code": org.code,
        "finance_org_name": org.name,
        "document_date": case.document_date,
        "customer_code": case.customer_code,
        "currency": case.currency,
        "bank_label": case.bank_label,
        "bank_account": bank_account,
        "amount": case.amount,
        "fee": case.fee,
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
    }


def fill_minimal_header(jab, business):
    steps = []
    for item in [
        ("财务组织", business["finance_org_code"], "control_name", False),
        ("客户", business["customer_code"], "header_form", True),
        ("单据日期", business["document_date"], "header_form", False),
    ]:
        label, value, method, require_strict = item
        start = time.perf_counter()
        if method == "control_name":
            result = set_text_by_control_name(jab, "财务组织(O)", value)
        elif label == "客户":
            result = set_customer_header_field(jab, value)
        else:
            result = set_receipt_header_form_field(jab, label, value)
        result = dict(result)
        result.update(
            {
                "step": "header",
                "label": label,
                "value": value,
                "seconds": elapsed(start),
            }
        )
        result["strict_ok"] = bool(result.get("ok"))
        result["require_strict"] = require_strict
        result["soft_ok"] = (
            result["strict_ok"] if require_strict else header_step_soft_ok(result)
        )
        steps.append(result)
        if not result.get("soft_ok"):
            break
    return steps


def header_step_soft_ok(result):
    if result.get("ok"):
        return True
    # NC/JAB 对部分表头字段会出现 setTextContents 成功但读回为空。
    # T0 保存循环以保存后【新增】为最终 oracle，因此这里不因读回空阻断。
    if result.get("path") and not result.get("exception"):
        return True
    return False


def set_customer_header_field(jab, value):
    result = dict(set_receipt_header_form_field(jab, "客户", value, commit_key="tab"))
    if result.get("ok"):
        result["screen_fallback"] = {"skipped": True, "reason": "JAB 后台写入已成功"}
        return result

    fallback = screen_write_header_field(jab, "客户", value)
    result["screen_fallback"] = fallback
    if not fallback.get("ok"):
        result["ok"] = False
        result["reason"] = fallback.get("reason") or "客户屏幕输入兜底失败"
        return result

    checked = wait_header_field_non_empty(jab, "客户", timeout=4.0)
    result["screen_fallback_check"] = checked
    result["ok"] = bool(checked.get("ok"))
    result["reason"] = None if result["ok"] else checked.get("reason")
    if result["ok"]:
        result["text_after"] = checked.get("text")
        result["description_after"] = checked.get("description")
    return result


def screen_write_header_field(jab, label, value):
    found = find_receipt_header_form_field(jab, label)
    if not found.get("ok"):
        return {
            "ok": False,
            "label": label,
            "reason": found.get("reason") or "表头字段未找到，无法屏幕输入",
            "found": found,
        }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    try:
        info = jab.get_context_info(vm_id, context)
        if not info:
            return {"ok": False, "label": label, "reason": "无法读取客户字段 bounds"}
        bounds = [info.x, info.y, info.width, info.height]
        if info.x < 0 or info.y < 0 or info.width <= 0 or info.height <= 0:
            return {
                "ok": False,
                "label": label,
                "reason": f"客户字段 bounds 不可见：{bounds}",
                "bounds": bounds,
            }
        window = found.get("window") or {}
        guard = same_nc_root_foreground(window)
        if not guard.get("ok"):
            return {
                "ok": False,
                "label": label,
                "reason": guard.get("reason"),
                "guard": guard,
                "bounds": bounds,
            }
        target_x = int(info.x + info.width / 2)
        target_y = int(info.y + info.height / 2)
        try:
            if hasattr(jab.dll, "requestFocus"):
                jab.dll.requestFocus(vm_id, context)
                time.sleep(0.1)
            move_mouse(target_x, target_y)
            mouse_click()
            time.sleep(0.08)
            mouse_click()
            time.sleep(0.15)
            send_hotkey_ctrl_a()
            time.sleep(0.08)
            send_text(value)
            time.sleep(0.3)
            commit = click_header_field_center(jab, "单据日期")
            time.sleep(0.8)
        except Exception as exc:
            return {
                "ok": False,
                "label": label,
                "reason": f"客户屏幕输入失败：{type(exc).__name__}: {exc}",
                "target": [target_x, target_y],
                "bounds": bounds,
                "guard": guard,
            }
        return {
            "ok": True,
            "label": label,
            "method": "guarded_screen_text",
            "target": [target_x, target_y],
            "bounds": bounds,
            "guard": guard,
            "commit": commit,
            "path": found.get("path"),
            "fallback_path_used": bool(found.get("fallback_path_used")),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def click_header_field_center(jab, label):
    found = find_receipt_header_form_field(jab, label)
    if not found.get("ok"):
        return {
            "ok": False,
            "label": label,
            "reason": found.get("reason") or "提交目标字段未找到",
            "found": found,
        }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    try:
        info = jab.get_context_info(vm_id, context)
        if not info:
            return {"ok": False, "label": label, "reason": "无法读取提交目标 bounds"}
        bounds = [info.x, info.y, info.width, info.height]
        if info.x < 0 or info.y < 0 or info.width <= 0 or info.height <= 0:
            return {
                "ok": False,
                "label": label,
                "reason": f"提交目标 bounds 不可见：{bounds}",
                "bounds": bounds,
            }
        window = found.get("window") or {}
        guard = same_nc_root_foreground(window)
        if not guard.get("ok"):
            return {
                "ok": False,
                "label": label,
                "reason": guard.get("reason"),
                "guard": guard,
                "bounds": bounds,
            }
        target_x = int(info.x + info.width / 2)
        target_y = int(info.y + info.height / 2)
        move_mouse(target_x, target_y)
        mouse_click()
        return {
            "ok": True,
            "label": label,
            "target": [target_x, target_y],
            "bounds": bounds,
            "guard": guard,
            "path": found.get("path"),
            "fallback_path_used": bool(found.get("fallback_path_used")),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def read_header_field_non_empty(jab, label):
    found = find_receipt_header_form_field(jab, label)
    if not found.get("ok"):
        return {
            "ok": False,
            "label": label,
            "reason": found.get("reason") or "表头字段未找到",
            "found": found,
        }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    try:
        info = jab.get_context_info(vm_id, context)
        text = str(jab.get_text_context_value(vm_id, context) or "").strip()
        description = info.description.strip() if info else ""
        name = info.name.strip() if info else ""
        non_empty = bool(text or description)
        return {
            "ok": non_empty,
            "label": label,
            "path": found.get("path"),
            "fallback_path_used": bool(found.get("fallback_path_used")),
            "text": text,
            "description": description,
            "name": name,
            "reason": None if non_empty else f"表头【{label}】为空",
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def wait_header_field_non_empty(jab, label, timeout=3.0):
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        last = read_header_field_non_empty(jab, label)
        if last.get("ok"):
            last["wait_seconds"] = round(timeout - max(deadline - time.time(), 0), 3)
            return last
        time.sleep(0.3)
    if last is None:
        return {
            "ok": False,
            "label": label,
            "reason": f"表头【{label}】非空检测没有取得结果",
            "wait_seconds": timeout,
        }
    last["wait_seconds"] = timeout
    return last


def check_customer_non_empty(jab, stage):
    result = wait_header_field_non_empty(jab, "客户", timeout=3.0)
    result["stage"] = stage
    return result


def fill_main_detail(jab, business):
    start = time.perf_counter()
    located = locate_receipt_body_table(jab, max_rows=3)
    if not located.get("best"):
        return (
            located,
            [{"ok": False, "name": "明细表", "reason": "body table not found"}],
            elapsed(start),
        )
    steps = write_detail_line_by_screen(jab, business, located)
    return located, steps, elapsed(start)


def fill_fee_if_needed(jab, located, case):
    if str(case.fee).strip() in ("", "0", "0.00"):
        return {
            "ok": True,
            "skipped": True,
            "reason": "手续费为 0，跳过手续费分支",
            "seconds": 0.0,
        }
    start = time.perf_counter()
    add_row, steps, clear_account, delete_extra = run_fee_only(jab, located, case.fee)
    ok = (
        bool(add_row.get("ok"))
        and all(bool(step.get("ok")) for step in steps)
        and bool(clear_account.get("ok"))
        and bool(delete_extra.get("ok"))
    )
    return {
        "ok": ok,
        "skipped": False,
        "seconds": elapsed(start),
        "add_row": add_row,
        "steps": steps,
        "clear_account": clear_account,
        "delete_extra": delete_extra,
    }


def same_nc_root_foreground(table_window):
    if sys.platform != "win32":
        return {"ok": False, "reason": "必须在 Windows Python 下运行"}
    user32 = ctypes.windll.user32
    table_hwnd = int((table_window or {}).get("hwnd") or 0)
    table_root = user32.GetAncestor(table_hwnd, 2) if table_hwnd else 0
    foreground = user32.GetForegroundWindow()
    foreground_root = user32.GetAncestor(foreground, 2) if foreground else 0
    ok = bool(
        foreground
        and table_hwnd
        and (
            foreground == table_hwnd
            or foreground == table_root
            or foreground_root == table_root
        )
    )
    return {
        "ok": ok,
        "foreground": int(foreground or 0),
        "foreground_root": int(foreground_root or 0),
        "table_hwnd": table_hwnd,
        "table_root": int(table_root or 0),
        "reason": None
        if ok
        else "当前前台窗口不是本次定位到的 NC 收款单窗口，未发送 Ctrl+S",
    }


def send_ctrl_s():
    try:
        send_virtual_key(0x11, key_up=False)
        send_virtual_key(0x53, key_up=False)
        send_virtual_key(0x53, key_up=True)
        send_virtual_key(0x11, key_up=True)
    except Exception:
        try:
            send_virtual_key(0x11, key_up=True)
        finally:
            raise


def send_virtual_key(vk, key_up=False):
    inp = INPUT()
    inp.type = 1
    inp.ki = KEYBDINPUT(vk, 0, 0x0002 if key_up else 0, 0, None)
    ctypes.windll.kernel32.SetLastError(0)
    sent = ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    if sent != 1:
        error_code = ctypes.windll.kernel32.GetLastError()
        raise RuntimeError(
            f"SendInput failed, vk={vk}, key_up={key_up}, error={error_code}"
        )


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
        _fields_ = [("ki", KEYBDINPUT)]

    _anonymous_ = ("union",)
    _fields_ = [("type", ctypes.c_ulong), ("union", _INPUT_UNION)]


def save_and_wait_new(jab, located):
    start = time.perf_counter()
    best = located.get("best") or {}
    guard = same_nc_root_foreground(best.get("window") or {})
    if not guard.get("ok"):
        return {"ok": False, "seconds": elapsed(start), "guard": guard}
    try:
        send_ctrl_s()
    except Exception as exc:
        return {
            "ok": False,
            "seconds": elapsed(start),
            "guard": guard,
            "reason": f"发送 Ctrl+S 失败：{type(exc).__name__}: {exc}",
        }
    time.sleep(0.8)
    waited = wait_new_visible(jab, timeout=SAVE_SUCCESS_TIMEOUT)
    return {
        "ok": bool(waited.get("ok")),
        "seconds": elapsed(start),
        "guard": guard,
        "wait_new": waited,
    }


def print_case_summary(case_report):
    print()
    print(f"案例：{case_report.get('name')}")
    print(f"  总耗时：{case_report.get('seconds')} 秒")
    for item in case_report.get("timings", []):
        print(f"  - {item['name']}: {item['seconds']} 秒")
    header_steps = case_report.get("header_steps") or []
    if header_steps:
        print("  表头：")
        for step in header_steps:
            print(
                f"    {step.get('label')}: "
                f"{'严格成功' if step.get('strict_ok') else '软通过' if step.get('soft_ok') else '失败'} "
                f"| path={step.get('path')}"
            )
            if step.get("fallback_path_used"):
                print("      注意：本字段使用了固定 path 兜底。")
            fallback = step.get("screen_fallback") or {}
            if fallback:
                if fallback.get("skipped"):
                    print("      屏幕兜底：未执行。")
                else:
                    print(
                        "      屏幕兜底："
                        f"{'成功' if fallback.get('ok') else '失败'} "
                        f"| target={fallback.get('target')} "
                        f"| bounds={fallback.get('bounds')}"
                    )
                    commit = fallback.get("commit") or {}
                    if commit:
                        print(
                            "      兜底提交："
                            f"{'成功' if commit.get('ok') else '失败'} "
                            f"| label={commit.get('label')} "
                            f"| target={commit.get('target')} "
                            f"| bounds={commit.get('bounds')}"
                        )
                    if not fallback.get("ok"):
                        print(f"      兜底原因：{fallback.get('reason')}")
            fallback_check = step.get("screen_fallback_check") or {}
            if fallback_check:
                print(
                    "      兜底后检测："
                    f"{'通过' if fallback_check.get('ok') else '失败'} "
                    f"| wait={fallback_check.get('wait_seconds')}s "
                    f"| text={fallback_check.get('text')!r}, "
                    f"description={fallback_check.get('description')!r}"
                )
            if not step.get("strict_ok"):
                state = step.get("backend_state") or {}
                print(
                    "      读回："
                    f"text={state.get('text')!r}, "
                    f"description={state.get('description')!r}, "
                    f"reason={step.get('reason')}"
                )
    customer_checks = case_report.get("customer_checks") or []
    if customer_checks:
        print("  客户非空检测：")
        for check in customer_checks:
            print(
                f"    {check.get('stage')}: "
                f"{'通过' if check.get('ok') else '失败'} | "
                f"wait={check.get('wait_seconds')}s | "
                f"text={check.get('text')!r}, description={check.get('description')!r}"
            )
            if check.get("fallback_path_used"):
                print("      注意：本次读取使用了固定 path 兜底。")
            if not check.get("ok"):
                print(f"      原因：{check.get('reason')}")
    if case_report.get("ok"):
        print("  结果：成功，保存后已看到【新增】。")
    else:
        print(f"  结果：失败，停止位置={case_report.get('failed_step')}")
        reason = case_report.get("reason")
        if reason:
            print(f"  原因：{reason}")


def run_one_case(config, case, allow_existing_entry=False):
    case_start = time.perf_counter()
    case_report = {"name": case.name, "timings": []}
    jab = JABOperator(config)
    try:
        jab.ensure_started()
        start_state = ensure_starts_from_new_state(jab)
        case_report["start_state"] = start_state
        if not start_state.get("ok"):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "start-new-state",
                    "reason": start_state.get("reason"),
                    "seconds": elapsed(case_start),
                }
            )
            return case_report
        if (
            start_state.get("state") == "existing-self-made-entry"
            and not allow_existing_entry
        ):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "start-new-state",
                    "reason": "当前已在自制录入态，但本案例必须从保存后【新增】入口开始。",
                    "seconds": elapsed(case_start),
                }
            )
            return case_report
    finally:
        jab.close()

    if start_state.get("state") == "existing-self-made-entry":
        opened = {
            "ok": True,
            "seconds": 0.0,
            "method": "reuse-existing-self-made-entry",
        }
    else:
        opened = run_receipt_new_probe()
    case_report["open_self_made"] = opened
    case_report["timings"].append({"name": "新增->自制", "seconds": opened["seconds"]})
    if not opened.get("ok"):
        case_report.update(
            {
                "ok": False,
                "failed_step": "open-self-made",
                "reason": "新增->自制失败",
                "seconds": elapsed(case_start),
            }
        )
        return case_report

    jab = JABOperator(config)
    try:
        jab.ensure_started()
        business = business_for_case(config, case)
        case_report["business"] = business
        case_report["customer_checks"] = []

        header_start = time.perf_counter()
        header_steps = fill_minimal_header(jab, business)
        case_report["header_steps"] = header_steps
        case_report["timings"].append(
            {"name": "表头", "seconds": elapsed(header_start)}
        )
        if not all(step.get("soft_ok") for step in header_steps):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "fill-header",
                    "reason": "至少一个表头字段失败",
                    "seconds": elapsed(case_start),
                }
            )
            return case_report

        customer_after_header = check_customer_non_empty(jab, "表头写完后")
        case_report["customer_checks"].append(customer_after_header)
        if not customer_after_header.get("ok"):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "check-customer-after-header",
                    "reason": "客户字段为空，未进入明细和保存",
                    "seconds": elapsed(case_start),
                }
            )
            return case_report

        located, detail_steps, detail_seconds = fill_main_detail(jab, business)
        case_report["detail_steps"] = detail_steps
        case_report["timings"].append({"name": "主行", "seconds": detail_seconds})
        if not all(step.get("ok") for step in detail_steps):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "fill-main-detail",
                    "reason": "至少一个主行字段失败",
                    "seconds": elapsed(case_start),
                }
            )
            return case_report

        fee_report = fill_fee_if_needed(jab, located, case)
        case_report["fee"] = fee_report
        case_report["timings"].append(
            {"name": "手续费", "seconds": fee_report.get("seconds", 0.0)}
        )
        if not fee_report.get("ok"):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "fill-fee",
                    "reason": "手续费分支失败",
                    "seconds": elapsed(case_start),
                }
            )
            return case_report

        customer_before_save = check_customer_non_empty(jab, "保存前")
        case_report["customer_checks"].append(customer_before_save)
        if not customer_before_save.get("ok"):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "check-customer-before-save",
                    "reason": "保存前客户字段为空，未发送 Ctrl+S",
                    "seconds": elapsed(case_start),
                }
            )
            return case_report

        refreshed = locate_receipt_body_table(jab, max_rows=5)
        save = save_and_wait_new(jab, refreshed)
        case_report["save"] = save
        case_report["timings"].append(
            {"name": "保存并等新增", "seconds": save["seconds"]}
        )
        case_report["ok"] = bool(save.get("ok"))
        if not case_report["ok"]:
            case_report["failed_step"] = "save-wait-new"
            case_report["reason"] = "Ctrl+S 后未等到【新增】"
        case_report["after_table"] = read_body_table(jab, "after_case")
        case_report["seconds"] = elapsed(case_start)
        return case_report
    finally:
        jab.close()


def main():
    print_header()
    print()
    print(
        f"请在 {START_DELAY_SECONDS} 秒内切到 NC【收款单录入】且能看到【新增】的页面..."
    )
    time.sleep(START_DELAY_SECONDS)
    print("开始真实保存循环测试。")

    report = {
        "launcher": "tmp_receipt_two_case_save_run.py",
        "start_delay_seconds": START_DELAY_SECONDS,
        "stop_hotkey": STOP_HOTKEY,
        "cases": [],
    }
    try:
        if is_stop_hotkey_pressed():
            print(f"检测到紧急停止键 {STOP_HOTKEY}，未开始。")
            return 1
        config = load_config(str(ROOT / "config.json"))
        health_jab = JABOperator(config)
        try:
            health_jab.ensure_started()
            health = check_jab_ready(health_jab)
            report["jab_health"] = health
            if not health.get("ok"):
                print("JAB 启动状态：")
                print_jab_health_failure(health)
                return 1
        finally:
            health_jab.close()

        for index, case in enumerate(TEST_CASES):
            if is_stop_hotkey_pressed():
                print(f"检测到紧急停止键 {STOP_HOTKEY}，停止后续案例。")
                break
            case_report = run_one_case(config, case, allow_existing_entry=(index == 0))
            report["cases"].append(case_report)
            print_case_summary(case_report)
            if not case_report.get("ok"):
                break
            time.sleep(0.8)
    except Exception as exc:
        report["exception"] = type(exc).__name__
        report["reason"] = str(exc)
        report["traceback"] = traceback.format_exc()
        print()
        print("脚本异常：")
        print(f"  {type(exc).__name__}: {exc}")
        return 1

    ok = len(report["cases"]) == len(TEST_CASES) and all(
        item.get("ok") for item in report["cases"]
    )
    print()
    print("总结果：")
    print(f"  {'成功' if ok else '失败'}")
    total = sum(float(item.get("seconds") or 0) for item in report["cases"])
    print(f"  两案例累计耗时：{round(total, 3)} 秒")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
