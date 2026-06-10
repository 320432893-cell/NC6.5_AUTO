# 生命周期：T0 一次性（删除条件：两条收款单真实保存循环验证后删除）
# 覆盖的业务阶段：收款单自制录入-两案例保存循环
# 依赖的服务/环境：Windows Python、NC 收款单录入页、Java Access Bridge
# 运行方式：python tools/tmp_receipt_two_case_save_run.py

import ctypes
from dataclasses import dataclass
import os
import re
import sys
import time
import traceback
from types import SimpleNamespace
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import ReceiptEntryConfig, ReceiptEntryWorkbook  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready, print_jab_health_failure  # noqa: E402
from tools.receipt_account_reference_try import STOP_HOTKEY, is_stop_hotkey_pressed  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    collect_receipt_new_windows,
    detect_self_made_entry_state,
    find_new_buttons,
    run as run_receipt_new_probe_in_process,
    summarize_report as summarize_receipt_new_probe_report,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    find_context_with_window,
    find_receipt_header_form_field,
    find_receipt_header_form_field_by_path,
    locate_receipt_header_scope,
    read_body_table,
)
from tools.tmp_receipt_cell_probe_run import (  # noqa: E402
    mouse_click,
    move_mouse,
    send_hotkey_ctrl_a,
    send_text,
    send_virtual_key,
)
from tools.tmp_receipt_detail_main_line_run import (  # noqa: E402
    run_fee_only,
    write_detail_line_by_screen,
)


START_DELAY_SECONDS = 2
SAVE_SUCCESS_TIMEOUT = 12.0
SAVE_ENABLED = False
READ_DETAIL_BEFORE_SAVE_SKIPPED = False
ALLOW_EXISTING_ENTRY_FOR_FIRST_CASE = (
    os.environ.get("RECEIPT_ALLOW_EXISTING_ENTRY", "").strip().lower()
    in {"1", "true", "yes", "y"}
)
TEST_CASE_LIMIT = 2
TEST_ORG_CODE = os.environ.get("RECEIPT_TEST_ORG_CODE", "A001")
TEST_PREFERRED_CURRENCY = os.environ.get("RECEIPT_TEST_CURRENCY", "美元")
TEST_BANK_LABEL = os.environ.get("RECEIPT_TEST_BANK_LABEL", "招行")
TEST_BANK_ACCOUNT_NO = os.environ.get(
    "RECEIPT_TEST_BANK_ACCOUNT_NO", "FTE1219165931831"
)
TEST_FEE_OVERRIDES = ("20.00", "33.00")


@dataclass(frozen=True)
class TestCase:
    name: str
    document_date: str
    customer_code: str
    bank_label: str
    currency: str
    amount: str
    fee: str
    bank_account_no: str
    excel_row: int | None = None
    payer_name: str = ""
    source_bank: str = ""


def print_header(test_cases):
    print("测试功能：收款单两案例真实保存循环")
    print()
    print("测试数据：")
    for index, case in enumerate(test_cases, start=1):
        print(
            f"{index}. {case.name}: Excel行={case.excel_row}, 日期={case.document_date}, "
            f"客户={case.customer_code}, 原Excel银行={case.source_bank}, "
            f"录入银行={case.bank_label}, 账号={case.bank_account_no}, "
            f"币种={case.currency}, 金额={case.amount}, 手续费={case.fee}"
        )
    print()
    print(
        f"选数口径：最近 {TEST_CASE_LIMIT} 条有效行，"
        f"主体={TEST_ORG_CODE}，币种={TEST_PREFERRED_CURRENCY}；"
        f"收款银行账号固定用 {TEST_BANK_ACCOUNT_NO}；"
        f"手续费用脚本测试值覆盖={TEST_FEE_OVERRIDES}"
    )
    print()
    print("本脚本会做：")
    print("1. 每条从【新增】入口进入【自制】")
    print("2. 写表头：财务组织、客户、单据日期、币种")
    print("3. 写明细主行：货款、收款银行账户、科目、金额、网银")
    print("4. 手续费非零时：Ctrl+I 增行，写手续费行，清账户，删多余空行")
    if SAVE_ENABLED:
        print("5. 前台守卫通过后发送 Ctrl+S 保存")
        print("6. 保存后等待【新增】再次出现，作为保存成功")
    else:
        print("5. 保存前停止：本轮只诊断填写，不发送 Ctrl+S")
    print()
    print("不会做：关闭窗口、写 Excel 状态、处理非测试数据")
    if not SAVE_ENABLED:
        print("不会做：保存、暂存")
    print(f"紧急停止：按 {STOP_HOTKEY}")
    print(f"启动后等待：{START_DELAY_SECONDS} 秒，用来切到 NC 窗口")
    print("=" * 60)


def elapsed(start):
    return round(time.perf_counter() - start, 3)


def build_latest_test_cases(config, limit=TEST_CASE_LIMIT):
    workbook = ReceiptEntryWorkbook(config)
    rows, issues, summary = workbook.build_local_plan(write_sheet=False)
    issue_rows = {issue.excel_row for issue in issues if issue.excel_row is not None}
    candidates = [
        row
        for row in rows
        if row.row not in issue_rows
        and row.organization_code == TEST_ORG_CODE
        and row.currency == TEST_PREFERRED_CURRENCY
    ]
    candidates.sort(key=lambda row: (row.receipt_date, row.row), reverse=True)
    selected = candidates[:limit]
    if len(selected) < limit:
        raise RuntimeError(
            f"有效测试行不足 {limit} 条：主体={TEST_ORG_CODE}, "
            f"可选={len(selected)}, 总计划行={summary.get('rows')}"
        )
    return [
        TestCase(
            name=f"最近有效行{index}",
            excel_row=row.row,
            document_date=row.receipt_date.isoformat(),
            customer_code=row.customer_code,
            payer_name=row.payer_name,
            source_bank=row.bank,
            bank_label=TEST_BANK_LABEL,
            bank_account_no=TEST_BANK_ACCOUNT_NO,
            currency=row.currency,
            amount=str(row.raw_amount),
            fee=TEST_FEE_OVERRIDES[(index - 1) % len(TEST_FEE_OVERRIDES)],
        )
        for index, row in enumerate(selected, start=1)
    ]


def run_receipt_new_probe(jab=None, before=None, buttons=None):
    start = time.perf_counter()
    args = SimpleNamespace(
        config=str(ROOT / "config.json"),
        method="button",
        path=None,
        title=None,
        class_name="SunAwtFrame",
        name="新增",
        role=None,
        action=None,
        return_timeout=0.2,
        wait=0.55,
        choose_self_made=True,
        self_made_index=0,
        json=False,
        summary=True,
    )
    report = run_receipt_new_probe_in_process(
        args, jab=jab, before=before, buttons=buttons
    )
    parsed = summarize_receipt_new_probe_report(report)
    entry_state = report.get("entry_state") or {}
    ok = (
        bool((report.get("open") or {}).get("ok"))
        and bool((report.get("choose_self_made") or {}).get("ok"))
        and bool(entry_state.get("ok") or entry_state.get("partial_ok"))
    )
    return {
        "ok": ok,
        "seconds": elapsed(start),
        "returncode": 0 if ok else 1,
        "stdout": "",
        "stderr": "",
        "parsed_summary": parsed,
        "timings": parsed.get("timings") or [],
        "raw_report": report,
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
    windows = collect_receipt_new_windows(jab)
    buttons = find_new_buttons(jab, name_query="新增", class_name="SunAwtFrame")
    if buttons:
        return {
            "ok": True,
            "state": "new-visible",
            "count": len(buttons),
            "first": buttons[0],
            "buttons": buttons,
            "windows": windows,
        }
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
    bank_account = case.bank_account_no or account.account_no
    if TEST_BANK_ACCOUNT_NO and bank_account != TEST_BANK_ACCOUNT_NO:
        raise RuntimeError(
            f"本轮测试只允许账号 {TEST_BANK_ACCOUNT_NO}，实际将使用 {bank_account}"
        )
    header_currency_code = account.header_currency_code
    if not header_currency_code:
        raise RuntimeError(f"账号 {bank_account} 未配置表头币种代码")
    return {
        "finance_org_code": org.code,
        "finance_org_name": org.name,
        "finance_org_short_name": org.short_name,
        "document_date": case.document_date,
        "customer_code": case.customer_code,
        "payer_name": case.payer_name,
        "currency": case.currency,
        "header_currency_code": header_currency_code,
        "bank_label": case.bank_label,
        "source_bank": case.source_bank,
        "bank_account": bank_account,
        "amount": case.amount,
        "fee": case.fee,
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
    }


def fill_minimal_header(jab, business, header_cache=None, scope_hwnd=None):
    header_cache = header_cache if header_cache is not None else {}
    steps = []
    for item in [
        ("财务组织", business["finance_org_code"], "header_form", True),
        ("客户", business["customer_code"], "header_form", True),
        ("单据日期", business["document_date"], "header_form", True),
        ("币种", business["header_currency_code"], "header_form", True),
        ("结算方式", business["settlement"], "header_form", True),
    ]:
        label, value, method, require_strict = item
        start = time.perf_counter()
        if label == "财务组织":
            result = set_finance_org_header_field(jab, value, header_cache, scope_hwnd)
        elif label == "客户":
            result = set_customer_header_field(jab, value, header_cache, scope_hwnd)
        elif label == "单据日期":
            result = set_document_date_header_field(jab, value, header_cache, scope_hwnd)
        elif label == "币种":
            result = set_currency_header_field(jab, value, header_cache, scope_hwnd)
        else:
            result = screen_write_header_field(jab, label, value, header_cache, scope_hwnd)
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
        remember_header_field(header_cache, label, result)
        if not result.get("soft_ok"):
            break
    return steps


def header_step_soft_ok(result):
    if result.get("ok"):
        return True
    return False


def set_finance_org_header_field(jab, value, header_cache=None, scope_hwnd=None):
    result = screen_write_header_field(jab, "财务组织", value, header_cache, scope_hwnd)
    result.update(
        {
            "label": "财务组织",
            "value": value,
            "method": "screen_input_enter",
        }
    )
    return result


def set_customer_header_field(jab, value, header_cache=None, scope_hwnd=None):
    fallback = screen_write_header_field(jab, "客户", value, header_cache, scope_hwnd)
    result = {
        "ok": bool(fallback.get("ok")),
        "label": "客户",
        "value": value,
        "method": "screen_input_enter",
        "path": fallback.get("path"),
        "fallback_path_used": bool(fallback.get("fallback_path_used")),
        "backend_write": {
            "skipped": True,
            "reason": "客户是 NC 参照型字段，JAB setTextContents 不触发解析，直接使用真实输入",
        },
        "screen_fallback": fallback,
    }
    if not fallback.get("ok"):
        result["ok"] = False
        result["reason"] = fallback.get("reason") or "客户屏幕输入兜底失败"
        return result
    result["commit_check"] = {
        "skipped": True,
        "reason": "客户输入后已 Enter 提交，不原地等待读回；保存前统一回看",
    }
    return result


def set_document_date_header_field(jab, value, header_cache=None, scope_hwnd=None):
    fallback = screen_write_header_field(jab, "单据日期", value, header_cache, scope_hwnd)
    result = {
        "ok": bool(fallback.get("ok")),
        "label": "单据日期",
        "value": value,
        "method": "screen_input_enter",
        "path": fallback.get("path"),
        "fallback_path_used": bool(fallback.get("fallback_path_used")),
        "backend_write": {
            "skipped": True,
            "reason": "日期字段 JAB setTextContents 读回迟钝，直接使用真实输入并失焦提交",
        },
        "screen_fallback": fallback,
    }
    if not fallback.get("ok"):
        result["ok"] = False
        result["reason"] = fallback.get("reason") or "单据日期屏幕输入失败"
        return result
    result["commit_check"] = {
        "skipped": True,
        "reason": "日期写入后不原地等待，保存前统一回看",
    }
    return result


def set_currency_header_field(jab, value, header_cache=None, scope_hwnd=None):
    fallback = screen_write_header_field(jab, "币种", value, header_cache, scope_hwnd)
    result = {
        "ok": bool(fallback.get("ok")),
        "label": "币种",
        "value": value,
        "method": "screen_input_enter",
        "path": fallback.get("path"),
        "fallback_path_used": bool(fallback.get("fallback_path_used")),
        "backend_write": {
            "skipped": True,
            "reason": "币种必须写表头，明细币种不会同步回表头",
        },
        "screen_fallback": fallback,
    }
    if not fallback.get("ok"):
        result["ok"] = False
        result["reason"] = fallback.get("reason") or "币种屏幕输入失败"
        return result
    result["commit_check"] = {
        "skipped": True,
        "reason": "币种写入后不原地等待，保存前统一回看",
    }
    return result


def screen_write_header_field(jab, label, value, header_cache=None, scope_hwnd=None):
    found = find_cached_header_field(jab, label, header_cache, scope_hwnd)
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
            return {
                "ok": False,
                "label": label,
                "reason": f"无法读取{label}字段 bounds",
            }
        bounds = [info.x, info.y, info.width, info.height]
        if info.x < 0 or info.y < 0 or info.width <= 0 or info.height <= 0:
            return {
                "ok": False,
                "label": label,
                "reason": f"{label}字段 bounds 不可见：{bounds}",
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
                time.sleep(0.05)
            move_mouse(target_x, target_y)
            mouse_click()
            time.sleep(0.04)
            mouse_click()
            time.sleep(0.08)
            send_hotkey_ctrl_a()
            time.sleep(0.04)
            send_text(value)
            time.sleep(0.08)
            press_virtual_key(0x0D)
            commit = {"ok": True, "key": "Enter"}
            time.sleep(0.08)
        except Exception as exc:
            return {
                "ok": False,
                "label": label,
                "reason": f"{label}屏幕输入失败：{type(exc).__name__}: {exc}",
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
            "window": window,
            "cache_hit": bool(found.get("cache_hit")),
            "fallback_path_used": bool(found.get("fallback_path_used")),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def screen_write_control_by_name(jab, control_name, value, header_cache=None):
    cached_label = "财务组织" if control_name == "财务组织(O)" else control_name
    cached = find_cached_header_field(jab, cached_label, header_cache)
    if cached.get("ok"):
        context = cached["context"]
        vm_id = cached["vm_id"]
        owned_contexts = cached["owned_contexts"]
        window = cached.get("window") or {}
        path = cached.get("path")
    else:
        context, vm_id, owned_contexts, owned_indexes, window = (
            find_context_with_window(
                jab,
                control_name,
                roles=("text",),
                timeout=3.0,
                require_showing=True,
                window_class="SunAwtCanvas",
                visible_only=True,
            )
        )
        path = "0" + "".join(f".{index}" for index in owned_indexes)
    if not context:
        return {"ok": False, "reason": f"控件未找到：{control_name}"}
    try:
        info = jab.get_context_info(vm_id, context)
        if not info:
            return {"ok": False, "reason": f"无法读取控件 bounds：{control_name}"}
        bounds = [info.x, info.y, info.width, info.height]
        if info.x < 0 or info.y < 0 or info.width <= 0 or info.height <= 0:
            return {
                "ok": False,
                "reason": f"控件 bounds 不可见：{bounds}",
                "bounds": bounds,
                "path": path,
            }
        guard = same_nc_root_foreground(window)
        if not guard.get("ok"):
            return {
                "ok": False,
                "reason": guard.get("reason"),
                "guard": guard,
                "bounds": bounds,
                "path": path,
            }
        target_x = int(info.x + info.width / 2)
        target_y = int(info.y + info.height / 2)
        if hasattr(jab.dll, "requestFocus"):
            jab.dll.requestFocus(vm_id, context)
            time.sleep(0.03)
        move_mouse(target_x, target_y)
        mouse_click()
        time.sleep(0.04)
        mouse_click()
        time.sleep(0.06)
        send_hotkey_ctrl_a()
        time.sleep(0.03)
        send_text(value)
        time.sleep(0.08)
        press_virtual_key(0x0D)
        time.sleep(0.08)
        info_after = jab.get_context_info(vm_id, context)
        return {
            "ok": True,
            "path": path,
            "cache_hit": bool(cached.get("ok")),
            "bounds": bounds,
            "target": [target_x, target_y],
            "guard": guard,
            "commit": {"ok": True, "key": "Enter"},
            "text_after": jab.get_text_context_value(vm_id, context),
            "description_after": info_after.description.strip() if info_after else None,
            "window": window,
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def remember_header_field(header_cache, label, result):
    if header_cache is None or not result.get("ok"):
        return
    path = result.get("path")
    window = result.get("window") or (result.get("guard") or {}).get("table_window")
    if not path:
        return
    header_cache[label] = {
        "path": path,
        "window": normalize_window_info(window),
    }


def normalize_window_info(window):
    window = dict(window or {})
    if "class_name" not in window and "class" in window:
        window["class_name"] = window.get("class")
    return window


def find_cached_header_field(jab, label, header_cache=None, scope_hwnd=None):
    cached = (header_cache or {}).get(label) or {}
    path = cached.get("path")
    window = normalize_window_info(cached.get("window"))
    if path:
        context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
            path,
            class_name=window.get("class_name") or "SunAwtCanvas",
            scope_hwnd=scope_hwnd or window.get("hwnd"),
            role="text",
            require_showing=True,
            require_valid_bounds=True,
        )
        if context:
            return {
                "ok": True,
                "context": context,
                "vm_id": vm_id,
                "owned_contexts": owned_contexts,
                "path": path,
                "window": normalize_window_info(window_info),
                "cache_hit": True,
            }
    found = find_receipt_header_form_field(jab, label, scope_hwnd=scope_hwnd)
    if found.get("ok"):
        found["cache_hit"] = False
        return found
    return found


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


def read_header_field_non_empty(jab, label, header_cache=None, scope_hwnd=None):
    found = find_cached_header_field(jab, label, header_cache, scope_hwnd)
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
            "cache_hit": bool(found.get("cache_hit")),
            "fallback_path_used": bool(found.get("fallback_path_used")),
            "text": text,
            "description": description,
            "name": name,
            "reason": None if non_empty else f"表头【{label}】为空",
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def wait_header_field_non_empty(jab, label, timeout=0.9, header_cache=None, scope_hwnd=None):
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        last = read_header_field_non_empty(jab, label, header_cache, scope_hwnd)
        if last.get("ok"):
            last["wait_seconds"] = round(timeout - max(deadline - time.time(), 0), 3)
            return last
        time.sleep(0.12)
    if last is None:
        return {
            "ok": False,
            "label": label,
            "reason": f"表头【{label}】非空检测没有取得结果",
            "wait_seconds": timeout,
        }
    last["wait_seconds"] = timeout
    return last


def read_finance_org_field(jab, header_cache=None, scope_hwnd=None):
    cached = find_cached_header_field(jab, "财务组织", header_cache, scope_hwnd)
    if cached.get("ok"):
        context = cached["context"]
        vm_id = cached["vm_id"]
        owned_contexts = cached["owned_contexts"]
        window = cached.get("window") or {}
        path = cached.get("path")
        cache_hit = True
    else:
        context, vm_id, owned_contexts, owned_indexes, window = (
            find_context_with_window(
                jab,
                "财务组织(O)",
                roles=("text",),
                timeout=1.0,
                require_showing=True,
                window_class="SunAwtCanvas",
                visible_only=True,
                scope_hwnd=scope_hwnd,
            )
        )
        path = "0" + "".join(f".{index}" for index in owned_indexes)
        cache_hit = False
    if not context:
        return {"ok": False, "label": "财务组织", "reason": "财务组织控件未找到"}
    try:
        info = jab.get_context_info(vm_id, context)
        text = str(jab.get_text_context_value(vm_id, context) or "").strip()
        description = info.description.strip() if info else ""
        name = info.name.strip() if info else ""
        non_empty = bool(text or description)
        return {
            "ok": non_empty,
            "label": "财务组织",
            "path": path,
            "cache_hit": cache_hit,
            "text": text,
            "description": description,
            "name": name,
            "window": window,
            "reason": None if non_empty else "财务组织为空",
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def wait_finance_org_field_non_empty(jab, timeout=1.2, header_cache=None, scope_hwnd=None):
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        last = read_finance_org_field(jab, header_cache, scope_hwnd)
        if last.get("ok"):
            last["wait_seconds"] = round(timeout - max(deadline - time.time(), 0), 3)
            return last
        time.sleep(0.1)
    if last is None:
        return {
            "ok": False,
            "label": "财务组织",
            "reason": "财务组织非空检测没有取得结果",
            "wait_seconds": timeout,
        }
    last["wait_seconds"] = timeout
    return last


def wait_header_field_has_text(jab, label, value, timeout=1.0, header_cache=None, scope_hwnd=None):
    expected = str(value or "").strip()
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        last = read_header_field_non_empty(jab, label, header_cache, scope_hwnd)
        text = str((last or {}).get("text") or "").strip()
        description = str((last or {}).get("description") or "").strip()
        if last and last.get("ok") and (text == expected or description == expected):
            last["wait_seconds"] = round(timeout - max(deadline - time.time(), 0), 3)
            return last
        time.sleep(0.1)
    if last is None:
        return {
            "ok": False,
            "label": label,
            "reason": f"表头【{label}】没有取得结果",
            "wait_seconds": timeout,
        }
    last["ok"] = False
    last["wait_seconds"] = timeout
    last["reason"] = f"表头【{label}】未匹配期望值 {expected!r}"
    return last


def wait_header_field_matches_any(
    jab, label, accepted_values, timeout=1.0, header_cache=None, scope_hwnd=None
):
    accepted = [
        normalize_review_value(value)
        for value in accepted_values
        if str(value or "").strip()
    ]
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        last = read_header_field_non_empty(jab, label, header_cache, scope_hwnd)
        observed = [
            normalize_review_value(value) for value in header_observed_values(last or {})
        ]
        if last and last.get("ok") and any(value in accepted for value in observed):
            last["wait_seconds"] = round(timeout - max(deadline - time.time(), 0), 3)
            return last
        time.sleep(0.1)
    if last is None:
        return {
            "ok": False,
            "label": label,
            "reason": f"表头【{label}】没有取得结果",
            "wait_seconds": timeout,
        }
    last["ok"] = False
    last["wait_seconds"] = timeout
    last["reason"] = f"表头【{label}】未匹配任一期望值 {accepted_values!r}"
    return last


def currency_review_aliases(currency_code):
    aliases = {
        "USD": ("USD", "美元"),
        "CNY": ("CNY", "人民币"),
    }
    return aliases.get(str(currency_code or "").strip().upper(), (currency_code,))


def normalize_review_value(value):
    return re.sub(r"\s+", "", str(value or "").strip()).casefold()


def header_observed_values(result):
    return [
        value
        for value in (
            result.get("text"),
            result.get("description"),
            result.get("name"),
        )
        if str(value or "").strip()
    ]


def org_review_aliases(business):
    return [
        value
        for value in (
            business.get("finance_org_code"),
            business.get("finance_org_name"),
            business.get("finance_org_short_name"),
        )
        if str(value or "").strip()
    ]


def finance_org_matches_exact(result, business):
    observed = [
        normalize_review_value(value) for value in header_observed_values(result)
    ]
    code = normalize_review_value(business.get("finance_org_code"))
    name = normalize_review_value(business.get("finance_org_name"))
    aliases = [normalize_review_value(value) for value in org_review_aliases(business)]
    if any(value in aliases for value in observed):
        return True
    # Some NC fields display "A001 上海移为通信技术股份有限公司".
    # Requiring both code and full name avoids short-name collisions such as A001 vs A00101.
    if code and name and any(code in value and name in value for value in observed):
        return True
    return False


def review_finance_org_after_fill(jab, business, header_cache=None, scope_hwnd=None):
    checks = [review_finance_org_exact(jab, business, header_cache, scope_hwnd)]
    failed_required = [
        item
        for item in checks
        if item.get("required", True) and not bool(item.get("ok"))
    ]
    return {
        "ok": not failed_required,
        "stage": "财务组织写完后",
        "checks": checks,
        "reason": None
        if not failed_required
        else "财务组织写完后回看失败："
        + "、".join(item["label"] for item in failed_required),
    }


def review_header_before_save(jab, business, header_cache=None, scope_hwnd=None):
    checks = [
        review_finance_org_exact(jab, business, header_cache, scope_hwnd),
        review_header_has_text(
            jab,
            "单据日期",
            business["document_date"],
            header_cache=header_cache,
            scope_hwnd=scope_hwnd,
        ),
        review_header_matches_any(
            jab,
            "币种",
            currency_review_aliases(business["header_currency_code"]),
            header_cache=header_cache,
            scope_hwnd=scope_hwnd,
        ),
        review_header_has_text(
            jab,
            "结算方式",
            business["settlement"],
            header_cache=header_cache,
            scope_hwnd=scope_hwnd,
        ),
        review_customer_resolved(jab, business, header_cache, scope_hwnd),
    ]
    failed_required = [
        item
        for item in checks
        if item.get("required", True) and not bool(item.get("ok"))
    ]
    return {
        "ok": not failed_required,
        "stage": "保存前",
        "checks": checks,
        "reason": None
        if not failed_required
        else "保存前表头回看失败："
        + "、".join(item["label"] for item in failed_required),
    }


def review_finance_org_exact(jab, business, header_cache=None, scope_hwnd=None):
    result = wait_finance_org_field_non_empty(
        jab, timeout=1.2, header_cache=header_cache, scope_hwnd=scope_hwnd
    )
    ok = bool(result.get("ok")) and finance_org_matches_exact(result, business)
    expected = {
        "code": business.get("finance_org_code"),
        "name": business.get("finance_org_name"),
        "short_name": business.get("finance_org_short_name"),
    }
    return {
        **result,
        "ok": ok,
        "required": True,
        "expected": expected,
        "verification_mode": "exact_or_strict_alias",
        "reason": None if ok else f"财务组织未匹配预期组织 {expected!r}",
    }


def review_customer_resolved(jab, business, header_cache=None, scope_hwnd=None):
    result = wait_header_field_non_empty(
        jab, "客户", timeout=0.9, header_cache=header_cache, scope_hwnd=scope_hwnd
    )
    ok = bool(result.get("ok"))
    expected = {
        "input_code": business.get("customer_code"),
        "payer_name": business.get("payer_name"),
    }
    return {
        **result,
        "ok": ok,
        "required": True,
        "expected": expected,
        "verification_mode": "resolved_non_empty_after_exact_code_input",
        "reason": None
        if ok
        else f"客户编码 {business.get('customer_code')!r} 输入后未解析出客户",
    }


def review_header_has_text(jab, label, value, header_cache=None, scope_hwnd=None):
    result = wait_header_field_has_text(
        jab, label, value, timeout=1.0, header_cache=header_cache, scope_hwnd=scope_hwnd
    )
    return {
        **result,
        "required": True,
        "expected": value,
        "verification_mode": "exact_text",
    }


def review_header_matches_any(jab, label, accepted_values, header_cache=None, scope_hwnd=None):
    result = wait_header_field_matches_any(
        jab,
        label,
        accepted_values,
        timeout=1.0,
        header_cache=header_cache,
        scope_hwnd=scope_hwnd,
    )
    return {
        **result,
        "required": True,
        "expected": list(accepted_values),
        "verification_mode": "any_alias",
    }


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
        "foreground_class": window_class_name(foreground),
        "foreground_title": window_text(foreground),
        "foreground_root_class": window_class_name(foreground_root),
        "foreground_root_title": window_text(foreground_root),
        "table_hwnd": table_hwnd,
        "table_root": int(table_root or 0),
        "table_class": window_class_name(table_hwnd),
        "table_title": window_text(table_hwnd),
        "table_root_class": window_class_name(table_root),
        "table_root_title": window_text(table_root),
        "reason": None
        if ok
        else "当前前台窗口不是本次定位到的 NC 收款单窗口，未发送 Ctrl+S",
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


def send_ctrl_s():
    try:
        send_virtual_key(0x11, key_up=False)
        press_virtual_key(0x53)
        send_virtual_key(0x11, key_up=True)
    except Exception:
        try:
            send_virtual_key(0x11, key_up=True)
        finally:
            raise


def press_virtual_key(vk):
    send_virtual_key(vk, key_up=False)
    time.sleep(0.03)
    send_virtual_key(vk, key_up=True)


def save_and_wait_new(jab, located):
    start = time.perf_counter()
    diagnostics = {}
    best = located.get("best") or {}
    diagnostics["prelocated_table"] = {
        "ok": bool(best),
        "seconds": 0.0,
        "row_count": best.get("row_count"),
        "col_count": best.get("col_count"),
        "window": best.get("window"),
    }
    if not best:
        return {
            "ok": False,
            "seconds": elapsed(start),
            "diagnostics": diagnostics,
            "reason": "保存前未定位到收款明细表",
        }
    guard_start = time.perf_counter()
    guard = same_nc_root_foreground(best.get("window") or {})
    guard["seconds"] = elapsed(guard_start)
    diagnostics["foreground_guard"] = guard
    if not guard.get("ok"):
        return {
            "ok": False,
            "seconds": elapsed(start),
            "diagnostics": diagnostics,
            "guard": guard,
            "reason": guard.get("reason"),
        }
    send_start = time.perf_counter()
    try:
        send_ctrl_s()
    except Exception as exc:
        diagnostics["send_ctrl_s"] = {
            "ok": False,
            "seconds": elapsed(send_start),
            "reason": f"{type(exc).__name__}: {exc}",
        }
        return {
            "ok": False,
            "seconds": elapsed(start),
            "diagnostics": diagnostics,
            "guard": guard,
            "reason": f"发送 Ctrl+S 失败：{type(exc).__name__}: {exc}",
        }
    diagnostics["send_ctrl_s"] = {"ok": True, "seconds": elapsed(send_start)}
    time.sleep(0.8)
    wait_start = time.perf_counter()
    waited = wait_new_visible(jab, timeout=SAVE_SUCCESS_TIMEOUT)
    waited["outer_seconds"] = elapsed(wait_start)
    diagnostics["wait_new_visible"] = waited
    return {
        "ok": bool(waited.get("ok")),
        "seconds": elapsed(start),
        "diagnostics": diagnostics,
        "guard": guard,
        "wait_new": waited,
        "reason": None if waited.get("ok") else waited.get("reason"),
    }


def print_save_diagnostics(save):
    if not save:
        return
    print("  保存诊断：")
    print(f"    总耗时：{save.get('seconds')} 秒")
    if save.get("reason"):
        print(f"    原因：{save.get('reason')}")
    diagnostics = save.get("diagnostics") or {}
    table = diagnostics.get("prelocated_table") or {}
    if table:
        print(
            "    保存前明细表："
            f"{'定位成功' if table.get('ok') else '定位失败'} | "
            f"row_count={table.get('row_count')} | col_count={table.get('col_count')} | "
            f"window={table.get('window')}"
        )
    guard = diagnostics.get("foreground_guard") or save.get("guard") or {}
    if guard:
        print(
            "    前台守卫："
            f"{'通过' if guard.get('ok') else '失败'} | {guard.get('seconds')} 秒"
        )
        print(
            "      foreground="
            f"{guard.get('foreground')} {guard.get('foreground_class')!r} "
            f"{guard.get('foreground_title')!r}"
        )
        print(
            "      foreground_root="
            f"{guard.get('foreground_root')} {guard.get('foreground_root_class')!r} "
            f"{guard.get('foreground_root_title')!r}"
        )
        print(
            "      table="
            f"{guard.get('table_hwnd')} {guard.get('table_class')!r} "
            f"root={guard.get('table_root')} {guard.get('table_root_class')!r} "
            f"{guard.get('table_root_title')!r}"
        )
    sent = diagnostics.get("send_ctrl_s") or {}
    if sent:
        print(
            "    发送Ctrl+S："
            f"{'成功' if sent.get('ok') else '失败'} | {sent.get('seconds')} 秒"
        )
        if sent.get("reason"):
            print(f"      原因：{sent.get('reason')}")
    waited = diagnostics.get("wait_new_visible") or save.get("wait_new") or {}
    if waited:
        print(
            "    等待新增："
            f"{'成功' if waited.get('ok') else '失败'} | "
            f"{waited.get('outer_seconds', waited.get('seconds'))} 秒 | "
            f"count={waited.get('count')}"
        )
        if waited.get("reason"):
            print(f"      原因：{waited.get('reason')}")


def print_detail_steps(title, steps):
    if not steps:
        return
    print(f"  {title}字段诊断：")
    for step in steps:
        print(
            f"    {step.get('name')}: "
            f"{'成功' if step.get('ok') else '失败' if step.get('ok') is False else '待提交'} | "
            f"col={step.get('col')} | 期望={step.get('value')!r} | "
            f"实际={step.get('actual')!r} | 输入={step.get('input_ok')} | "
            f"target={step.get('target')}"
        )
        attempts = step.get("attempts") or []
        if attempts:
            print(f"      尝试次数：{len(attempts)}")
            for attempt in attempts:
                print(
                    "      - "
                    f"#{attempt.get('attempt')} "
                    f"{'成功' if attempt.get('ok') else '失败'} | "
                    f"{attempt.get('seconds')} 秒 | "
                    f"input={attempt.get('input_ok')} | "
                    f"commit_col={attempt.get('commit_col')} | "
                    f"commit={attempt.get('commit_ok')} | "
                    f"actual={attempt.get('actual')!r}"
                )
                reason = attempt.get("input_reason") or attempt.get("commit_reason")
                if reason:
                    print(f"        原因：{reason}")
        if step.get("before") not in (None, ""):
            print(f"      写前={step.get('before')!r}")
        geometry = step.get("geometry") or {}
        if geometry:
            print(
                "      几何："
                f"bounds={geometry.get('table_bounds')} | "
                f"rows={geometry.get('row_count')} cols={geometry.get('col_count')} | "
                f"cell={geometry.get('cell_width')}x{geometry.get('cell_height')}"
            )
        commit = step.get("commit_click") or {}
        if commit:
            print(
                "      提交点击："
                f"{'成功' if commit.get('ok') else '失败'} | "
                f"target={commit.get('target')} | reason={commit.get('reason')}"
            )
        if step.get("reason"):
            print(f"      原因：{step.get('reason')}")


def print_fee_diagnostics(fee_report):
    if not fee_report:
        return
    print("  手续费诊断：")
    print(
        f"    总体：{'通过' if fee_report.get('ok') else '失败'} | "
        f"跳过={fee_report.get('skipped')} | {fee_report.get('seconds')} 秒"
    )
    if fee_report.get("reason"):
        print(f"    原因：{fee_report.get('reason')}")
    add_row = fee_report.get("add_row") or {}
    if add_row:
        print(
            "    增行："
            f"{'成功' if add_row.get('ok') else '失败'} | "
            f"before={add_row.get('before_rows')} after={add_row.get('after_rows')} | "
            f"reason={add_row.get('reason')}"
        )
        pressed = add_row.get("pressed") or {}
        if pressed:
            print(f"      热键：{pressed}")
    print_detail_steps("手续费行", fee_report.get("steps") or [])
    clear_account = fee_report.get("clear_account") or {}
    if clear_account:
        print(
            "    清手续费账户："
            f"{'成功' if clear_account.get('ok') else '失败'} | "
            f"skipped={clear_account.get('skipped')} | "
            f"before={clear_account.get('before')!r} after={clear_account.get('after')!r} | "
            f"reason={clear_account.get('reason')}"
        )
    delete_extra = fee_report.get("delete_extra") or {}
    if delete_extra:
        print(
            "    清理多余行："
            f"{'成功' if delete_extra.get('ok') else '失败'} | "
            f"skipped={delete_extra.get('skipped')} | "
            f"before={delete_extra.get('before_rows')} after={delete_extra.get('after_rows')} | "
            f"reason={delete_extra.get('reason')}"
        )
        cleanup = delete_extra.get("cleanup") or delete_extra
        for step in cleanup.get("steps") or []:
            print(
                "      - "
                f"删除目标行={step.get('target_row')} | "
                f"{step.get('before_rows')} -> {step.get('after_rows')} | "
                f"{'成功' if step.get('ok') else '失败'} | "
                f"reason={step.get('reason')}"
            )


def print_case_summary(case_report):
    print()
    print(f"案例：{case_report.get('name')}")
    business = case_report.get("business") or {}
    if business:
        print(
            f"  Excel行：{case_report.get('excel_row')} | "
            f"客户={business.get('customer_code')} | 金额={business.get('amount')} | "
            f"手续费={business.get('fee')} | 账号={business.get('bank_account')}"
        )
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
                f"| {step.get('seconds')} 秒 | path={step.get('path')}"
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
            commit_check = step.get("commit_check") or {}
            if commit_check:
                print(f"      原地检测：跳过 | {commit_check.get('reason')}")
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
    header_reviews = case_report.get("header_reviews") or []
    if header_reviews:
        print("  表头回看：")
        for review in header_reviews:
            print(f"    阶段：{review.get('stage')}")
            for item in review.get("checks") or []:
                print(
                    f"      {item.get('label')}: "
                    f"{'通过' if item.get('ok') else '失败'} | "
                    f"mode={item.get('verification_mode')} | "
                    f"expected={item.get('expected')!r} | "
                    f"text={item.get('text')!r}, description={item.get('description')!r}"
                )
                if item.get("reason"):
                    print(f"        原因：{item.get('reason')}")
    print_detail_steps("主行", case_report.get("detail_steps") or [])
    print_fee_diagnostics(case_report.get("fee"))
    if case_report.get("ok"):
        if case_report.get("save_skipped"):
            print("  结果：填写诊断通过，按本轮要求未保存。")
        else:
            print("  结果：成功，保存后已看到【新增】。")
    else:
        print(f"  结果：失败，停止位置={case_report.get('failed_step')}")
        reason = case_report.get("reason")
        if reason:
            print(f"  原因：{reason}")
    print_save_diagnostics(case_report.get("save"))


def run_one_case(config, case, allow_existing_entry=False):
    case_start = time.perf_counter()
    case_report = {"name": case.name, "excel_row": case.excel_row, "timings": []}
    jab = JABOperator(config)
    try:
        start = time.perf_counter()
        jab.ensure_started()
        case_report["timings"].append(
            {"name": "起始检测.JAB启动", "seconds": elapsed(start)}
        )
        start = time.perf_counter()
        start_state = ensure_starts_from_new_state(jab)
        case_report["timings"].append(
            {"name": "起始检测.新增入口/自制态", "seconds": elapsed(start)}
        )
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

        if start_state.get("state") == "existing-self-made-entry":
            opened = {
                "ok": True,
                "seconds": 0.0,
                "method": "reuse-existing-self-made-entry",
            }
        else:
            opened = run_receipt_new_probe(
                jab=jab,
                before=start_state.get("windows"),
                buttons=start_state.get("buttons"),
            )
        case_report["open_self_made"] = opened
        case_report["timings"].append(
            {"name": "新增->自制", "seconds": opened["seconds"]}
        )
        for item in opened.get("timings") or []:
            case_report["timings"].append(
                {
                    "name": f"新增->自制.{item.get('name')}",
                    "seconds": item.get("seconds"),
                }
            )
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

        scope_start = time.perf_counter()
        header_scope = locate_receipt_header_scope(jab)
        case_report["header_scope"] = header_scope
        case_report["timings"].append(
            {"name": "表头scope定位", "seconds": elapsed(scope_start)}
        )
        if not header_scope.get("ok"):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "locate-header-scope",
                    "reason": header_scope.get("reason"),
                    "seconds": elapsed(case_start),
                }
            )
            return case_report
        scope_hwnd = header_scope["scope_hwnd"]

        case_report["timings"].append(
            {"name": "填写阶段.JAB启动", "seconds": 0.0, "reused": True}
        )
        business = business_for_case(config, case)
        case_report["business"] = business
        case_report["customer_checks"] = []
        case_report["header_reviews"] = []
        header_cache = {}

        header_start = time.perf_counter()
        header_steps = fill_minimal_header(jab, business, header_cache, scope_hwnd)
        case_report["header_steps"] = header_steps
        case_report["header_cache"] = header_cache
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

        start = time.perf_counter()
        finance_org_after_fill = review_finance_org_after_fill(
            jab, business, header_cache, scope_hwnd
        )
        case_report["header_reviews"].append(finance_org_after_fill)
        case_report["timings"].append(
            {"name": "财务组织回看.写完后", "seconds": elapsed(start)}
        )
        if not finance_org_after_fill.get("ok"):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "review-finance-org-after-fill",
                    "reason": finance_org_after_fill.get("reason"),
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
        for item in (
            ((fee_report.get("delete_extra") or {}).get("timings") or [])
            if isinstance(fee_report, dict)
            else []
        ):
            case_report["timings"].append(
                {
                    "name": f"手续费.{item.get('name')}",
                    "seconds": item.get("seconds"),
                }
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

        start = time.perf_counter()
        header_review = review_header_before_save(
            jab, business, header_cache, scope_hwnd
        )
        case_report["header_reviews"].append(header_review)
        case_report["header_review"] = header_review
        case_report["timings"].append(
            {"name": "保存前表头回看", "seconds": elapsed(start)}
        )
        if not header_review.get("ok"):
            case_report.update(
                {
                    "ok": False,
                    "failed_step": "review-header-before-save",
                    "reason": header_review.get("reason"),
                    "seconds": elapsed(case_start),
                }
            )
            return case_report

        if not SAVE_ENABLED:
            if READ_DETAIL_BEFORE_SAVE_SKIPPED:
                start = time.perf_counter()
                after_table = read_body_table(jab, "before_save_skipped")
                case_report["timings"].append(
                    {"name": "保存前停止.读明细", "seconds": elapsed(start)}
                )
            else:
                after_table = {
                    "ok": True,
                    "skipped": True,
                    "reason": "速度测试跳过 no-save 保存前明细诊断读表",
                }
                case_report["timings"].append(
                    {"name": "保存前停止.读明细", "seconds": 0.0}
                )
            case_report.update(
                {
                    "ok": True,
                    "save_skipped": True,
                    "failed_step": None,
                    "reason": "本轮配置为不保存，已在保存前停止",
                    "after_table": after_table,
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
    run_start = time.perf_counter()
    report = {
        "launcher": "tmp_receipt_two_case_save_run.py",
        "start_delay_seconds": START_DELAY_SECONDS,
        "stop_hotkey": STOP_HOTKEY,
        "cases": [],
        "timings": [],
    }
    try:
        start = time.perf_counter()
        config = load_config(str(ROOT / "config.json"))
        report["timings"].append({"name": "全局.读取配置", "seconds": elapsed(start)})
        start = time.perf_counter()
        test_cases = build_latest_test_cases(config)
        report["timings"].append(
            {"name": "全局.读取Excel并选择测试行", "seconds": elapsed(start)}
        )
        report["selected_cases"] = [case.__dict__ for case in test_cases]
        print_header(test_cases)
        print()
        print(
            f"请在 {START_DELAY_SECONDS} 秒内切到 NC【收款单录入】且能看到【新增】的页面..."
        )
        start = time.perf_counter()
        time.sleep(START_DELAY_SECONDS)
        report["timings"].append({"name": "全局.启动等待", "seconds": elapsed(start)})
        print("开始真实保存循环测试。")

        if is_stop_hotkey_pressed():
            print(f"检测到紧急停止键 {STOP_HOTKEY}，未开始。")
            return 1
        health_jab = JABOperator(config)
        try:
            start = time.perf_counter()
            health_jab.ensure_started()
            report["timings"].append(
                {"name": "全局.健康检查JAB启动", "seconds": elapsed(start)}
            )
            start = time.perf_counter()
            health = check_jab_ready(health_jab)
            report["timings"].append(
                {"name": "全局.JAB健康检查", "seconds": elapsed(start)}
            )
            report["jab_health"] = health
            if not health.get("ok"):
                print("JAB 启动状态：")
                print_jab_health_failure(health)
                return 1
        finally:
            health_jab.close()

        for index, case in enumerate(test_cases):
            if is_stop_hotkey_pressed():
                print(f"检测到紧急停止键 {STOP_HOTKEY}，停止后续案例。")
                break
            case_report = run_one_case(
                config,
                case,
                allow_existing_entry=(
                    index == 0 and ALLOW_EXISTING_ENTRY_FOR_FIRST_CASE
                ),
            )
            report["cases"].append(case_report)
            print_case_summary(case_report)
            if case_report.get("save_skipped"):
                break
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

    if SAVE_ENABLED:
        ok = len(report["cases"]) == len(report.get("selected_cases", [])) and all(
            item.get("ok") for item in report["cases"]
        )
    else:
        ok = bool(report["cases"]) and all(item.get("ok") for item in report["cases"])
    print()
    print("总结果：")
    print(f"  {'成功' if ok else '失败'}")
    total = sum(float(item.get("seconds") or 0) for item in report["cases"])
    print(f"  两案例累计耗时：{round(total, 3)} 秒")
    print(f"  脚本总耗时：{elapsed(run_start)} 秒")
    if report.get("timings"):
        print("  全局计时：")
        for item in report["timings"]:
            print(f"  - {item['name']}: {item['seconds']} 秒")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
