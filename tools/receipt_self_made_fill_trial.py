import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime
from decimal import Decimal
import json
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import ReceiptEntryConfig  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.read_receipt_excel_row import DEFAULT_FIELDS  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    collect_receipt_new_windows,
    detect_self_made_entry_state,
)
from tools.receipt_account_reference_try import (  # noqa: E402
    OK_PATH as ACCOUNT_REF_OK_PATH,
    RESULT_TABLE_PATH as ACCOUNT_REF_RESULT_TABLE_PATH,
    wait_table,
)
from tools.receipt_table_cell_probe import select_cell  # noqa: E402


CURRENCY_NAMES = {"USD": "美元", "RMB": "人民币", "CNY": "人民币"}
HEADER_FORM_BASE_PATH = "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0"
HEADER_DYNAMIC_PREFIX_BASE = "0.0.1.0.0.0.0"
HEADER_DYNAMIC_MAX_INDEX = 8
HEADER_COMMON_SUFFIX_TEMPLATE = (
    "0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.{index}.0"
)
HEADER_COMMON_LABEL_SUFFIX_TEMPLATE = (
    "0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.{index}"
)
FINANCE_ORG_SUFFIX = "0.0.0.1.1.0.0.0.0.1.1.1.0"
FINANCE_ORG_LABEL_SUFFIX = "0.0.0.1.1.0.0.0.0.1.1.0"
ACCOUNT_REFERENCE_BUTTON_PATH = f"{HEADER_FORM_BASE_PATH}.15.1"
ACCOUNT_REF_SEARCH_TEXT_PATH = "0.0.1.0.0.0.1.0.0.0.0"
HEADER_FORM_TEXT_INDEXES = {
    "单据日期": 5,
    "币种": 13,
    "收款银行账户": 15,
    "客户": 17,
    "结算方式": 31,
}
HEADER_DYNAMIC_PROBE_LABELS = ("单据日期", "币种", "客户")
HEADER_LABEL_ALIASES = {
    "财务组织": ("财务组织", "财务组织(O)"),
    "客户": ("客户",),
    "单据日期": ("单据日期",),
    "币种": ("币种",),
}
FAST_HEADER_SCOPE_LABEL = "财务组织"


def print_json(data):
    text = json.dumps(data, ensure_ascii=False, indent=2, default=str)
    try:
        sys.stdout.write(text + "\n")
    except UnicodeEncodeError:
        sys.stdout.buffer.write((text + "\n").encode("utf-8"))


def main():
    parser = argparse.ArgumentParser(
        description="Trial-fill NC receipt self-made entry from one Excel row."
    )
    parser.add_argument("row", type=int)
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--open-self-made", action="store_true")
    parser.add_argument(
        "--stop-before-account-reference",
        action="store_true",
        help="fill verified header fields only, then stop before opening/searching account reference",
    )
    parser.add_argument(
        "--continue-account-reference",
        action="store_true",
        help="continue past opening the account reference; reserved for explicitly staged trials",
    )
    parser.add_argument(
        "--fill-detail",
        action="store_true",
        help="fill receipt detail cells after header verification",
    )
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    config = load_config(args.config)
    row_data = read_excel_row(config, args.row)
    business = build_business_values(config, row_data)
    report = {"row": args.row, "excel": row_data, "business": business, "steps": []}

    if args.open_self_made:
        open_step = detect_existing_self_made_entry(config)
        if not open_step.get("ok"):
            open_step = run_receipt_new_probe()
        report["steps"].append(open_step)
        if not open_step.get("ok"):
            report["stopped"] = "open_self_made"
            print_json(report)
            return 1

    jab = JABOperator(config)
    try:
        jab.ensure_started()
        header_steps = fill_header(
            jab,
            business,
            stop_before_account_reference=args.stop_before_account_reference,
            continue_account_reference=args.continue_account_reference,
        )
        report["steps"].extend(header_steps)
        if any(step.get("step") == "blocked" for step in header_steps):
            report["stopped"] = "header"
            print_json(report)
            return 1
        report["steps"].append(read_body_table(jab, "before_detail_fill"))
        if args.fill_detail:
            report["steps"].extend(fill_detail_line(jab, business))
            report["steps"].append(read_body_table(jab, "after_detail_fill"))
        else:
            report["steps"].append(
                {
                    "step": "detail_fill",
                    "ok": False,
                    "blocked": True,
                    "reason": "detail fill requires explicit --fill-detail",
                }
            )
    finally:
        jab.close()

    print_json(report)
    return 0


def read_excel_row(config, row):
    import openpyxl

    excel_cfg = config["receipt_entry"]["excel"]
    workbook = openpyxl.load_workbook(excel_cfg["path"], data_only=True, read_only=True)
    try:
        sheet = workbook[excel_cfg["sheet_name"]]
        header_row = excel_cfg.get("header_row", 1)
        headers = [
            sheet.cell(header_row, col).value for col in range(1, sheet.max_column + 1)
        ]
        data = {
            headers[col - 1]: sheet.cell(row, col).value
            for col in range(1, sheet.max_column + 1)
            if headers[col - 1]
        }
    finally:
        workbook.close()
    return {field: data.get(field) for field in DEFAULT_FIELDS}


def build_business_values(config, row_data):
    receipt_config = ReceiptEntryConfig(config)
    bank = str(row_data.get("银行") or "").strip()
    account = receipt_config.accounts_by_label.get(
        bank.upper()
    ) or receipt_config.accounts_by_label.get(bank)
    if not account:
        normalized_bank = "".join(
            ch for ch in bank.upper() if ch.isalnum() or "\u4e00" <= ch <= "\u9fff"
        )
        account = receipt_config.accounts_by_label.get(normalized_bank)
    if not account:
        raise SystemExit(f"bank account config not found: {bank!r}")
    organization = receipt_config.organizations[account.organization_code]
    currency_code = str(row_data.get("币种") or "").strip().upper()
    receipt_date = row_data.get("到款日期")
    if isinstance(receipt_date, datetime):
        receipt_date_text = receipt_date.strftime("%Y-%m-%d")
    else:
        receipt_date_text = str(receipt_date)[:10]
    amount = Decimal(str(row_data.get("🟪到账金额") or "0"))
    fee_raw = row_data.get("手续费")
    fee = Decimal(str(fee_raw or "0"))
    return {
        "finance_org_code": organization.code,
        "finance_org_name": organization.name,
        "document_date": receipt_date_text,
        "customer_code": str(row_data.get("客户编码") or "").strip(),
        "currency": CURRENCY_NAMES.get(currency_code, currency_code),
        "header_currency_code": account.header_currency_code,
        "bank_label": bank,
        "bank_account": account.account_no,
        "amount": str(amount),
        "fee": str(fee),
        "has_fee": fee != 0,
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
        "fee_subject": "660305",
        "fee_business_type": "手续费",
    }


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
    proc = subprocess.run(
        cmd,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    parsed = None
    ok = False
    try:
        parsed = json.loads(proc.stdout)
        entry_state = parsed.get("entry_state") or {}
        ok = (
            proc.returncode == 0
            and bool((parsed.get("open") or {}).get("ok"))
            and bool((parsed.get("choose_self_made") or {}).get("ok"))
            and bool(entry_state.get("ok"))
        )
    except json.JSONDecodeError:
        ok = False
    return {
        "step": "open_self_made",
        "ok": ok,
        "returncode": proc.returncode,
        "parsed": parsed,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
    }


def detect_existing_self_made_entry(config):
    jab = JABOperator(config)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        windows = collect_receipt_new_windows(jab)
        entry_state = detect_self_made_entry_state(windows)
        java_windows = [window for window in windows if window.get("is_java")]
        return {
            "step": "open_self_made",
            "ok": bool(entry_state.get("ok")),
            "method": "existing_entry_state_probe",
            "entry_state": entry_state,
            "java_window_count": len(java_windows),
            "reason": (
                "already in self-made entry state"
                if entry_state.get("ok")
                else "self-made entry state not detected before opening"
            ),
        }
    finally:
        jab.hide_blank_awt_windows_enabled = False
        jab.close()


def fill_header(
    jab,
    business,
    stop_before_account_reference=False,
    continue_account_reference=False,
):
    steps = []
    for field in [
        {
            "label": "财务组织",
            "value": business["finance_org_code"],
            "control_name": "财务组织(O)",
        },
        {
            "label": "客户",
            "value": business["customer_code"],
            "header_form": True,
        },
        {
            "label": "单据日期",
            "value": business["document_date"],
            "header_form": True,
        },
        {
            "label": "币种",
            "value": business["header_currency_code"],
            "header_form": True,
        },
        {
            "label": "收款银行账户",
            "value": business["bank_account"],
            "header_account_reference": True,
        },
    ]:
        label = field["label"]
        value = field["value"]
        if field.get("header_account_reference") and stop_before_account_reference:
            steps.append(
                {
                    "step": "blocked",
                    "reason": "stopped before account reference by request",
                    "label": label,
                    "value": value,
                }
            )
            break
        if field.get("control_name"):
            result = set_text_by_control_name(
                jab,
                field["control_name"],
                value,
                commit_key=field.get("commit_key"),
                accepted_text=field.get("accepted_text"),
                unlock_query=field.get("unlock_query"),
            )
            ok = bool(result.get("ok"))
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "control_name",
                    "control_name": field["control_name"],
                    **result,
                }
            )
        elif field.get("header_account_reference"):
            result = set_header_account_by_reference(
                jab,
                value,
                continue_after_open=continue_account_reference,
            )
            ok = bool(result.get("ok"))
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "account_reference",
                    **result,
                }
            )
        elif field.get("header_form"):
            result = set_receipt_header_form_field(
                jab,
                label,
                value,
                commit_key=field.get("commit_key"),
            )
            ok = bool(result.get("ok"))
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "header_form_label",
                    **result,
                }
            )
        else:
            ok = jab.set_text_near_label(
                label,
                value,
                class_name="SunAwtCanvas",
                require_showing=True,
                timeout=2.0,
                wait=0.2,
            )
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "near_label",
                    "ok": bool(ok),
                }
            )
        if not ok:
            steps.append(
                {
                    "step": "blocked",
                    "reason": "header field failed; blocking workflow stopped",
                    "label": label,
                }
            )
            break
    return steps


def set_header_account_by_reference(jab, account, continue_after_open=False):
    existing_dialog = wait_reference_dialog(jab, timeout=0.4)
    if existing_dialog and not continue_after_open:
        return {
            "ok": False,
            "blocked": True,
            "reason": "account reference already open; search/select must be run as separate staged actions",
            "button_path": ACCOUNT_REFERENCE_BUTTON_PATH,
            "dialog": existing_dialog,
            "next_required": "foreground_check_account_reference",
        }

    opened = jab.do_action_by_path(
        ACCOUNT_REFERENCE_BUTTON_PATH,
        class_name="SunAwtCanvas",
        action_name="单击",
        wait=0.8,
        timeout=3.0,
        require_showing=True,
        require_valid_bounds=False,
    )
    if not opened:
        return {
            "ok": False,
            "reason": "account reference button action failed",
            "button_path": ACCOUNT_REFERENCE_BUTTON_PATH,
        }

    dialog = wait_reference_dialog(jab, timeout=6.0)
    if not dialog:
        return {
            "ok": False,
            "reason": "使用权参照 dialog not found after account button",
            "reference_windows": collect_reference_window_candidates(jab),
        }

    if not continue_after_open:
        return {
            "ok": False,
            "blocked": True,
            "reason": "account reference opened; search/select must be run as separate staged actions",
            "button_path": ACCOUNT_REFERENCE_BUTTON_PATH,
            "dialog": dialog,
            "next_required": "foreground_check_account_reference",
        }

    search = set_reference_search_text(jab, dialog["hwnd"], account)
    table = wait_table(jab, dialog["hwnd"], timeout=30.0)
    if not table.get("ok") or table.get("row_count", 0) <= 0:
        return {
            "ok": False,
            "reason": "account reference search returned no rows",
            "dialog": dialog,
            "search": search,
            "table": table,
        }

    selected = select_reference_result_first_row(jab, dialog["hwnd"])
    if not selected.get("ok"):
        return {
            "ok": False,
            "reason": "account reference first row selection failed",
            "dialog": dialog,
            "search": search,
            "table": table,
            "selection": selected,
        }

    confirmed = jab.do_action_by_path(
        ACCOUNT_REF_OK_PATH,
        scope_hwnd=dialog["hwnd"],
        action_name="单击",
        wait=1.2,
        timeout=3.0,
        require_showing=True,
        require_valid_bounds=False,
    )
    verified = wait_header_account_description(jab, timeout=5.0)
    return {
        "ok": bool(confirmed and verified.get("accepted")),
        "button_path": ACCOUNT_REFERENCE_BUTTON_PATH,
        "dialog": dialog,
        "search": search,
        "table": table,
        "selection": selected,
        "confirmed": bool(confirmed),
        "verified": verified,
    }


def wait_reference_dialog(jab, timeout=6.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for hwnd, title, class_name, pid, visible in jab.get_scoped_windows(
            None, include_children=True
        ):
            if (
                visible
                and title == "使用权参照"
                and class_name == "SunAwtDialog"
                and jab.dll.isJavaWindow(hwnd)
            ):
                return {
                    "hwnd": int(hwnd),
                    "title": title,
                    "class_name": class_name,
                    "pid": pid,
                    "visible": visible,
                }
        time.sleep(0.2)
    return None


def collect_reference_window_candidates(jab):
    candidates = []
    for hwnd, title, class_name, pid, visible in jab.get_scoped_windows(
        None, include_children=True
    ):
        if not (
            "参照" in title
            or class_name.startswith("SunAwt")
            or class_name.startswith("Yonyou")
        ):
            continue
        candidates.append(
            {
                "hwnd": int(hwnd),
                "title": title,
                "class_name": class_name,
                "pid": pid,
                "visible": visible,
                "is_java": bool(jab.dll.isJavaWindow(hwnd)),
            }
        )
    return candidates[:50]


def set_reference_search_text(jab, hwnd, account):
    context, vm_id, owned, window = jab.find_context_by_path_once(
        ACCOUNT_REF_SEARCH_TEXT_PATH,
        scope_hwnd=hwnd,
        role="text",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "reason": "reference search text field not found",
            "path": ACCOUNT_REF_SEARCH_TEXT_PATH,
        }
    try:
        before = jab.get_text_context_value(vm_id, context)
        info_before = jab.get_context_info(vm_id, context)
        write_ok = jab.set_text_context(vm_id, context, account)
        time.sleep(0.3)
        enter_ok = post_key_to_hwnd(hwnd, "enter")
        time.sleep(1.0)
        after = jab.get_text_context_value(vm_id, context)
        info_after = jab.get_context_info(vm_id, context)
        return {
            "ok": bool(write_ok and enter_ok),
            "method": "jab_text_path_post_enter",
            "path": ACCOUNT_REF_SEARCH_TEXT_PATH,
            "window": window,
            "text_before": before,
            "description_before": (
                info_before.description.strip() if info_before else None
            ),
            "text_after": after,
            "description_after": (
                info_after.description.strip() if info_after else None
            ),
            "enter_ok": bool(enter_ok),
        }
    finally:
        jab.release_contexts(vm_id, owned)


def select_reference_result_first_row(jab, hwnd):
    context, vm_id, owned, window = jab.find_context_by_path_once(
        ACCOUNT_REF_RESULT_TABLE_PATH,
        scope_hwnd=hwnd,
        role="table",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "reason": "result table not found",
            "path": ACCOUNT_REF_RESULT_TABLE_PATH,
        }
    try:
        info = jab.get_table_info(vm_id, context)
        if not info:
            return {"ok": False, "reason": "table info unavailable", "window": window}
        if info.rowCount <= 0:
            return {
                "ok": False,
                "reason": "no rows",
                "row_count": info.rowCount,
                "col_count": info.columnCount,
            }
        before = jab.get_selected_child_indexes(
            vm_id, context, info.rowCount * info.columnCount
        )
        jab.dll.clearAccessibleSelectionFromContext(vm_id, context)
        jab.dll.addAccessibleSelectionFromContext(vm_id, context, 0)
        time.sleep(0.3)
        after = jab.get_selected_child_indexes(
            vm_id, context, info.rowCount * info.columnCount
        )
        cells = {
            str(col): jab.get_table_cell_text(vm_id, context, 0, col)
            for col in range(min(info.columnCount, 10))
        }
        return {
            "ok": 0 in after,
            "path": ACCOUNT_REF_RESULT_TABLE_PATH,
            "window": window,
            "row_count": info.rowCount,
            "col_count": info.columnCount,
            "selected_before": before,
            "selected_after": after,
            "first_row_cells": cells,
        }
    finally:
        jab.release_contexts(vm_id, owned)


def wait_header_account_description(jab, timeout=5.0):
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        found = find_receipt_header_form_field(jab, "收款银行账户")
        if found.get("ok"):
            context = found["context"]
            vm_id = found["vm_id"]
            owned_contexts = found["owned_contexts"]
            try:
                info = jab.get_context_info(vm_id, context)
                text = jab.get_text_context_value(vm_id, context)
                desc = info.description.strip() if info else ""
                last = {"text": text, "description": desc, "path": found.get("path")}
                if text or desc:
                    last["accepted"] = True
                    return last
            finally:
                jab.release_contexts(vm_id, owned_contexts)
        time.sleep(0.3)
    if last is None:
        last = {"text": None, "description": "", "accepted": False}
    else:
        last["accepted"] = False
    return last


def set_receipt_header_form_field(jab, label, value, commit_key=None):
    found = find_receipt_header_form_field(jab, label)
    if not found.get("ok"):
        return found
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    window_info = found["window"]
    try:
        info_before = jab.get_context_info(vm_id, context)
        before = jab.get_text_context_value(vm_id, context)
        ok = jab.set_text_context(vm_id, context, value)
        time.sleep(0.3)
        commit = None
        if ok and commit_key:
            if hasattr(jab.dll, "requestFocus"):
                jab.dll.requestFocus(vm_id, context)
            action_result = do_context_commit_action(jab, vm_id, context)
            post_key_ok = False
            if not action_result.get("ok"):
                post_key_ok = post_key_to_hwnd(window_info.get("hwnd"), commit_key)
            backend_state = wait_backend_field_state(
                jab,
                vm_id,
                context,
                value=value,
                timeout=3.0,
            )
            commit = {
                "key": commit_key,
                "action": action_result,
                "post_key_ok": bool(post_key_ok),
                "backend_state": backend_state,
            }
        backend_state = wait_backend_field_state(
            jab,
            vm_id,
            context,
            value=value,
            timeout=1.0,
        )
        info_after = jab.get_context_info(vm_id, context)
        after = jab.get_text_context_value(vm_id, context)
        desc_after = info_after.description.strip() if info_after else ""
        return {
            "ok": bool(ok and backend_state.get("accepted")),
            "path": found["path"],
            "label_path": found["label_path"],
            "fallback_path_used": bool(found.get("fallback_path_used")),
            "text_before": before,
            "description_before": (
                info_before.description.strip() if info_before else None
            ),
            "text_after": after,
            "description_after": desc_after,
            "backend_state": backend_state,
            "commit": commit,
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def find_receipt_header_form_field(jab, label, scope_hwnd=None):
    # 收款单页签路径会随已打开 NC 页面数量漂移；表头字段优先按 label
    # 在当前可见 SunAwtCanvas 中语义定位，固定 path 只作为现场兜底。
    aliases = HEADER_LABEL_ALIASES.get(label, (label,))
    for window in header_scope_windows(jab, scope_hwnd):
        hwnd, title, class_name, pid, visible = window
        if (
            not visible
            or class_name != "SunAwtCanvas"
            or not jab.dll.isJavaWindow(hwnd)
        ):
            continue
        from tools.jab_probe import JOBJECT

        vm_id_ref = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id_ref),
            ctypes.byref(root_context),
        ):
            continue
        result = None
        matched_alias = None
        for alias in aliases:
            result = find_label_following_text(
                jab,
                vm_id_ref.value,
                root_context.value,
                alias,
                path="0",
                depth=0,
                owned_contexts=[root_context.value],
            )
            if result:
                matched_alias = alias
                break
        if result:
            context, owned_contexts, path, label_path = result
            return {
                "ok": True,
                "matched_alias": matched_alias,
                "context": context,
                "vm_id": vm_id_ref.value,
                "owned_contexts": owned_contexts,
                "path": path,
                "label_path": label_path,
                "window": {
                    "hwnd": int(hwnd),
                    "title": title,
                    "class_name": class_name,
                    "pid": pid,
                    "visible": visible,
                },
                "dynamic_index": extract_receipt_header_dynamic_index(path),
                "dynamic_prefix": (
                    receipt_header_dynamic_prefix(
                        extract_receipt_header_dynamic_index(path)
                    )
                    if extract_receipt_header_dynamic_index(path) is not None
                    else None
                ),
        }
        jab.release_contexts(vm_id_ref.value, [root_context.value])
    fallback = find_receipt_header_form_field_by_path(jab, label, scope_hwnd=scope_hwnd)
    if fallback.get("ok"):
        fallback["fallback_path_used"] = True
        return fallback
    return {"ok": False, "reason": "header form field not found", "label": label}


def header_scope_windows(jab, scope_hwnd=None):
    if scope_hwnd is None or os.name != "nt" or not hasattr(ctypes, "windll"):
        return jab.get_scoped_windows(scope_hwnd, include_children=True)
    item = describe_hwnd_for_scope(scope_hwnd)
    if not item:
        return []
    return [item]


def build_receipt_header_dynamic_path(dynamic_index, label):
    if label == "财务组织":
        return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{FINANCE_ORG_SUFFIX}"
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    if index is None:
        return None
    suffix = HEADER_COMMON_SUFFIX_TEMPLATE.format(index=index)
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{suffix}"


def build_receipt_header_dynamic_label_path(dynamic_index, label):
    if label == "财务组织":
        return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{FINANCE_ORG_LABEL_SUFFIX}"
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    if index is None:
        return None
    suffix = HEADER_COMMON_LABEL_SUFFIX_TEMPLATE.format(index=index - 1)
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{suffix}"


def receipt_header_dynamic_prefix(dynamic_index):
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}"


def extract_receipt_header_dynamic_index(path):
    prefix = f"{HEADER_DYNAMIC_PREFIX_BASE}."
    if not path or not path.startswith(prefix):
        return None
    first = path[len(prefix) :].split(".", 1)[0]
    try:
        return int(first)
    except ValueError:
        return None


def find_receipt_header_field_by_dynamic_path(
    jab,
    label,
    dynamic_index,
    scope_hwnd=None,
    require_showing=True,
    require_valid_bounds=True,
):
    text_path = build_receipt_header_dynamic_path(dynamic_index, label)
    if not text_path:
        return {
            "ok": False,
            "reason": "header dynamic path not configured",
            "label": label,
            "dynamic_index": dynamic_index,
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        text_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=require_showing,
        require_valid_bounds=require_valid_bounds,
    )
    if not context:
        return {
            "ok": False,
            "reason": "header dynamic path not found",
            "label": label,
            "path": text_path,
            "dynamic_index": dynamic_index,
        }
    return {
        "ok": True,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": build_receipt_header_dynamic_label_path(dynamic_index, label),
        "window": window_info,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
    }


def infer_receipt_header_dynamic_prefix(
    jab,
    scope_hwnd=None,
    dynamic_max=HEADER_DYNAMIC_MAX_INDEX,
    require_showing=True,
    require_valid_bounds=True,
):
    attempts = []
    for dynamic_index in range(dynamic_max + 1):
        ok_labels = []
        label_results = {}
        for label in HEADER_DYNAMIC_PROBE_LABELS:
            found = find_receipt_header_field_by_dynamic_path(
                jab,
                label,
                dynamic_index,
                scope_hwnd=scope_hwnd,
                require_showing=require_showing,
                require_valid_bounds=require_valid_bounds,
            )
            label_results[label] = {
                "ok": bool(found.get("ok")),
                "path": found.get("path"),
                "reason": found.get("reason"),
            }
            if found.get("ok"):
                ok_labels.append(label)
                jab.release_contexts(found["vm_id"], found["owned_contexts"])
        if ok_labels:
            attempts.append(
                {
                    "dynamic_index": dynamic_index,
                    "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                    "ok_labels": ok_labels,
                    "ok_count": len(ok_labels),
                    "labels": label_results,
                }
            )
        if len(ok_labels) == len(HEADER_DYNAMIC_PROBE_LABELS):
            return {
                "ok": True,
                "dynamic_index": dynamic_index,
                "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
                "matched_labels": ok_labels,
                "attempts": attempts,
            }
    return {
        "ok": False,
        "reason": "header dynamic prefix not found",
        "attempts": attempts,
    }


def infer_receipt_header_scope_fast(jab):
    foreground_root = foreground_root_hwnd()
    windows = jab.get_scoped_windows(None, include_children=True)
    candidates = []
    for window in windows:
        hwnd, _title, class_name, _pid, visible = window
        if not visible or class_name != "SunAwtCanvas" or not jab.dll.isJavaWindow(hwnd):
            continue
        if foreground_root and window_root_hwnd(hwnd) != foreground_root:
            continue
        prefix = infer_receipt_header_dynamic_prefix(
            jab,
            scope_hwnd=hwnd,
            require_showing=False,
            require_valid_bounds=False,
        )
        if not prefix.get("ok"):
            continue
        finance_org = find_receipt_header_field_by_dynamic_path(
            jab,
            FAST_HEADER_SCOPE_LABEL,
            prefix["dynamic_index"],
            scope_hwnd=hwnd,
            require_showing=False,
            require_valid_bounds=False,
        )
        if not finance_org.get("ok"):
            continue
        jab.release_contexts(finance_org["vm_id"], finance_org["owned_contexts"])
        candidates.append(
            {
                "hwnd": int(hwnd),
                "path": finance_org.get("path"),
                "dynamic_index": prefix["dynamic_index"],
                "dynamic_prefix": prefix["dynamic_prefix"],
                "matched_labels": prefix.get("matched_labels"),
            }
        )
    if len(candidates) == 1:
        return {
            "ok": True,
            "scope_hwnd": candidates[0]["hwnd"],
            "mode": "fast-path",
            "dynamic_index": candidates[0]["dynamic_index"],
            "dynamic_prefix": candidates[0]["dynamic_prefix"],
            "candidates": candidates,
        }
    return {
        "ok": False,
        "mode": "fast-path",
        "candidates": candidates,
        "reason": f"快速表头 scope 候选数量不是 1：{len(candidates)}",
    }


def describe_hwnd_for_scope(hwnd):
    user32 = ctypes.windll.user32
    hwnd_obj = wintypes.HWND(int(hwnd))
    if not user32.IsWindow(hwnd_obj):
        return None
    title_len = user32.GetWindowTextLengthW(hwnd_obj)
    title_buffer = ctypes.create_unicode_buffer(title_len + 1)
    user32.GetWindowTextW(hwnd_obj, title_buffer, title_len + 1)
    class_buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd_obj, class_buffer, 256)
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd_obj, ctypes.byref(pid))
    return (
        int(hwnd),
        title_buffer.value,
        class_buffer.value,
        int(pid.value),
        bool(user32.IsWindowVisible(hwnd_obj)),
    )


def find_receipt_header_form_field_by_path(
    jab,
    label,
    scope_hwnd=None,
    dynamic_index=None,
):
    if dynamic_index is not None:
        return find_receipt_header_field_by_dynamic_path(
            jab,
            label,
            dynamic_index,
            scope_hwnd=scope_hwnd,
        )
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    if index is None:
        return {
            "ok": False,
            "reason": "header form path not configured",
            "label": label,
        }
    text_path = f"{HEADER_FORM_BASE_PATH}.{index}.0"
    label_path = f"{HEADER_FORM_BASE_PATH}.{index - 1}"
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        text_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=True,
        require_valid_bounds=True,
    )
    if not context:
        return {
            "ok": False,
            "reason": "header form path not found",
            "label": label,
            "path": text_path,
        }
    return {
        "ok": True,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": label_path,
        "window": window_info,
    }


def locate_receipt_header_scope(jab):
    fast = infer_receipt_header_scope_fast(jab)
    if fast.get("ok"):
        return fast

    foreground_root = foreground_root_hwnd()
    if foreground_root:
        fast_matches = collect_complete_header_scope_matches(
            jab, foreground_root=foreground_root
        )
        if len(fast_matches) == 1:
            match = fast_matches[0]
            return {
                "ok": True,
                "scope_hwnd": match["hwnd"],
                "matches": fast_matches,
                "mode": "foreground-fast",
                "fast_path_attempt": fast,
                "dynamic_index": match.get("dynamic_index"),
                "dynamic_prefix": match.get("dynamic_prefix"),
            }

    matches = collect_complete_header_scope_matches(jab)
    if len(matches) == 1:
        match = matches[0]
        return {
            "ok": True,
            "scope_hwnd": match["hwnd"],
            "matches": matches,
            "mode": "full-scan",
            "fast_path_attempt": fast,
            "dynamic_index": match.get("dynamic_index"),
            "dynamic_prefix": match.get("dynamic_prefix"),
        }
    return {
        "ok": False,
        "matches": matches,
        "fast_path_attempt": fast,
        "reason": f"完整表头 scope 数量不是 1：{len(matches)}",
    }


def collect_complete_header_scope_matches(jab, foreground_root=None):
    matches = []
    for window in jab.get_scoped_windows(None, include_children=True):
        hwnd, _title, class_name, _pid, visible = window
        if foreground_root and window_root_hwnd(hwnd) != foreground_root:
            continue
        if not visible or class_name != "SunAwtCanvas" or not jab.dll.isJavaWindow(hwnd):
            continue
        ok_labels = []
        dynamic_index = None
        dynamic_prefix = None
        for label in HEADER_LABEL_ALIASES:
            found = find_receipt_header_form_field(jab, label, scope_hwnd=hwnd)
            if found.get("ok"):
                ok_labels.append(label)
                if dynamic_index is None and found.get("dynamic_index") is not None:
                    dynamic_index = found.get("dynamic_index")
                    dynamic_prefix = found.get("dynamic_prefix")
                jab.release_contexts(found["vm_id"], found["owned_contexts"])
        if len(ok_labels) == len(HEADER_LABEL_ALIASES):
            matches.append(
                {
                    "hwnd": int(hwnd),
                    "labels": ok_labels,
                    "dynamic_index": dynamic_index,
                    "dynamic_prefix": dynamic_prefix,
                }
            )
    return matches


def foreground_root_hwnd():
    if os.name != "nt" or not hasattr(ctypes, "windll"):
        return 0
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    return window_root_hwnd(hwnd)


def window_root_hwnd(hwnd):
    if os.name != "nt" or not hasattr(ctypes, "windll") or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(wintypes.HWND(int(hwnd)), 2) or 0)


def find_label_following_text(jab, vm_id, context, label, path, depth, owned_contexts):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if depth >= jab.max_depth or role == "table":
        return None

    child_infos = []
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_info = jab.get_context_info(vm_id, child)
        child_path = f"{path}.{index}"
        if child_info:
            child_infos.append((index, child, child_info, child_path))
        else:
            jab.release_contexts(vm_id, [child])

    try:
        for position, (_index, child, child_info, child_path) in enumerate(child_infos):
            child_role = (
                child_info.role_en_US.strip() or child_info.role.strip()
            ).lower()
            child_states = (
                child_info.states_en_US.strip() or child_info.states.strip()
            ).lower()
            if (
                child_role == "label"
                and child_info.name.strip() == label
                and "visible" in child_states
            ):
                for _next_index, next_child, _next_info, next_path in child_infos[
                    position + 1 :
                ]:
                    text_context, text_owned, text_path = first_text_descendant(
                        jab,
                        vm_id,
                        next_child,
                        next_path,
                        depth + 1,
                    )
                    if text_context:
                        keep = [item[1] for item in child_infos] + text_owned
                        return (
                            text_context,
                            owned_contexts + keep,
                            text_path,
                            child_path,
                        )
                    continue

        for _index, child, _child_info, child_path in child_infos:
            result = find_label_following_text(
                jab,
                vm_id,
                child,
                label,
                child_path,
                depth + 1,
                owned_contexts + [item[1] for item in child_infos],
            )
            if result:
                return result
    finally:
        pass
    for _index, child, _child_info, _child_path in child_infos:
        jab.release_contexts(vm_id, [child])
    return None


def first_text_descendant(jab, vm_id, context, path, depth):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None, [], None
    role = (info.role_en_US.strip() or info.role.strip()).lower()
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    if (
        role == "text"
        and "visible" in states
        and info.x >= 0
        and info.y >= 0
        and info.width > 0
        and info.height > 0
    ):
        return context, [], path
    if depth >= jab.max_depth:
        return None, [], None
    owned = []
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        child_path = f"{path}.{index}"
        found, found_owned, found_path = first_text_descendant(
            jab,
            vm_id,
            child,
            child_path,
            depth + 1,
        )
        if found:
            return found, owned + [child] + found_owned, found_path
        jab.release_contexts(vm_id, [child])
    return None, owned, None


def set_text_by_control_name(
    jab,
    control_name,
    value,
    commit_key=None,
    accepted_text=None,
    unlock_query=None,
):
    context, vm_id, owned_contexts, owned_indexes, window_info = (
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
    if not context:
        return {"ok": False, "reason": "control not found"}
    path = "0" + "".join(f".{index}" for index in owned_indexes)
    try:
        before = jab.get_text_context_value(vm_id, context)
        info_before = jab.get_context_info(vm_id, context)
        ok = jab.set_text_context(vm_id, context, value)
        time.sleep(0.3)
        commit = None
        if ok and commit_key:
            request_focus_ok = True
            if hasattr(jab.dll, "requestFocus"):
                request_focus_ok = bool(jab.dll.requestFocus(vm_id, context))
            action_result = do_context_commit_action(jab, vm_id, context)
            post_key_ok = False
            if not action_result.get("ok"):
                post_key_ok = post_key_to_hwnd(window_info.get("hwnd"), commit_key)
            backend_state = wait_backend_field_state(
                jab,
                vm_id,
                context,
                value=value,
                accepted_text=accepted_text,
                unlock_query=unlock_query,
                timeout=3.5,
            )
            commit = {
                "key": commit_key,
                "request_focus_ok": bool(request_focus_ok),
                "target_window": window_info,
                "action": action_result,
                "post_key_ok": bool(post_key_ok),
                "accepted_text": accepted_text,
                "unlock_query": unlock_query,
                "backend_state": backend_state,
            }
            ok = bool(
                ok
                and (
                    backend_state.get("written")
                    or backend_state.get("accepted")
                    or backend_state.get("unlocked")
                )
            )
        else:
            backend_state = wait_backend_field_state(
                jab,
                vm_id,
                context,
                value=value,
                accepted_text=accepted_text,
                unlock_query=unlock_query,
                timeout=1.0,
            )
            ok = bool(
                ok
                and (
                    backend_state.get("written")
                    or backend_state.get("accepted")
                    or backend_state.get("unlocked")
                )
            )
        after = jab.get_text_context_value(vm_id, context)
        info_after = jab.get_context_info(vm_id, context)
        return {
            "ok": bool(ok),
            "path": path,
            "name_before": (info_before.name.strip() if info_before else None),
            "description_before": (
                info_before.description.strip() if info_before else None
            ),
            "text_before": before,
            "name_after": (info_after.name.strip() if info_after else None),
            "description_after": (
                info_after.description.strip() if info_after else None
            ),
            "text_after": after,
            "backend_state": backend_state,
            "commit": commit,
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def find_context_with_window(
    jab,
    name,
    roles=(),
    timeout=None,
    require_showing=False,
    window_title=None,
    window_class=None,
    visible_only=True,
    scope_hwnd=None,
):
    deadline = time.time() + (timeout or jab.search_timeout)
    normalized_roles = {role.lower() for role in roles}
    while time.time() < deadline:
        windows = jab.get_scoped_windows(scope_hwnd, include_children=True)
        for hwnd, title, class_name, pid, visible in windows:
            if visible_only and not visible:
                continue
            if (
                scope_hwnd is None
                and window_title is not None
                and title != window_title
            ):
                continue
            if (
                scope_hwnd is None
                and window_class is not None
                and class_name != window_class
            ):
                continue
            if not jab.dll.isJavaWindow(hwnd):
                continue
            from tools.jab_probe import JOBJECT

            vm_id_ref = ctypes.c_long()
            root_context = JOBJECT()
            if not jab.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id_ref),
                ctypes.byref(root_context),
            ):
                continue
            context, owned_contexts, owned_indexes = jab.find_in_tree_with_path(
                vm_id_ref.value,
                root_context.value,
                name,
                normalized_roles,
                require_showing,
                depth=0,
                owned_contexts=[],
                owned_indexes=[],
            )
            if context:
                return (
                    context,
                    vm_id_ref.value,
                    owned_contexts,
                    owned_indexes,
                    {
                        "hwnd": int(hwnd),
                        "title": title,
                        "class_name": class_name,
                        "pid": pid,
                        "visible": visible,
                    },
                )
            jab.release_contexts(vm_id_ref.value, [root_context.value])
        time.sleep(0.2)
    return None, None, [], [], {}


def post_key_to_hwnd(hwnd, key):
    if os.name != "nt" or not hwnd:
        return False
    key_map = {
        "enter": 0x0D,
        "tab": 0x09,
    }
    vk = key_map.get(str(key).lower())
    if not vk:
        return False
    user32 = ctypes.windll.user32
    hwnd = wintypes.HWND(int(hwnd))
    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    down_ok = bool(user32.PostMessageW(hwnd, WM_KEYDOWN, vk, 0))
    up_ok = bool(user32.PostMessageW(hwnd, WM_KEYUP, vk, 0))
    return down_ok and up_ok


def do_context_commit_action(jab, vm_id, context):
    actions = jab.get_action_names(vm_id, context)
    preferred = ("确认", "确定", "提交", "单击", "click", "press")
    for action_name in preferred:
        if action_name not in actions:
            continue
        try:
            ok = bool(jab.do_action(vm_id, context, action_name=action_name))
        except Exception as exc:
            return {
                "ok": False,
                "action": action_name,
                "exception": repr(exc),
                "actions": actions,
            }
        if ok:
            time.sleep(0.8)
            return {"ok": True, "action": action_name, "actions": actions}
    return {"ok": False, "reason": "no commit action", "actions": actions}


def wait_backend_field_state(
    jab,
    vm_id,
    context,
    value=None,
    accepted_text=None,
    unlock_query=None,
    timeout=3.0,
):
    deadline = time.time() + timeout
    last = {
        "accepted": False,
        "written": False,
        "unlocked": False,
        "text": None,
        "name": None,
        "description": None,
    }
    while time.time() < deadline:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        last = describe_backend_field_state(info, text, value, accepted_text)
        if unlock_query:
            last["unlocked"] = contains_visible_control(jab, unlock_query, timeout=0.2)
        if last.get("accepted") or last.get("written") or last.get("unlocked"):
            return last
        time.sleep(0.2)
    return last


def describe_backend_field_state(info, text, value=None, accepted_text=None):
    name = info.name.strip() if info else ""
    description = info.description.strip() if info else ""
    accepted = backend_field_accepts(info, text, value, accepted_text)
    written = backend_field_has_written_value(info, text, value)
    return {
        "accepted": bool(accepted),
        "written": bool(written),
        "unlocked": False,
        "text": text,
        "name": name,
        "description": description,
    }


def backend_field_has_written_value(info, text, value=None):
    expected = str(value).strip() if value is not None else ""
    if not info or not expected:
        return False
    actual_text = str(text or "").strip()
    description = info.description.strip()
    return actual_text == expected or description == expected


def backend_field_accepts(info, text, value=None, accepted_text=None):
    if not info:
        return False
    if accepted_text:
        return context_contains(info, accepted_text)
    expected = str(value).strip() if value is not None else ""
    actual_text = str(text or "").strip()
    description = info.description.strip()
    if expected and (actual_text == expected or description == expected):
        return True
    return bool(description)


def context_contains(info, expected_text):
    if not info or not expected_text:
        return False
    expected = str(expected_text).strip()
    haystack = " ".join(
        part
        for part in (
            info.name.strip(),
            info.description.strip(),
            info.role.strip(),
            info.role_en_US.strip(),
            info.states.strip(),
            info.states_en_US.strip(),
        )
        if part
    )
    return expected in haystack


def contains_visible_control(jab, query, timeout=2.0):
    if not query:
        return False
    deadline = time.time() + timeout
    needle = str(query).lower()
    while time.time() < deadline:
        for window in jab.get_scoped_windows(None, include_children=True):
            hwnd, title, class_name, pid, visible = window
            if (
                not visible
                or class_name != "SunAwtCanvas"
                or not jab.dll.isJavaWindow(hwnd)
            ):
                continue
            import ctypes
            from tools.jab_probe import JOBJECT

            vm_id = ctypes.c_long()
            root_context = JOBJECT()
            if not jab.dll.getAccessibleContextFromHWND(
                hwnd,
                ctypes.byref(vm_id),
                ctypes.byref(root_context),
            ):
                continue
            try:
                if tree_contains_control(
                    jab, vm_id.value, root_context.value, needle, depth=0
                ):
                    return True
            finally:
                jab.release_contexts(vm_id.value, [root_context.value])
        time.sleep(0.2)
    return False


def tree_contains_control(jab, vm_id, context, needle, depth):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return False
    role = info.role_en_US.strip() or info.role.strip()
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    text = f"{info.name} {info.description} {role} {states}".lower()
    if "visible" in states and "showing" in states and needle in text:
        return True
    if depth >= jab.max_depth or role.lower() == "table":
        return False
    for index in range(min(info.childrenCount, jab.max_children)):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        try:
            if tree_contains_control(jab, vm_id, child, needle, depth + 1):
                return True
        finally:
            jab.release_contexts(vm_id, [child])
    return False


def fill_detail_line(jab, business):
    steps = []
    for col, name, value in [
        (1, "收款业务类型", business["main_business_type"]),
        (4, "收款银行账户", business["bank_account"]),
        (5, "科目", business["main_subject"]),
        (7, "贷方原币金额", business["amount"]),
        (11, "结算方式", business["settlement"]),
    ]:
        steps.append(fill_cell(jab, 0, col, name, value))
    return steps


def fill_cell(jab, row, col, name, value):
    attempts = []
    for mode, kwargs in [
        (
            "direct",
            {"set_cell_text": value, "commit_key": "none", "focus_target": "cell"},
        ),
    ]:
        try:
            result = select_cell(
                jab,
                table_index=0,
                row=row,
                col=col,
                window_title=None,
                locate_body_table=True,
                wait=0.35,
                **kwargs,
            )
        except Exception as exc:
            result = {"ok": False, "exception": repr(exc)}
        attempts.append({"mode": mode, "result": compact_selection_result(result)})
        if cell_changed(result, value):
            break
        time.sleep(0.2)
    ok = any(cell_changed(attempt["result"], value) for attempt in attempts)
    return {
        "step": "detail_cell",
        "ok": ok,
        "blocked": not ok,
        "reason": (
            None
            if ok
            else "backend cell input failed; global keyboard fallbacks are disabled"
        ),
        "row": row,
        "col": col,
        "name": name,
        "value": value,
        "attempts": attempts,
    }


def compact_selection_result(result):
    if not isinstance(result, dict):
        return result
    keep = {
        key: result.get(key)
        for key in (
            "ok",
            "reason",
            "row_count",
            "col_count",
            "child_index",
            "selected_before",
            "selected_after",
            "cell_text_before",
            "cell_text_after",
            "edit",
            "exception",
        )
        if key in result
    }
    return keep


def cell_changed(result, value):
    after = str((result or {}).get("cell_text_after") or "")
    return (
        after
        and after != str((result or {}).get("cell_text_before") or "")
        and str(value) in after
    )


def read_body_table(jab, step, scope_hwnd=None):
    located = locate_receipt_body_table(jab, max_rows=3, scope_hwnd=scope_hwnd)
    best = located.get("best")
    if not best:
        return {
            "step": step,
            "ok": False,
            "reason": "body table not found",
            "candidates": located.get("candidates", [])[:3],
        }
    return {
        "step": step,
        "ok": True,
        "path": best.get("path"),
        "row_count": best.get("row_count"),
        "col_count": best.get("col_count"),
        "rows": best.get("rows"),
    }


if __name__ == "__main__":
    raise SystemExit(main())
