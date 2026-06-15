import argparse
import ctypes
from ctypes import wintypes
from datetime import datetime
from decimal import Decimal
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_config import ReceiptEntryConfig  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.read_receipt_excel_row import DEFAULT_FIELDS  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_new_probe import (  # noqa: E402
    collect_receipt_new_windows,
    detect_self_made_entry_state,
)
from tools.receipt_table_cell_probe import select_cell  # noqa: E402


CURRENCY_NAMES = {"USD": "美元", "RMB": "人民币", "CNY": "人民币"}
HEADER_DYNAMIC_PREFIX_BASE = "0.0.1.0.0.0.0"
HEADER_DYNAMIC_MAX_INDEX = 8
HEADER_COMMON_SUFFIX_TEMPLATE = "0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.{index}.0"
HEADER_COMMON_LABEL_SUFFIX_TEMPLATE = "0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.{index}"
FINANCE_ORG_SUFFIX = "0.0.0.1.1.0.0.0.0.1.1.1.0"
FINANCE_ORG_LABEL_SUFFIX = "0.0.0.1.1.0.0.0.0.1.1.0"
HEADER_FORM_TEXT_INDEXES = {
    "单据日期": 5,
    "币种": 13,
    "收款银行账户": 15,
    "客户": 17,
    "结算方式": 31,
}
HEADER_REQUIRED_LABELS = ("财务组织", "客户", "单据日期", "币种", "结算方式")
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
        header_steps = fill_header(jab, business)
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


def fill_header(jab, business):
    steps = []
    scope = locate_receipt_header_scope(jab)
    if not scope.get("ok"):
        return [
            {
                "step": "blocked",
                "reason": "header scope not resolved",
                "scope": scope,
            }
        ]
    dynamic_index = scope.get("dynamic_index")
    scope_hwnd = scope.get("scope_hwnd")
    semantic_preload = HeaderSemanticPreload(
        jab.config if hasattr(jab, "config") else None,
        scope_hwnd,
        [label for label in HEADER_REQUIRED_LABELS if label != "财务组织"],
    )
    semantic_preload.start()
    for field in [
        {
            "label": "财务组织",
            "value": business["finance_org_code"],
            "accepted_text": business.get("finance_org_name"),
            "dynamic_path": True,
        },
        {
            "label": "客户",
            "value": business["customer_code"],
            "dynamic_path": True,
        },
        {
            "label": "单据日期",
            "value": business["document_date"],
            "dynamic_path": True,
        },
        {
            "label": "币种",
            "value": business.get("header_currency_code") or business.get("currency"),
            "dynamic_path": True,
        },
        {
            "label": "结算方式",
            "value": business.get("settlement") or "网银",
            "dynamic_path": True,
        },
    ]:
        label = field["label"]
        value = field["value"]
        if field.get("dynamic_path"):
            result = set_receipt_header_dynamic_field(
                jab,
                label,
                value,
                dynamic_index,
                scope_hwnd,
                accepted_text=field.get("accepted_text"),
                semantic_snapshot=semantic_preload.snapshot(timeout=0.0),
                semantic_preload=semantic_preload,
            )
            ok = bool(result.get("ok"))
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "dynamic_path_commit",
                    "scope": {
                        "mode": scope.get("mode"),
                        "scope_hwnd": scope_hwnd,
                        "dynamic_index": dynamic_index,
                        "dynamic_prefix": scope.get("dynamic_prefix"),
                    },
                    **result,
                    "semantic_preload": semantic_preload.snapshot(timeout=0.0),
                }
            )
        else:
            ok = False
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "unsupported",
                    "ok": False,
                    "reason": "header field must use dynamic path",
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


class HeaderSemanticPreload:
    def __init__(self, config, scope_hwnd, labels):
        self.config = config
        self.scope_hwnd = scope_hwnd
        self.labels = list(labels)
        self.started_at = None
        self.finished_at = None
        self.result = {
            "status": "not_started",
            "scope_hwnd": scope_hwnd,
            "fields": {},
        }
        self._thread = None

    def start(self):
        if not self.config:
            self.result["status"] = "skipped"
            self.result["reason"] = "JAB config unavailable"
            return
        self.started_at = time.perf_counter()
        self.result["status"] = "running"
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self):
        jab = JABOperator(self.config)
        try:
            jab.ensure_started()
            fields = {}
            for label in self.labels:
                found = find_receipt_header_field_by_semantic_label(
                    jab,
                    label,
                    scope_hwnd=self.scope_hwnd,
                )
                fields[label] = {
                    "ok": bool(found.get("ok")),
                    "path": found.get("path"),
                    "label_path": found.get("label_path"),
                    "reason": found.get("reason"),
                }
                if found.get("ok"):
                    jab.release_contexts(found["vm_id"], found["owned_contexts"])
            self.result["fields"] = fields
            self.result["ok"] = all(item.get("ok") for item in fields.values())
            self.result["status"] = "ready"
        except Exception as exc:
            self.result["ok"] = False
            self.result["status"] = "error"
            self.result["reason"] = f"{type(exc).__name__}: {exc}"
        finally:
            self.finished_at = time.perf_counter()
            if self.started_at is not None:
                self.result["seconds"] = round(self.finished_at - self.started_at, 3)
            jab.close()

    def snapshot(self, timeout=0.0):
        if self._thread is not None:
            self._thread.join(timeout=max(float(timeout or 0), 0.0))
        data = dict(self.result)
        if self.started_at is not None and self.finished_at is None:
            data["elapsed_seconds"] = round(time.perf_counter() - self.started_at, 3)
        return data


def wait_header_account_description(jab, timeout=5.0, scope=None):
    deadline = time.time() + timeout
    last = None
    if scope is None:
        scope = locate_receipt_header_scope(jab)
    if not scope.get("ok"):
        return {"text": None, "description": "", "accepted": False, "scope": scope}
    dynamic_index = scope.get("dynamic_index")
    scope_hwnd = scope.get("scope_hwnd")
    while time.time() < deadline:
        found = find_receipt_header_field_by_dynamic_path(
            jab,
            "收款银行账户",
            dynamic_index,
            scope_hwnd=scope_hwnd,
            require_showing=False,
            require_valid_bounds=False,
        )
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


def set_receipt_header_dynamic_field(
    jab,
    label,
    value,
    dynamic_index,
    scope_hwnd,
    accepted_text=None,
    semantic_snapshot=None,
    semantic_preload=None,
):
    found = find_receipt_header_field_by_dynamic_path(
        jab,
        label,
        dynamic_index,
        scope_hwnd=scope_hwnd,
        require_showing=label != "财务组织",
        require_valid_bounds=label != "财务组织",
    )
    path_attempt = found
    if not found.get("ok"):
        if (
            semantic_preload is not None
            and semantic_snapshot
            and semantic_snapshot.get("status") == "running"
        ):
            semantic_snapshot = semantic_preload.snapshot(timeout=1.5)
        takeover = find_receipt_header_field_by_semantic_cache(
            jab,
            label,
            semantic_snapshot,
            scope_hwnd,
        )
        if takeover.get("ok"):
            found = takeover
        else:
            return {
                "ok": False,
                "stage": "resolve",
                "path_attempt": path_attempt,
                "semantic_takeover": takeover,
            }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    window_info = found["window"]
    try:
        info_before = jab.get_context_info(vm_id, context)
        before = jab.get_text_context_value(vm_id, context)
        set_ok = bool(jab.set_text_context(vm_id, context, value))
        if set_ok and hasattr(jab.dll, "requestFocus"):
            jab.dll.requestFocus(vm_id, context)
        commit_action = (
            do_context_commit_action(jab, vm_id, context) if set_ok else None
        )
        enter_ok = (
            post_key_to_hwnd(window_info.get("hwnd"), "enter") if set_ok else False
        )
        backend_state = wait_backend_field_state(
            jab,
            vm_id,
            context,
            value=value,
            accepted_text=accepted_text,
            timeout=3.5,
        )
        info_after = jab.get_context_info(vm_id, context)
        after = jab.get_text_context_value(vm_id, context)
        return {
            "ok": bool(
                set_ok
                and (
                    backend_state.get("written")
                    or backend_state.get("accepted")
                    or backend_state.get("unlocked")
                )
            ),
            "path": found.get("path"),
            "label_path": found.get("label_path"),
            "dynamic_index": found.get("dynamic_index"),
            "dynamic_prefix": found.get("dynamic_prefix"),
            "source": found.get("source") or "path",
            "path_attempt": path_attempt,
            "semantic_takeover": None
            if found is path_attempt
            else {
                "ok": True,
                "path": found.get("path"),
                "label_path": found.get("label_path"),
                "source": found.get("source"),
            },
            "text_before": before,
            "description_before": (
                info_before.description.strip() if info_before else None
            ),
            "set_ok": bool(set_ok),
            "commit_action": commit_action,
            "enter_ok": bool(enter_ok),
            "backend_state": backend_state,
            "accepted_text": accepted_text_from_backend(
                backend_state,
                value,
                accepted_text,
            ),
            "text_after": after,
            "description_after": (
                info_after.description.strip() if info_after else None
            ),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def find_receipt_header_field_by_semantic_cache(
    jab,
    label,
    semantic_snapshot,
    scope_hwnd,
):
    field = ((semantic_snapshot or {}).get("fields") or {}).get(label) or {}
    path = field.get("path")
    if not path:
        return {"ok": False, "reason": "semantic preload path missing", "label": label}
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "reason": "semantic preload path not found",
            "label": label,
            "path": path,
        }
    dynamic_index = extract_receipt_header_dynamic_index(path)
    return {
        "ok": True,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": path,
        "label_path": field.get("label_path"),
        "window": window_info,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index)
        if dynamic_index is not None
        else None,
        "source": "semantic_preload",
    }


def find_receipt_header_field_by_semantic_label(jab, label, scope_hwnd=None):
    label_found = find_context_with_window(
        jab,
        label,
        roles=("label",),
        timeout=1.5,
        require_showing=True,
        window_class="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
    )
    label_context, vm_id, owned_contexts, owned_indexes, window = label_found
    if not label_context:
        return {"ok": False, "label": label, "reason": "semantic label not found"}
    label_path = None
    if owned_indexes:
        label_path = "0" + "".join(f".{index}" for index in owned_indexes)
    jab.release_contexts(vm_id, owned_contexts)
    if not label_path:
        return {"ok": False, "label": label, "reason": "semantic label path missing"}
    text_path = infer_header_text_path_from_label_path(label, label_path)
    if not text_path:
        return {
            "ok": False,
            "label": label,
            "label_path": label_path,
            "reason": "semantic label path cannot infer text path",
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        text_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "label": label,
            "label_path": label_path,
            "path": text_path,
            "reason": "semantic inferred text path not found",
        }
    return {
        "ok": True,
        "label": label,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": label_path,
        "window": window_info,
    }


def infer_header_text_path_from_label_path(label, label_path):
    parts = split_header_path(label_path)
    if not parts:
        return None
    if label == "财务组织":
        if parts[-1] == 0:
            return ".".join(str(part) for part in [*parts[:-1], 1, 0])
        return None
    if parts[-1] % 2 != 0:
        return None
    return ".".join(str(part) for part in [*parts[:-1], parts[-1] + 1, 0])


def split_header_path(path):
    try:
        return [int(part) for part in str(path).split(".") if part != ""]
    except ValueError:
        return []


def accepted_text_from_backend(backend_state, raw_value=None, preferred=None):
    if preferred:
        return str(preferred).strip()
    for key in ("description", "text", "name"):
        text = str((backend_state or {}).get(key) or "").strip()
        if text and text != str(raw_value or "").strip():
            return text
    return ""


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
        return (
            f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{FINANCE_ORG_LABEL_SUFFIX}"
        )
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
        for label in HEADER_REQUIRED_LABELS:
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
        if len(ok_labels) == len(HEADER_REQUIRED_LABELS):
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
        if (
            not visible
            or class_name != "SunAwtCanvas"
            or not jab.dll.isJavaWindow(hwnd)
        ):
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


def locate_receipt_header_scope(jab):
    fast = infer_receipt_header_scope_fast(jab)
    if fast.get("ok"):
        return fast
    semantic = infer_receipt_header_scope_by_semantic(jab)
    if semantic.get("ok"):
        semantic["fast_path_attempt"] = fast
        return semantic
    return {
        "ok": False,
        "fast_path_attempt": fast,
        "semantic_attempt": semantic,
        "reason": "未能用固定表头后缀 path 或语义路径推断唯一定位当前收款单表头 scope",
    }


def infer_receipt_header_scope_by_semantic(jab):
    found = find_receipt_header_field_by_semantic_label(
        jab,
        FAST_HEADER_SCOPE_LABEL,
        scope_hwnd=None,
    )
    if not found.get("ok"):
        return {
            "ok": False,
            "mode": "semantic-path-inference",
            "reason": found.get("reason") or "财务组织语义定位失败",
            "attempt": found,
        }
    dynamic_index = extract_receipt_header_dynamic_index(found.get("path"))
    scope_hwnd = ((found.get("window") or {}).get("hwnd")) or None
    jab.release_contexts(found["vm_id"], found["owned_contexts"])
    if dynamic_index is None or not scope_hwnd:
        return {
            "ok": False,
            "mode": "semantic-path-inference",
            "reason": "财务组织语义路径无法推出动态前缀或窗口",
            "attempt": {
                "path": found.get("path"),
                "label_path": found.get("label_path"),
                "window": found.get("window"),
            },
        }
    prefix = infer_receipt_header_dynamic_prefix(
        jab,
        scope_hwnd=scope_hwnd,
        dynamic_max=dynamic_index,
        require_showing=False,
        require_valid_bounds=False,
    )
    if not prefix.get("ok"):
        return {
            "ok": False,
            "mode": "semantic-path-inference",
            "reason": "语义推出前缀后未能通过必填表头字段校验",
            "attempt": found,
            "prefix_check": prefix,
        }
    return {
        "ok": True,
        "scope_hwnd": scope_hwnd,
        "mode": "semantic-path-inference",
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        "matched_labels": prefix.get("matched_labels"),
        "semantic_label_path": found.get("label_path"),
        "semantic_text_path": found.get("path"),
    }


def foreground_root_hwnd():
    if os.name != "nt" or not hasattr(ctypes, "windll"):
        return 0
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    return window_root_hwnd(hwnd)


def window_root_hwnd(hwnd):
    if os.name != "nt" or not hasattr(ctypes, "windll") or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(wintypes.HWND(int(hwnd)), 2) or 0)


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
            else "backend cell input failed; global keyboard input is disabled"
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
