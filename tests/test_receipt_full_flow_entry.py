# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程入口的计划行选择、业务值映射和保存安全确认
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_entry.py

from datetime import date
from decimal import Decimal

import pytest

from core.receipt_models import ReceiptPlanIssue, ReceiptPlanRow
from tools import receipt_full_flow_entry as full_flow
from core import receipt_counterparty as cp
from core import receipt_locator_cache as _locator
from core import receipt_save_cancel as _save_cancel
from core import receipt_report as _report
from core import receipt_row_stages as _row_stages
from core.receipt_locator_cache import (
    extract_entry_anchor_path,
    extract_entry_dynamic_index,
    extract_entry_scope_hwnd,
)

from core.receipt_report import build_console_report_lines
from tools.receipt_full_flow_entry import (
    business_from_plan_row,
    build_header_verify_expectations,
    cache_receipt_header_scope,
    customer_name_similarity,
    confirm_save,
    ensure_header_counterparty_customer,
    extract_header_accepted_text,
    header_currency_matches,
    open_self_made_entry,
    parse_args,
    post_query_failure_reasons,
    read_customer_name_after_header,
    run_one_row,
    save_receipt_by_ctrl_s,
    select_plan_rows,
    write_extra_text_field_by_dynamic_path,
    verify_and_repair_header_targets,
)
from core.receipt_post_save_query import target_to_match_row
from core.receipt_post_save_query import format_query_exception
from core.receipt_save_cancel import should_retry_row_by_cancel_reopen

_PATCH_TARGET_MODULES = [full_flow, cp, _locator, _save_cancel, _report, _row_stages]


def patch_all(monkeypatch, name, value):
    """把 helper 同时打到所有导入它的拆分模块,使调用方无论在哪都命中补丁。"""
    patched = False
    for _m in _PATCH_TARGET_MODULES:
        if hasattr(_m, name):
            monkeypatch.setattr(_m, name, value)
            patched = True
    if not patched:
        raise AttributeError(name)

REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER = ensure_header_counterparty_customer


def plan_row(row, fee=Decimal("0.00"), extra_text_fields=None):
    return ReceiptPlanRow(
        row=row,
        receipt_date=date(2026, 5, 22),
        payer_name="ACME LTD",
        raw_amount=Decimal("1090.00"),
        bank="招行",
        currency="人民币",
        customer_code="YW03574",
        fee=fee,
        organization_code="A001",
        organization_name="上海移为通信技术股份有限公司",
        organization_short_name="移为",
        account_id="cmb_a001",
        account_label="大陆招行",
        account_no="FTE1219165931831",
        header_currency_code="CNY",
        duplicate_key=("A001", "2026-05-22", "招行"),
        extra_text_fields=extra_text_fields or {},
    )


def open_report_with_header_anchor(hwnd=2002, dynamic_index=5):
    return {
        "ok": True,
        "entry_state": {
            "hits": [
                {
                    "window": {
                        "hwnd": hwnd,
                        "class_name": "SunAwtCanvas",
                        "visible": True,
                    },
                    "control": {
                        "path": f"0.0.1.0.0.0.0.{dynamic_index}.0.0",
                        "dynamic_index": dynamic_index,
                    },
                }
            ]
        },
    }


def pipeline_wait_ok_with_cny_snapshot(task_ids=()):
    return {
        "ok": True,
        "submitted": list(task_ids),
        "done": len(list(task_ids)),
        "snapshots": [
            {
                "id": "snapshot-0",
                "ok": True,
                "snapshot": {
                    "ok": True,
                    "rows": [
                        {
                            "row_index": 0,
                            "cells": {"6": "1.0000", "7": "1,090.00"},
                        }
                    ],
                },
            }
        ],
    }


class FakeInfo:
    role = "text"
    role_en_US = "text"
    states = "enabled,visible,showing,editable"
    states_en_US = "enabled,visible,showing,editable"
    childrenCount = 0

    def __init__(self, name="", description=""):
        self.name = name
        self.description = description


class FakeComboInfo(FakeInfo):
    role = "combo box"
    role_en_US = "combo box"
    states = "enabled,focusable,visible,showing,opaque,collapsed"
    states_en_US = "enabled,focusable,visible,showing,opaque,collapsed"


class Args:
    start_row: int | None = None
    limit: int | None = None


@pytest.fixture(autouse=True)
def default_counterparty_header_ok(monkeypatch):
    patch_all(monkeypatch, "ensure_header_counterparty_customer",
        lambda *_args, **_kwargs: {
            "ok": True,
            "skipped": True,
            "actual": "客户",
            "path": "counterparty.path",
        },
    )


def test_extract_header_accepted_text_rejects_java_object_string():
    assert (
        extract_header_accepted_text(
            [
                {
                    "label": "客户",
                    "value": "YW00178",
                    "post_write_snapshot": {
                        "description": "[Ljava.lang.String;@75acf5a0",
                    },
                }
            ],
            "客户",
        )
        == ""
    )


def test_header_unified_check_failure_triggers_cancel_reopen_retry():
    assert should_retry_row_by_cancel_reopen(
        {
            "ok": False,
            "failed_step": "header-unified-check",
            "reason": "表头统一校验补写后仍有缺失",
            "save": {"skipped": True},
        }
    ) is True


def test_detail_main_line_field_failure_triggers_cancel_reopen_retry():
    assert should_retry_row_by_cancel_reopen(
        {
            "ok": False,
            "failed_step": "detail-main-line",
            "reason": "明细主行写入失败",
            "save": {"skipped": True},
            "detail_steps": [
                {
                    "ok": False,
                    "name": "收款银行账户",
                    "reason": "即时校验未匹配：字段=收款银行账户",
                }
            ],
        }
    ) is True


def test_detail_main_line_stop_hotkey_does_not_retry():
    assert should_retry_row_by_cancel_reopen(
        {
            "ok": False,
            "failed_step": "detail-main-line",
            "reason": "明细主行写入失败",
            "save": {"skipped": True},
            "detail_steps": [
                {
                    "ok": False,
                    "name": "科目",
                    "reason": "检测到紧急停止键 ctrl+shift+q",
                }
            ],
        }
    ) is False


def test_detail_main_line_table_shape_failure_does_not_retry():
    assert should_retry_row_by_cancel_reopen(
        {
            "ok": False,
            "failed_step": "detail-main-line",
            "reason": "明细主行写入失败",
            "save": {"skipped": True},
            "detail_steps": [
                {
                    "ok": False,
                    "name": "明细表",
                    "reason": "明细表尺寸异常：0 行 x 0 列，目标第 1 行",
                }
            ],
        }
    ) is False


def test_header_counterparty_api_repair_failure_triggers_cancel_reopen_retry():
    assert should_retry_row_by_cancel_reopen(
        {
            "ok": False,
            "failed_step": "header-counterparty-type",
            "reason": "往来对象未确认客户",
            "save": {"skipped": True},
            "header_counterparty": {
                "ok": False,
                "actual": "供应商",
                "state": {"state": "repairable-conflict"},
                "repair": {"method": "embedded-selection-api", "ok": True},
                "after_detail": {"value": "供应商"},
            },
        }
    ) is True


def test_header_counterparty_locator_failure_does_not_retry():
    assert should_retry_row_by_cancel_reopen(
        {
            "ok": False,
            "failed_step": "header-counterparty-type",
            "reason": "nearby scope 不是 Java 窗口",
            "save": {"skipped": True},
            "header_counterparty": {
                "ok": False,
                "actual": "供应商",
                "detail": {"value": "供应商"},
            },
        }
    ) is False


def test_read_customer_name_after_header_uses_customer_description(monkeypatch):
    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description="INDUSTRIAS METALURGICAS PESCARMONA")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "path": "customer.path",
            "label_path": "customer.label",
        },
    )

    result = read_customer_name_after_header(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "客户",
                "value": "YW00178",
                "dynamic_index": 6,
                "path": "customer.path",
            }
        ],
        6,
        197550,
    )

    assert result["ok"] is True
    assert result["value"] == "INDUSTRIAS METALURGICAS PESCARMONA"
    assert result["source"] == "path-readback"


def test_read_customer_name_after_header_polls_until_customer_description(monkeypatch):
    class FakeJAB:
        def __init__(self):
            self.descriptions = ["", "ACME NC"]

        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description=self.descriptions.pop(0))

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "path": "customer.path",
            "label_path": "customer.label",
        },
    )

    result = read_customer_name_after_header(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "客户",
                "value": "YW00178",
                "dynamic_index": 6,
                "path": "customer.path",
            }
        ],
        6,
        197550,
        timeout=0.2,
        poll_interval=0.01,
    )

    assert result["ok"] is True
    assert result["value"] == "ACME NC"
    assert len(result["attempts"]) == 2


def test_read_customer_name_after_header_failure_reports_readback(monkeypatch):
    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return FakeInfo(name="客户", description="YW00178")

        def get_text_context_value(self, _vm_id, _context):
            return "YW00178"

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "path": "customer.path",
            "label_path": "customer.label",
        },
    )

    result = read_customer_name_after_header(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "客户",
                "value": "YW00178",
                "dynamic_index": 6,
                "path": "customer.path",
            }
        ],
        6,
        197550,
        timeout=0,
    )

    assert result["ok"] is False
    assert "客户名称未确认" in result["reason"]
    assert "YW00178" in result["reason"]


def test_select_plan_rows_skips_issue_rows_and_defaults_to_all_runnable_rows():
    rows = [plan_row(2), plan_row(3), plan_row(4)]
    issues = [
        ReceiptPlanIssue(
            excel_row=2,
            stage="本地数据校验",
            issue_type="CUSTOMER_CODE_EMPTY",
            field="客户编码",
            raw_value="",
            config_node="receipt_entry.excel.customer_code_column",
            message="客户编码为空",
            action="跳过",
        )
    ]

    args = Args()
    args.start_row = 3
    args.limit = 10
    selected = select_plan_rows(rows, issues, args)

    assert [row.row for row in selected] == [3, 4]


def test_select_plan_rows_requires_start_row():
    args = Args()
    args.limit = 10

    with pytest.raises(SystemExit, match="start-row"):
        select_plan_rows([plan_row(3), plan_row(4)], [], args)


def test_select_plan_rows_limit_zero_means_until_end():
    args = Args()
    args.start_row = 4
    args.limit = 0

    selected = select_plan_rows([plan_row(3), plan_row(4), plan_row(5)], [], args)

    assert [row.row for row in selected] == [4, 5]


def test_select_plan_rows_missing_limit_means_until_end():
    args = Args()
    args.start_row = 4

    selected = select_plan_rows([plan_row(3), plan_row(4), plan_row(5)], [], args)

    assert [row.row for row in selected] == [4, 5]


def test_select_plan_rows_uses_start_row_to_run_until_end():
    args = Args()
    args.start_row = 4
    args.limit = 10

    selected = select_plan_rows([plan_row(2), plan_row(4), plan_row(5)], [], args)

    assert [row.row for row in selected] == [4, 5]


def test_select_plan_rows_can_limit_count_after_start_row():
    args = Args()
    args.start_row = 4
    args.limit = 2

    selected = select_plan_rows([plan_row(2), plan_row(4), plan_row(5), plan_row(6)], [], args)

    assert [row.row for row in selected] == [4, 5]


def test_business_from_plan_row_maps_receipt_plan_to_entry_values():
    business = business_from_plan_row(plan_row(8, fee=Decimal("20.00")))

    assert business["finance_org_code"] == "A001"
    assert business["document_date"] == "2026-05-22"
    assert business["customer_code"] == "YW03574"
    assert business["header_currency_code"] == "CNY"
    assert business["bank_account"] == "FTE1219165931831"
    assert business["amount"] == "1090.00"
    assert business["fee"] == "20.00"
    assert business["has_fee"] is True
    assert business["settlement"] == "网银"


def test_post_save_match_uses_nc_gross_amount_for_fee_rows():
    row = plan_row(8, fee=Decimal("13.00"))

    target = target_to_match_row(
        type(
            "Target",
            (),
            {
                "row": row,
                "row_report": {"nc_customer_name": "ACME NC"},
            },
        )()
    )

    assert target.raw_amount == Decimal("1103.00")


def test_post_query_failure_reasons_collects_group_issues():
    assert post_query_failure_reasons(
        {
            "ok": True,
            "groups": [
                {
                    "ok": True,
                    "match": {
                        "matched": {"839": "D1"},
                        "issues": {"811": "后验未匹配-金额不一致"},
                    },
                }
            ],
        }
    ) == {"811": "后验未匹配-金额不一致"}


def test_post_query_failure_reasons_extracts_group_reason_when_top_level_failed():
    assert post_query_failure_reasons(
        {
            "ok": False,
            "groups": [
                {
                    "ok": False,
                    "target_rows": [851, 852],
                    "reason": "查询失败-鼠标位于屏幕角落",
                    "match": {"matched": {}, "issues": {"851": "查询失败-鼠标位于屏幕角落"}},
                }
            ],
        }
    ) == {
        "851": "查询失败-鼠标位于屏幕角落",
        "852": "查询失败-鼠标位于屏幕角落",
    }


def test_format_query_exception_translates_pyautogui_failsafe():
    class FailSafeException(Exception):
        pass

    exc = FailSafeException("PyAutoGUI fail-safe triggered from mouse moving to a corner of the screen")

    assert format_query_exception(exc) == (
        "鼠标位于屏幕角落，PyAutoGUI 安全保护中断了查询快捷键，请把鼠标移出屏幕角落后重试"
    )


def test_console_summary_reports_post_query_failure():
    lines = build_console_report_lines(
        {
            "ok": False,
            "total_seconds": 12.3,
            "rows": [
                {"excel_row": 811, "ok": True},
                {"excel_row": 839, "ok": True},
            ],
            "post_query_failed_rows": {"811": "后验未匹配-金额不一致"},
        }
    )

    assert "结果：失败" in lines
    assert "录入保存通过行：[811, 839]" in lines
    assert "失败阶段：post-query" in lines
    assert "后验未匹配行 811：后验未匹配-金额不一致" in lines


def test_confirm_save_requires_uppercase_save_without_bypass(monkeypatch):
    class SaveArgs:
        yes_i_understand = False

    monkeypatch.setattr("builtins.input", lambda _prompt: "no")

    with pytest.raises(SystemExit, match="用户取消保存"):
        confirm_save(SaveArgs())


def test_confirm_save_bypass_is_explicit():
    class SaveArgs:
        yes_i_understand = True

    assert confirm_save(SaveArgs()) is None


def test_parse_args_defaults_to_no_start_delay():
    assert parse_args([]).start_delay == 0.0


def test_parse_args_supports_detail_repair_drill():
    assert parse_args(["--diagnose-detail-repair"]).diagnose_detail_repair is True


def test_extract_entry_anchor_path_uses_exact_finance_org_anchor():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.0",
                        "name": "财务组织(O)",
                        "description": "财务组织(O)",
                    }
                }
            ]
        }
    }

    assert (
        extract_entry_anchor_path(report) == "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.1.1.1.0"
    )


def test_open_self_made_entry_always_runs_new_probe(monkeypatch):
    calls = []

    patch_all(monkeypatch, "run_receipt_new_probe",
        lambda: calls.append("new-probe") or {"ok": True, "mode": "new-self-made"},
    )

    assert open_self_made_entry({"receipt_entry": {}}) == {
        "ok": True,
        "mode": "new-self-made",
    }
    assert calls == ["new-probe"]


def test_open_self_made_entry_reuses_existing_jab(monkeypatch):
    calls = []
    jab = object()

    patch_all(monkeypatch, "run_receipt_new_probe",
        lambda: (_ for _ in ()).throw(AssertionError("不应起子进程开单")),
    )
    patch_all(monkeypatch, "run_receipt_new_probe_with_jab",
        lambda actual_jab: (
            calls.append(actual_jab) or {"ok": True, "mode": "in-process"}
        ),
    )

    assert open_self_made_entry({"receipt_entry": {}}, jab) == {
        "ok": True,
        "mode": "in-process",
    }
    assert calls == [jab]


def test_save_receipt_uses_sendinput_ctrl_s_not_jab_button(monkeypatch):
    calls = {"hotkey": 0, "states": 0, "maximize": []}

    class FakeJAB:
        def click_save(self, timeout=None):
            raise AssertionError("收款单保存不能调用凭证/制单保存按钮查找")

        def wait_save_success(self, timeout=None):
            raise AssertionError("收款单保存不等保存成功提示作为触发闭包")

        def press_hotkey(self, *keys, wait=None):
            raise AssertionError("收款单保存应使用 SendInput Ctrl+S")

        def maximize_window_by_handle(self, hwnd):
            calls["maximize"].append(hwnd)
            return True

    patch_all(monkeypatch, "foreground_matches_window",
        lambda window: {"ok": True, "target_window": window, "foreground": {}},
    )
    patch_all(monkeypatch, "collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 12345}],
    )
    monkeypatch.setattr(
        "core.receipt_save_cancel.detect_receipt_parent_new_ready",
        lambda _windows: {
            "ok": True,
            "usable_new_button_count": 1,
            "usable_new_buttons": [{"control": {"description": "新增(Ctrl+N)"}}],
        },
    )

    def fake_detect(_windows):
        calls["states"] += 1
        return {"ok": False, "reason": "已回到新增态"}

    patch_all(monkeypatch, "detect_self_made_entry_state",
        fake_detect,
    )
    patch_all(monkeypatch, "root_hwnd",
        lambda hwnd: hwnd,
    )
    patch_all(monkeypatch, "send_hotkey_ctrl_s",
        lambda: calls.__setitem__("hotkey", calls["hotkey"] + 1),
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), scope_hwnd=12345, timeout=0.5)

    assert result["ok"] is True
    assert result["triggered"] is True
    assert calls["maximize"] == [12345]
    assert calls["hotkey"] == 1
    assert result["hotkey"]["mode"] == "send_input"
    assert result["oracle"]["name"] == "receipt_parent_new_ready_after_save"
    assert result["oracle"]["parent_new_state"]["ok"] is True
    assert calls["states"] == 1


def test_save_receipt_stops_before_oracle_when_foreground_guard_fails(monkeypatch):
    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            raise AssertionError("前台保护失败时不应触发 Ctrl+S")

        def maximize_window_by_handle(self, hwnd):
            return True

    patch_all(monkeypatch, "foreground_matches_window",
        lambda _window: {"ok": False, "reason": "当前前台窗口不是目标 NC 窗口"},
    )
    patch_all(monkeypatch, "root_hwnd",
        lambda hwnd: hwnd,
    )
    patch_all(monkeypatch, "send_hotkey_ctrl_s",
        lambda: (_ for _ in ()).throw(
            AssertionError("前台保护失败时不应触发 Ctrl+S")
        ),
    )
    patch_all(monkeypatch, "collect_receipt_new_windows",
        lambda _jab: (_ for _ in ()).throw(
            AssertionError("Ctrl+S 未发出时不应继续等保存 oracle")
        ),
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), scope_hwnd=12345, timeout=0.5)

    assert result["ok"] is False
    assert result["triggered"] is False
    assert "当前前台窗口不是目标 NC 窗口" in result["reason"]


def test_save_receipt_promotes_scope_hwnd_to_root_before_hotkey(monkeypatch):
    calls = {"guard": [], "maximize": [], "hotkey": 0}

    class FakeJAB:
        def maximize_window_by_handle(self, hwnd):
            calls["maximize"].append(hwnd)
            return True

    patch_all(monkeypatch, "root_hwnd",
        lambda hwnd: 13579 if hwnd == 24680 else hwnd,
    )
    patch_all(monkeypatch, "foreground_matches_window",
        lambda window: calls["guard"].append(window) or {"ok": True},
    )
    patch_all(monkeypatch, "send_hotkey_ctrl_s",
        lambda: calls.__setitem__("hotkey", calls["hotkey"] + 1),
    )
    patch_all(monkeypatch, "collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 13579}],
    )
    patch_all(monkeypatch, "detect_self_made_entry_state",
        lambda _windows: {"ok": False, "reason": "已回到新增态"},
    )
    monkeypatch.setattr(
        "core.receipt_save_cancel.detect_receipt_parent_new_ready",
        lambda _windows: {"ok": True, "usable_new_button_count": 1},
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), scope_hwnd=24680, timeout=0.5)

    assert result["ok"] is True
    assert calls["maximize"] == [13579]
    assert calls["guard"] == [{"hwnd": 13579}]
    assert calls["hotkey"] == 1
    assert result["precondition"]["scope_hwnd"] == 24680
    assert result["precondition"]["target_hwnd"] == 13579


def test_save_receipt_does_not_treat_missing_entry_buttons_as_success_without_new_button(
    monkeypatch,
):
    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            raise AssertionError("收款单保存应使用 SendInput Ctrl+S")

        def maximize_window_by_handle(self, hwnd):
            return True

    patch_all(monkeypatch, "foreground_matches_window",
        lambda _window: {"ok": True},
    )
    patch_all(monkeypatch, "collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 12345}],
    )
    patch_all(monkeypatch, "detect_self_made_entry_state",
        lambda _windows: {"ok": False, "reason": "三按钮读不到"},
    )
    monkeypatch.setattr(
        "core.receipt_save_cancel.detect_receipt_parent_new_ready",
        lambda _windows: {"ok": False, "usable_new_button_count": 0},
    )
    patch_all(monkeypatch, "root_hwnd",
        lambda hwnd: hwnd,
    )
    patch_all(monkeypatch, "send_hotkey_ctrl_s",
        lambda: None,
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), scope_hwnd=12345, timeout=0.01)

    assert result["ok"] is False
    assert result["oracle"]["ok"] is False
    assert result["oracle"]["parent_new_state"]["ok"] is False
    assert "新增" in result["reason"]


def test_run_one_row_uses_detail_pipeline_verifier(monkeypatch):
    calls = {
        "start": 0,
        "field": [],
        "snapshot": [],
        "rows": [],
        "wait": [],
        "fill_header_kwargs": [],
        "body_locate_kwargs": [],
        "account_scope": [],
        "delete_extra_kwargs": [],
    }

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class FakeVerifier:
        def __init__(self, config, located, flow_started_at=None, **kwargs):
            self.config = config
            self.located = located
            self.flow_started_at = flow_started_at
            self.kwargs = kwargs

        def start(self):
            calls["start"] += 1

        def submit_field(self, row_index, field, business):
            task_id = f"field-{len(calls['field'])}"
            calls["field"].append(
                (row_index, field["name"], business[field["value_key"]])
            )
            return task_id

        def submit_snapshot(
            self, label, max_rows=5, timeout=1.2, interval=0.08, min_matches=0
        ):
            task_id = f"snapshot-{len(calls['snapshot'])}"
            calls["snapshot"].append((label, max_rows, min_matches))
            return task_id

        def submit_row_count(self, expected_rows, timeout=1.1, interval=0.06):
            task_id = f"rows-{len(calls['rows'])}"
            calls["rows"].append(expected_rows)
            return task_id

        def wait(self, task_ids, timeout=2.0):
            calls["wait"].append((list(task_ids), timeout))
            return {
                "ok": True,
                "submitted": list(task_ids),
                "done": len(task_ids),
                "snapshots": [
                    {
                        "id": "snapshot-0",
                        "ok": True,
                        "snapshot": {
                            "ok": True,
                            "rows": [
                                {
                                    "row_index": 0,
                                    "cells": {"6": "1", "7": "1090.00"},
                                }
                            ],
                        },
                    }
                ],
            }

        def close(self, timeout=1.0):
            calls["closed"] = timeout

    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 2002,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {
                            "path": "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0",
                            "name": "财务组织(O)",
                            "description": "财务组织(O)",
                        },
                    }
                ]
            },
        },
    )
    patch_all(monkeypatch, "JABOperator", FakeJAB)
    patch_all(monkeypatch, "fill_header",
        lambda _jab, _business, **kwargs: (
            calls["fill_header_kwargs"].append(kwargs)
            or [
                {
                    "ok": True,
                    "label": "客户",
                    "value": "YW03574",
                    "accepted_text": "ACME LTD",
                }
            ]
        ),
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **kwargs: (
            calls["body_locate_kwargs"].append(kwargs)
            or {
                "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
                "candidates": [],
            }
        ),
    )

    def fail_sync_read_before(*_args, **_kwargs):
        raise AssertionError("完整流程不应在明细写入前同步读整表")

    patch_all(monkeypatch, "read_body_table",
        fail_sync_read_before,
    )

    def fake_write_detail(_jab, business, _located, after_field=None, **_kwargs):
        field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}
        step = {"ok": True, "input_ok": True, "name": field["name"]}
        assert after_field is not None
        step["async_verify_task"] = after_field(0, field, business, step)
        return [step]

    patch_all(monkeypatch, "write_detail_line_by_screen",
        fake_write_detail,
    )
    patch_all(monkeypatch, "delete_extra_row_if_present",
        lambda *_args, **kwargs: (
            calls["delete_extra_kwargs"].append(kwargs) or {"ok": True}
        ),
    )
    patch_all(monkeypatch, "wait_header_account_description",
        lambda _jab, _timeout=0.0, scope=None: (
            calls["account_scope"].append(scope) or {"accepted": True}
        ),
    )
    patch_all(monkeypatch, "DetailPipelineVerifier", FakeVerifier
    )
    patch_all(monkeypatch, "recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("no-save 正常路径不应主动检查弹窗")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True, (
        report.get("failed_step"),
        report.get("reason"),
        report.get("exception"),
        report.get("detail_pipeline_repair"),
        report.get("detail_pipeline_verify_after_repair"),
    )
    assert calls["start"] == 1
    assert calls["field"] == [(0, "收款银行账户", "FTE1219165931831")]
    assert calls["snapshot"] == [("after-main-line", 3, 1)]
    assert calls["rows"] == [1]
    assert calls["delete_extra_kwargs"][0]["scope_hwnd"] == 2002
    assert calls["delete_extra_kwargs"][0]["defer_wait"] is True
    assert calls["wait"] == [(["field-0", "rows-0"], 2.0)]
    assert calls["fill_header_kwargs"][0]["scope_hwnd"] == 2002
    assert calls["fill_header_kwargs"][0]["dynamic_index"] == 5
    assert (
        calls["fill_header_kwargs"][0]["anchor_path"]
        == "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0"
    )
    assert calls["body_locate_kwargs"][0]["scope_hwnd"] == 2002
    assert calls["body_locate_kwargs"][0]["cached"] is None
    assert calls["account_scope"][0]["scope_hwnd"] == 2002
    assert calls["account_scope"][0]["dynamic_index"] == 5
    assert report["before_table"]["skipped"] is True
    assert report["after_table"]["skipped"] is True
    assert calls["closed"] == 0.2


def test_run_one_row_verifies_extra_text_fields_in_pipeline(monkeypatch):
    calls = {
        "text": [],
        "wait": [],
    }

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class FakeVerifier:
        def __init__(self, *_args, **_kwargs):
            pass

        def start(self):
            pass

        def submit_field(self, *_args, **_kwargs):
            return "field-0"

        def submit_snapshot(self, *_args, **_kwargs):
            return "snapshot-0"

        def submit_row_count(self, *_args, **_kwargs):
            return "rows-0"

        def submit_path_text(
            self,
            label,
            path,
            expected,
            scope_hwnd=None,
            *_args,
            **_kwargs,
        ):
            calls["text"].append(
                {
                    "label": label,
                    "path": path,
                    "expected": expected,
                    "scope_hwnd": scope_hwnd,
                }
            )
            return "text-0"

        def wait(self, task_ids, timeout=2.0):
            calls["wait"].append(list(task_ids))
            result = pipeline_wait_ok_with_cny_snapshot(task_ids)
            result.update({
                "ok": True,
                "submitted": list(task_ids),
                "done": len(task_ids),
                "results": {
                    "field-0": {"ok": True},
                    "rows-0": {"ok": True},
                    "text-0": {"ok": True, "type": "path_text"},
                },
            })
            return result

        def close(self, timeout=1.0):
            pass

    patch_all(monkeypatch, "JABOperator", FakeJAB)
    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    patch_all(monkeypatch, "fill_header",
        lambda _jab, _business, **_kwargs: [
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "accepted_text": "ACME LTD",
            }
        ],
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    patch_all(monkeypatch, "write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    patch_all(monkeypatch, "write_extra_text_fields",
        lambda *_args, **_kwargs: {
            "ok": True,
            "fields": [
                {
                    "ok": True,
                    "label": "商务领款备忘",
                    "value": "PI-001",
                    "path": "memo.path",
                }
            ],
        },
    )
    patch_all(monkeypatch, "delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    patch_all(monkeypatch, "wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    patch_all(monkeypatch, "verify_and_repair_header_targets",
        lambda *_args, **_kwargs: {"ok": True, "reads": [], "missing": []},
    )
    patch_all(monkeypatch, "DetailPipelineVerifier", FakeVerifier
    )

    report = run_one_row(
        {},
        plan_row(10, extra_text_fields={"商务领款备忘": "PI-001"}),
        save_enabled=False,
    )

    assert report["ok"] is True
    assert calls["text"] == [
        {
            "label": "商务领款备忘",
            "path": "memo.path",
            "expected": "PI-001",
            "scope_hwnd": 2002,
        }
    ]
    assert calls["wait"] == [["rows-0", "text-0"]]
    assert report["extra_text_verify_tasks"] == ["text-0"]


def test_write_extra_text_field_rewrites_when_first_paste_does_not_land(monkeypatch):
    descriptions = iter(["", "", "", "PI-001"])
    paste_values = []

    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description=next(descriptions))

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "get_receipt_header_path_template",
        lambda dynamic_index: {"text_suffix_template": "memo.{index}.0"},
    )
    patch_all(monkeypatch, "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 2002},
            "path": "memo.path",
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
            "source": "dynamic-path",
        },
    )

    def fake_paste(_jab, _vm_id, _context, _window, value):
        paste_values.append(value)
        return {"ok": True, "enter_ok": True}

    patch_all(monkeypatch, "guarded_paste_header_value", fake_paste)
    monkeypatch.setattr(full_flow.time, "sleep", lambda *_args, **_kwargs: None)

    result = write_extra_text_field_by_dynamic_path(
        FakeJAB(),
        "商务领款备忘",
        "PI-001",
        5,
        scope_hwnd=2002,
    )

    assert result["ok"] is True
    assert paste_values == ["PI-001", "PI-001"]
    assert len(result["attempts"]) == 2
    assert len(result["rewrites"]) == 1
    assert result["description_after"] == "PI-001"


def test_write_extra_text_field_can_defer_readback(monkeypatch):
    paste_values = []

    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            raise AssertionError("延迟校验模式不应立即读回")

        def get_text_context_value(self, _vm_id, _context):
            raise AssertionError("延迟校验模式不应立即读回")

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "get_receipt_header_path_template",
        lambda dynamic_index: {"text_suffix_template": "memo.{index}.0"},
    )
    patch_all(monkeypatch, "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 2002},
            "path": "memo.path",
            "source": "dynamic-path",
        },
    )

    def fake_paste(_jab, _vm_id, _context, _window, value):
        paste_values.append(value)
        return {"ok": True, "enter_ok": True}

    patch_all(monkeypatch, "guarded_paste_header_value", fake_paste)

    result = write_extra_text_field_by_dynamic_path(
        FakeJAB(),
        "商务领款备忘",
        "PI-001",
        5,
        scope_hwnd=2002,
        verify_after_write=False,
    )

    assert result["ok"] is True
    assert result["verify_after_write"] is False
    assert paste_values == ["PI-001"]
    assert result["attempts"][0]["post_write_snapshot"] == {}


def test_write_extra_text_field_fails_when_rewrite_still_does_not_land(monkeypatch):
    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description="")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "get_receipt_header_path_template",
        lambda dynamic_index: {"text_suffix_template": "memo.{index}.0"},
    )
    patch_all(monkeypatch, "find_receipt_header_field_by_dynamic_path",
        lambda *_args, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 2002},
            "path": "memo.path",
            "source": "dynamic-path",
        },
    )
    patch_all(monkeypatch, "guarded_paste_header_value",
        lambda *_args: {"ok": True, "enter_ok": True},
    )
    monkeypatch.setattr(full_flow.time, "sleep", lambda *_args, **_kwargs: None)

    result = write_extra_text_field_by_dynamic_path(
        FakeJAB(),
        "商务领款备忘",
        "PI-001",
        5,
        scope_hwnd=2002,
    )

    assert result["ok"] is False
    assert len(result["attempts"]) == 2
    assert result["reason"] == "写入后未读回目标文本字段值：商务领款备忘='PI-001'"


def test_header_unified_check_verifies_extra_text_fields(monkeypatch):
    states = {"customer.path": "ACME LTD", "memo.path": "PI-001"}
    paths = []

    class FakeJAB:
        def find_context_by_path_once(self, path, **_kwargs):
            paths.append(path)
            return path, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, context):
            return FakeInfo(description=states[str(context)])

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    result = verify_and_repair_header_targets(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "path": "customer.path",
                "accepted_text": "ACME LTD",
            }
        ],
        {
            "ok": True,
            "fields": [
                {
                    "ok": True,
                    "label": "商务领款备忘",
                    "value": "PI-001",
                    "path": "memo.path",
                }
            ],
        },
        5,
        2002,
        expectations={
            "客户": {
                "mode": "customer_name_similarity",
                "expected": "ACME LTD",
                "threshold": 80,
            },
            "商务领款备忘": {"mode": "text_exact", "expected": "PI-001"},
        },
    )

    assert result["ok"] is True
    assert [(item["label"], item["ok"]) for item in result["reads"]] == [
        ("客户", True),
        ("商务领款备忘", True),
    ]
    assert paths == ["customer.path", "memo.path"]
    assert any(target["kind"] == "extra_text" for target in result["targets"])


def test_header_unified_check_fails_when_header_repair_still_missing(monkeypatch):
    class FakeJAB:
        def find_context_by_path_once(self, path, **_kwargs):
            return path, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description="")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "guarded_paste_header_value",
        lambda *_args: {"ok": True, "enter_ok": True},
    )

    result = verify_and_repair_header_targets(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "结算方式",
                "value": "网银",
                "path": "settlement.path",
            }
        ],
        {
            "ok": True,
            "fields": [],
        },
        5,
        2002,
        expectations={"结算方式": {"mode": "text_exact", "expected": "网银"}},
    )

    assert result["ok"] is False
    assert "表头字段核验失败" in result["reason"]


def test_header_unified_check_skips_finance_org_path_read(monkeypatch):
    class FakeJAB:
        def find_context_by_path_once(self, path, **_kwargs):
            raise AssertionError(f"财务组织不应进入统一校验 exact path 读取: {path}")

    result = verify_and_repair_header_targets(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "财务组织",
                "value": "A001",
                "path": "finance.org.path",
                "accepted_text": "上海移为通信技术股份有限公司",
            }
        ],
        {"ok": True, "fields": []},
        5,
        2002,
        expectations={"币种": {"mode": "currency", "expected": "USD"}},
    )

    assert result["ok"] is True
    assert result["skipped"] is True
    assert result["targets"] == []


def test_header_currency_matches_code_and_chinese_name():
    assert header_currency_matches("USD", ["USD", "美元"], "美元") is True
    assert header_currency_matches("USD", ["USD", "美元"], "USD/美元") is True
    assert header_currency_matches("CNY", ["CNY", "人民币"], "人民币") is True
    assert header_currency_matches("CNY", ["CNY", "人民币"], "RMB 人民币") is True
    assert header_currency_matches("USD", ["USD", "美元"], "人民币") is False


def test_header_unified_check_accepts_currency_chinese_readback(monkeypatch):
    class FakeJAB:
        def find_context_by_path_once(self, path, **_kwargs):
            return path, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description="美元")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    result = verify_and_repair_header_targets(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "币种",
                "value": "USD",
                "path": "currency.path",
                "accepted_text": ["USD", "美元"],
            }
        ],
        {"ok": True, "fields": []},
        5,
        2002,
        expectations={"币种": {"mode": "currency", "expected": "USD"}},
    )

    assert result["ok"] is True
    assert result["reads"][0]["ok"] is True
    assert result["reads"][0]["actual_value"] == "美元"


def test_header_unified_check_rejects_wrong_date_after_two_repairs(monkeypatch):
    class FakeJAB:
        def find_context_by_path_once(self, path, **_kwargs):
            return path, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description="2026-06-30")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    patch_all(monkeypatch, "guarded_paste_header_value",
        lambda *_args: {"ok": True, "enter_ok": True},
    )

    result = verify_and_repair_header_targets(
        FakeJAB(),
        [
            {
                "ok": True,
                "label": "单据日期",
                "value": "2026-06-15",
                "path": "date.path",
            }
        ],
        {"ok": True, "fields": []},
        5,
        2002,
        expectations={"单据日期": {"mode": "date_exact", "expected": "2026-06-15"}},
    )

    assert result["ok"] is False
    assert len(result["repairs"]) == 2
    assert result["summary"]["failed"][0]["actual"] == "2026-06-30"


def test_customer_verify_uses_payer_name_similarity_threshold():
    assert (
        customer_name_similarity(
            "TECNOMOTUM SOCIEDAD ANONIMA PROMOT+",
            "TECNOMOTUM SOCIEDAD ANONIMA PROMOTO",
        )
        == 98
    )
    assert customer_name_similarity("ACME LTD", "移为") < 80


def test_build_header_verify_expectations_uses_payer_name_for_customer():
    row = plan_row(10, extra_text_fields={"商务领款备忘": "PI-001"})
    result = build_header_verify_expectations(
        row,
        business_from_plan_row(row),
        row.extra_text_fields,
    )

    assert result["客户"]["expected"] == "ACME LTD"
    assert result["客户"]["threshold"] == 80
    assert result["商务领款备忘"]["expected"] == "PI-001"


def test_ensure_header_counterparty_customer_skips_when_already_customer(monkeypatch):
    class FakeJAB:
        def __init__(self):
            self.actions = []

        def find_context_by_path_once(self, path, **_kwargs):
            return object(), 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeComboInfo(description="客户")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def do_action(self, *_args, **_kwargs):
            self.actions.append(_kwargs.get("action_name"))
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    jab = FakeJAB()
    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "detail-row0-col0",
            "row": 0,
            "col": 0,
            "value": "客户",
            "text": "客户",
        },
    )

    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(jab, 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["skipped"] is True
    assert result["actual"] == "客户"
    assert jab.actions == []


def test_ensure_header_counterparty_customer_skips_when_embedded_selected_customer(monkeypatch):
    class FakeNode:
        def __init__(
            self,
            role,
            name="",
            description="",
            states="enabled,visible,showing",
            children=None,
            bounds=(0, 0, 10, 10),
        ):
            self.name = name
            self.description = description
            self.role = role
            self.role_en_US = role
            self.states = states
            self.states_en_US = states
            self.children = children or []
            self.childrenCount = len(self.children)
            self.x, self.y, self.width, self.height = bounds

    class FakeJAB:
        def __init__(self):
            self.actions = []
            self.customer = FakeNode(
                "label",
                name="客户",
                states="enabled,focusable,visible,opaque,selectable,selected,showing",
                bounds=(-31998, -31998, 196, 20),
            )
            self.department = FakeNode(
                "label",
                name="部门",
                states="enabled,focusable,visible,opaque,selectable,showing",
                bounds=(-31998, -31978, 196, 20),
            )
            self.list_node = FakeNode(
                "list",
                description="客户",
                states="enabled,visible,showing,opaque",
                children=[self.customer, self.department],
                bounds=(-31998, -31998, 196, 80),
            )
            self.popup = FakeNode(
                "popup menu",
                states="enabled,visible,showing,selectable,selected",
                children=[self.list_node],
                bounds=(-32000, -32000, 200, 84),
            )
            self.combo = FakeNode(
                "combo box",
                states="enabled,focusable,visible,showing,opaque,expanded",
                children=[self.popup],
                bounds=(998, 232, 198, 22),
            )
            self.dll = self

        def find_context_by_path_once(self, path, **_kwargs):
            return self.combo, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return _context

        def get_text_context_value(self, _vm_id, _context):
            return "翸" if _context is self.combo else ""

        def getAccessibleChildFromContext(self, _vm_id, context, index):
            return context.children[index]

        def do_action(self, *_args, **_kwargs):
            self.actions.append(_kwargs.get("action_name"))
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    jab = FakeJAB()
    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "detail-row0-col0",
            "row": 0,
            "col": 0,
            "value": "客户",
            "text": "客户",
        },
    )

    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(jab, 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["skipped"] is True
    assert result["actual"] == "客户"
    assert result["source"] == "detail-row0-col0"
    assert jab.actions == []


def test_ensure_header_counterparty_customer_repairs_conflict_with_selection_api(
    monkeypatch,
):
    class FakeNode:
        def __init__(
            self,
            role,
            name="",
            description="",
            states="enabled,visible,showing",
            children=None,
            bounds=(0, 0, 10, 10),
        ):
            self.name = name
            self.description = description
            self.role = role
            self.role_en_US = role
            self.states = states
            self.states_en_US = states
            self.children = children or []
            self.childrenCount = len(self.children)
            self.x, self.y, self.width, self.height = bounds

    class FakeJAB:
        def __init__(self):
            self.customer = FakeNode(
                "label",
                name="客户",
                states="enabled,focusable,visible,opaque,selectable,selected,showing",
            )
            self.department = FakeNode(
                "label",
                name="部门",
                states="enabled,focusable,visible,opaque,selectable,showing",
            )
            self.list_node = FakeNode(
                "list",
                description="客户",
                children=[self.customer, self.department],
            )
            self.popup = FakeNode("popup menu", children=[self.list_node])
            self.combo = FakeNode("combo box", children=[self.popup])
            self.dll = self
            self.selection_calls = []
            self.keys = []

        def find_context_by_path_once(self, path, **_kwargs):
            return self.combo, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return _context

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def getAccessibleChildFromContext(self, _vm_id, context, index):
            return context.children[index]

        def clearAccessibleSelectionFromContext(self, vm_id, context):
            self.selection_calls.append(("clear", vm_id, context))
            return True

        def addAccessibleSelectionFromContext(self, vm_id, context, index):
            self.selection_calls.append(("add", vm_id, context, index))
            return True

        def requestFocus(self, vm_id, context):
            self.selection_calls.append(("focus", vm_id, context))
            return True

        def press_key(self, key, wait=0):
            self.keys.append((key, wait))
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    reads = iter(
        [
            {
                "ok": True,
                "source": "detail-row0-col0",
                "row": 0,
                "col": 0,
                "value": "部门",
                "text": "部门",
            },
            {
                "ok": True,
                "source": "detail-row0-col0",
                "row": 0,
                "col": 0,
                "value": "客户",
                "text": "客户",
            },
        ]
    )
    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: next(reads),
    )

    jab = FakeJAB()
    monkeypatch.setattr(cp, "find_counterparty_combo",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "nearby",
            "context": jab.combo,
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 2002},
            "path": "0.0.1.0.0.0.0.5.nearby.suffix",
        },
    )
    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(jab, 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["repaired"] is True
    assert result["repaired_from_conflict"] is True
    assert result["actual"] == "客户"
    assert result["detail"]["value"] == "部门"
    assert result["after_detail"]["value"] == "客户"
    assert ("add", 1, jab.list_node, 0) in jab.selection_calls
    assert jab.keys == [("home", 0.02), ("enter", 0)]


def test_ensure_header_counterparty_customer_repairs_blank_detail_with_embedded_selection(
    monkeypatch,
):
    class FakeNode:
        def __init__(
            self,
            role,
            name="",
            description="",
            states="enabled,visible,showing",
            children=None,
            bounds=(0, 0, 10, 10),
        ):
            self.name = name
            self.description = description
            self.role = role
            self.role_en_US = role
            self.states = states
            self.states_en_US = states
            self.children = children or []
            self.childrenCount = len(self.children)
            self.x, self.y, self.width, self.height = bounds

    class FakeJAB:
        def __init__(self):
            self.customer = FakeNode(
                "label",
                name="客户",
                states="enabled,focusable,visible,opaque,selectable,selected,showing",
            )
            self.department = FakeNode(
                "label",
                name="部门",
                states="enabled,focusable,visible,opaque,selectable,showing",
            )
            self.list_node = FakeNode(
                "list",
                description="客户",
                children=[self.customer, self.department],
            )
            self.popup = FakeNode("popup menu", children=[self.list_node])
            self.combo = FakeNode("combo box", children=[self.popup])
            self.dll = self
            self.selection_calls = []
            self.keys = []

        def find_context_by_path_once(self, path, **_kwargs):
            return self.combo, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return _context

        def get_text_context_value(self, _vm_id, _context):
            return "翸" if _context is self.combo else ""

        def getAccessibleChildFromContext(self, _vm_id, context, index):
            return context.children[index]

        def clearAccessibleSelectionFromContext(self, vm_id, context):
            self.selection_calls.append(("clear", vm_id, context))
            return True

        def addAccessibleSelectionFromContext(self, vm_id, context, index):
            self.selection_calls.append(("add", vm_id, context, index))
            return True

        def requestFocus(self, vm_id, context):
            self.selection_calls.append(("focus", vm_id, context))
            return True

        def press_key(self, key, wait=0):
            self.keys.append((key, wait))
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    reads = iter(
        [
            {
                "ok": False,
                "source": "detail-row0-col0",
                "row": 0,
                "col": 0,
                "value": "",
                "text": "翸",
                "reason": "明细表往来对象单元格为空",
            },
            {
                "ok": True,
                "source": "detail-row0-col0",
                "row": 0,
                "col": 0,
                "value": "客户",
                "text": "客户",
            },
        ]
    )
    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: next(reads),
    )
    monkeypatch.setattr(full_flow.time, "sleep", lambda *_args, **_kwargs: None)

    jab = FakeJAB()
    monkeypatch.setattr(cp, "find_counterparty_combo",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "nearby",
            "context": jab.combo,
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 2002},
            "path": "0.0.1.0.0.0.0.5.nearby.suffix",
        },
    )
    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(jab, 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["repaired"] is True
    assert result["source"] == "embedded-selection-api"
    assert result["after_detail"]["value"] == "客户"
    assert ("add", 1, jab.list_node, 0) in jab.selection_calls
    assert jab.keys == [("home", 0.02), ("enter", 0)]


def test_ensure_header_counterparty_customer_fails_when_conflict_repair_not_confirmed(
    monkeypatch,
):
    class FakeNode:
        def __init__(self, role, name="", description="", states="enabled,visible,showing", children=None):
            self.name = name
            self.description = description
            self.role = role
            self.role_en_US = role
            self.states = states
            self.states_en_US = states
            self.children = children or []
            self.childrenCount = len(self.children)
            self.x, self.y, self.width, self.height = (0, 0, 10, 10)

    class FakeJAB:
        def __init__(self):
            self.customer = FakeNode(
                "label",
                name="客户",
                states="enabled,focusable,visible,opaque,selectable,selected,showing",
            )
            self.supplier = FakeNode(
                "label",
                name="供应商",
                states="enabled,focusable,visible,opaque,selectable,showing",
            )
            self.list_node = FakeNode(
                "list",
                description="客户",
                children=[self.customer, self.supplier],
            )
            self.popup = FakeNode("popup menu", children=[self.list_node])
            self.combo = FakeNode("combo box", children=[self.popup])
            self.dll = self
            self.selection_calls = []

        def find_context_by_path_once(self, path, **_kwargs):
            return self.combo, 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return _context

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def getAccessibleChildFromContext(self, _vm_id, context, index):
            return context.children[index]

        def clearAccessibleSelectionFromContext(self, vm_id, context):
            self.selection_calls.append(("clear", vm_id, context))
            return True

        def addAccessibleSelectionFromContext(self, vm_id, context, index):
            self.selection_calls.append(("add", vm_id, context, index))
            return True

        def requestFocus(self, vm_id, context):
            self.selection_calls.append(("focus", vm_id, context))
            return True

        def press_key(self, key, wait=0):
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    reads = iter(
        [
            {
                "ok": True,
                "source": "detail-row0-col0",
                "row": 0,
                "col": 0,
                "value": "供应商",
                "text": "供应商",
            },
            {
                "ok": True,
                "source": "detail-row0-col0",
                "row": 0,
                "col": 0,
                "value": "供应商",
                "text": "供应商",
            },
        ]
    )
    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: next(reads),
    )

    jab = FakeJAB()
    monkeypatch.setattr(cp, "find_counterparty_combo",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "nearby",
            "context": jab.combo,
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 2002},
            "path": "0.0.1.0.0.0.0.5.nearby.suffix",
        },
    )
    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(jab, 5, scope_hwnd=2002)

    assert result["ok"] is False
    assert result["actual"] == "供应商"
    assert result["state"]["state"] == "repairable-conflict"
    assert result["after_detail"]["value"] == "供应商"
    assert ("add", 1, jab.list_node, 0) in jab.selection_calls


def test_ensure_header_counterparty_customer_skips_when_detail_row0_col0_customer(monkeypatch):
    class FakeJAB:
        def __init__(self):
            self.actions = []

        def find_context_by_path_once(self, path, **_kwargs):
            return object(), 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeComboInfo(description="")

        def get_text_context_value(self, _vm_id, _context):
            return "翸"

        def do_action(self, *_args, **_kwargs):
            self.actions.append(_kwargs.get("action_name"))
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "detail-row0-col0",
            "row": 0,
            "col": 0,
            "value": "客户",
            "text": "客户",
            "is_selected": True,
        },
    )

    jab = FakeJAB()

    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(jab, 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["skipped"] is True
    assert result["actual"] == "客户"
    assert result["source"] == "detail-row0-col0"
    assert result["detail"]["value"] == "客户"
    assert jab.actions == []


def test_ensure_header_counterparty_customer_skips_lower_detail_without_header_lookup(
    monkeypatch,
):
    class FakeJAB:
        def __init__(self):
            self.actions = []

        def get_context_info(self, _vm_id, _context):
            return FakeComboInfo(description="客户")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def do_action(self, *_args, **_kwargs):
            self.actions.append(_kwargs.get("action_name"))
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(cp, "_COUNTERPARTY_NEARBY_SUFFIX_CACHE", {})
    monkeypatch.setattr(cp, "find_counterparty_combo",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("lower detail 已是客户时不应定位上方往来对象")
        ),
    )
    patch_all(monkeypatch, "receipt_header_dynamic_prefix",
        lambda dynamic_index: "0.0.1.0.0.0.0.5",
    )
    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "detail-row0-col0",
            "row": 0,
            "col": 0,
            "value": "客户",
            "text": "客户",
        },
    )

    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(FakeJAB(), 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["path"] is None
    assert cp._COUNTERPARTY_NEARBY_SUFFIX_CACHE == {}


def test_find_counterparty_combo_uses_cached_nearby_suffix(monkeypatch):
    calls = []

    class FakeJAB:
        pass

    monkeypatch.setattr(cp, "_COUNTERPARTY_NEARBY_SUFFIX_CACHE", {})
    patch_all(monkeypatch, "receipt_header_dynamic_prefix",
        lambda dynamic_index: "0.0.1.0.0.0.0.5",
    )
    cp._COUNTERPARTY_NEARBY_SUFFIX_CACHE[2002] = {
        "suffix": "nearby.suffix",
        "path": "old",
    }

    def fake_find_by_path(_jab, path, **_kwargs):
        calls.append(path)
        return {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [],
            "window": {"hwnd": 2002},
            "path": path,
        }

    monkeypatch.setattr(cp, "find_counterparty_combo_by_path", fake_find_by_path)
    monkeypatch.setattr(cp, "find_counterparty_combo_nearby",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("nearby should not run when cached path works")
        ),
    )

    result = cp.find_counterparty_combo(FakeJAB(), 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["source"] == "nearby-cache-path"
    assert calls == ["0.0.1.0.0.0.0.5.nearby.suffix"]


def test_find_counterparty_combo_nearby_handles_label_path_triples(monkeypatch):
    class Node:
        def __init__(
            self,
            role,
            name="",
            description="",
            children=None,
            bounds=(0, 0, 10, 10),
            states="enabled,visible,showing",
        ):
            self.role = role
            self.role_en_US = role
            self.name = name
            self.description = description
            self.states = states
            self.states_en_US = states
            self.children = children or []
            self.childrenCount = len(self.children)
            self.x, self.y, self.width, self.height = bounds

    class FakeDLL:
        def __init__(self, nodes):
            self.nodes = nodes

        def isJavaWindow(self, _hwnd):
            return True

        def getAccessibleContextFromHWND(self, _hwnd, vm_id_ref, root_context_ref):
            vm_id_ref._obj.value = 1
            root_context_ref._obj.value = 100
            return True

        def getAccessibleChildFromContext(self, _vm_id, context, index):
            return self.nodes[int(context)].children[index]

    class FakeJAB:
        def __init__(self):
            self.nodes = {
                100: Node("panel", children=[101, 102]),
                101: Node("label", name="往来对象", bounds=(914, 232, 74, 22)),
                102: Node(
                    "combo box",
                    children=[],
                    bounds=(998, 232, 198, 22),
                    states="enabled,focusable,visible,showing,opaque,collapsed",
                ),
            }
            self.dll = FakeDLL(self.nodes)
            self.max_depth = 8
            self.max_children = 120
            self.released = []

        def get_context_info(self, _vm_id, context):
            return self.nodes[int(context)]

        def context_info_has_valid_bounds(self, info):
            return info.width > 0 and info.height > 0 and info.x >= 0 and info.y >= 0

        def get_action_names(self, _vm_id, context):
            return ["togglePopup"] if int(context) == 102 else []

        def info_to_dict(self, info):
            return {
                "name": info.name,
                "description": info.description,
                "role": info.role,
                "states": info.states,
                "x": info.x,
                "y": info.y,
                "width": info.width,
                "height": info.height,
            }

        def release_contexts(self, _vm_id, contexts):
            self.released.extend(contexts or [])

    patch_all(monkeypatch, "receipt_header_dynamic_prefix",
        lambda dynamic_index: "0.0.1.0.0.0.0.2",
    )

    result = cp.find_counterparty_combo_nearby(
        FakeJAB(),
        2,
        scope_hwnd=2002,
    )

    assert result["ok"] is True
    assert result["source"] == "nearby"
    assert result["path"] == "0.1"
    assert result["target"]["label"]["name"] == "往来对象"


def test_ensure_header_counterparty_customer_fails_when_existing_non_target_value(monkeypatch):
    class FakeJAB:
        def __init__(self):
            self.actions = []

        def find_context_by_path_once(self, path, **_kwargs):
            return object(), 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeComboInfo(description="供应商")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def press_key(self, key, wait=0):
            self.actions.append((key, wait))
            return True

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "detail-row0-col0",
            "row": 0,
            "col": 0,
            "value": "供应商",
            "text": "供应商",
        },
    )

    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(FakeJAB(), 5, scope_hwnd=2002)

    assert result["ok"] is False
    assert result["actual"] == "供应商"
    assert result["detail"]["value"] == "供应商"


def test_ensure_header_counterparty_customer_trusts_detail_over_stale_header_text(
    monkeypatch,
):
    class FakeJAB:
        def find_context_by_path_once(self, path, **_kwargs):
            return object(), 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeComboInfo(description="供应商")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: {
            "ok": True,
            "source": "detail-row0-col0",
            "row": 0,
            "col": 0,
            "value": "客户",
            "text": "客户",
        },
    )

    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(FakeJAB(), 5, scope_hwnd=2002)

    assert result["ok"] is True
    assert result["actual"] == "客户"
    assert result["source"] == "detail-row0-col0"
    assert result["detail"]["value"] == "客户"


def test_ensure_header_counterparty_customer_failure_includes_detail_diagnostic(
    monkeypatch,
):
    class FakeJAB:
        def find_context_by_path_once(self, path, **_kwargs):
            return object(), 1, [], {"hwnd": 2002}

        def get_context_info(self, _vm_id, _context):
            return FakeComboInfo(description="")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(cp, "read_detail_counterparty_value",
        lambda *_args, **_kwargs: {"ok": False, "value": "", "text": "", "reason": "明细表 path 未定位"},
    )
    result = REAL_ENSURE_HEADER_COUNTERPARTY_CUSTOMER(FakeJAB(), 5, scope_hwnd=2002)

    assert result["ok"] is False
    assert result["actual"] == ""
    assert result["detail"]["reason"] == "明细表 path 未定位"
    assert "header_selected=" in result["reason"]
    assert "combo_text=" in result["reason"]
    assert "detail_row0_col0=" in result["reason"]
    assert "detail_reason=明细表 path 未定位" in result["reason"]


def test_run_one_row_resolves_header_scope_by_finance_org_fast_path(monkeypatch):
    calls = {
        "finance_scope": [],
        "fill_header_kwargs": [],
        "fill_header_scope_cache": [],
        "body_locate_kwargs": [],
    }

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class FakeVerifier:
        def __init__(self, *_args, **_kwargs):
            pass

        def start(self):
            pass

        def submit_field(self, *_args, **_kwargs):
            return "field-0"

        def submit_snapshot(self, *_args, **_kwargs):
            return "snapshot-0"

        def submit_row_count(self, *_args, **_kwargs):
            return "rows-0"

        def wait(self, task_ids=(), *_args, **_kwargs):
            return pipeline_wait_ok_with_cny_snapshot(task_ids)

        def close(self, timeout=1.0):
            pass

    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 919586,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {"path": "not-a-header-path"},
                    }
                ]
            },
        },
    )
    patch_all(monkeypatch, "JABOperator", FakeJAB)
    patch_all(monkeypatch, "find_finance_org_header_scope_by_paths",
        lambda _jab, hwnd, **kwargs: (
            calls["finance_scope"].append((hwnd, kwargs))
            or {
                "ok": True,
                "scope_hwnd": hwnd,
                "dynamic_index": 5,
                "dynamic_prefix": "0.0.1.0.0.0.0.5",
                "semantic_label_path": "0.fast.semantic.label",
                "label_path": "0.fast.visible.label",
                "text_path": "0.fast.text",
                "variant": "observed-compact",
            }
        ),
    )
    patch_all(monkeypatch, "fill_header",
        lambda jab, _business, **kwargs: (
            calls["fill_header_kwargs"].append(kwargs)
            or calls["fill_header_scope_cache"].append(
                getattr(jab, "_receipt_header_scope_cache", None)
            )
            or [{"ok": True, "label": "客户", "accepted_text": "ACME LTD"}]
        ),
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **kwargs: (
            calls["body_locate_kwargs"].append(kwargs)
            or {
                "best": {
                    "path": "0.1",
                    "row_count": 1,
                    "col_count": 25,
                    "window": {"hwnd": 919586},
                },
                "candidates": [],
            }
        ),
    )
    patch_all(monkeypatch, "write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    patch_all(monkeypatch, "delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    patch_all(monkeypatch, "wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    patch_all(monkeypatch, "verify_and_repair_header_targets",
        lambda *_args, **_kwargs: {"ok": True, "reads": [], "missing": []},
    )
    patch_all(monkeypatch, "DetailPipelineVerifier", FakeVerifier
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True, report
    assert calls["finance_scope"] == [
        (
            919586,
            {"preferred_dynamic_index": None},
        )
    ]
    assert report["entry_dynamic_index"] == 5
    assert report["entry_finance_org_fast_scope"]["ok"] is True
    assert report["entry_dynamic_index_source"] == "finance-org-fast-scope"
    assert calls["fill_header_kwargs"][0]["scope_hwnd"] == 919586
    assert calls["fill_header_kwargs"][0]["dynamic_index"] == 5
    assert calls["fill_header_scope_cache"][0]["semantic_label_path"] == (
        "0.fast.semantic.label"
    )
    assert calls["fill_header_scope_cache"][0]["label_path"] == (
        "0.fast.visible.label"
    )
    assert calls["body_locate_kwargs"][0]["scope_hwnd"] == 919586
    assert calls["body_locate_kwargs"][0]["cached"] is None


def test_header_scope_cache_can_be_shared_between_row_jab_instances():
    class FakeJAB:
        pass

    shared_cache = {}
    first_jab = FakeJAB()
    scope = {
        "ok": True,
        "scope_hwnd": 919586,
        "mode": "header-anchor-retry-current-canvas",
        "dynamic_index": 5,
        "dynamic_prefix": "0.0.1.0.0.0.0.5",
        "label_path": "0.0.1.0.0.0.0.5.0.0",
    }

    cache_receipt_header_scope(first_jab, shared_cache, scope)

    second_jab = FakeJAB()
    assert not hasattr(second_jab, "_receipt_header_scope_cache")
    assert shared_cache["ok"] is True
    assert shared_cache["scope_hwnd"] == 919586
    assert shared_cache["dynamic_index"] == 5
    assert first_jab._receipt_header_scope_cache == shared_cache


def test_run_one_row_stops_when_current_canvas_header_scope_missing(monkeypatch):
    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 919586,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {"path": "not-a-header-path"},
                    }
                ]
            },
        },
    )
    patch_all(monkeypatch, "JABOperator", FakeJAB)
    patch_all(monkeypatch, "find_finance_org_header_scope_by_paths",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "当前 canvas 未通过财务组织(O) dynamic path 扫描",
        },
    )
    patch_all(monkeypatch, "fill_header",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("锚点失败后不应进入表头写入")
        ),
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("锚点失败后不应定位明细表")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is False
    assert report["failed_step"] == "header-anchor"
    assert "不走语义兜底" in report["reason"]


def test_run_one_row_repairs_pending_detail_field_with_cached_path(monkeypatch):
    calls = {
        "field": [],
        "snapshot": [],
        "rows": [],
        "wait": [],
        "repair": [],
        "locate": 0,
    }
    located = {
        "best": {
            "path": "0.1",
            "row_count": 1,
            "col_count": 25,
            "window": {"hwnd": 2002},
        },
        "candidates": [],
    }

    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class RepairingVerifier:
        def __init__(self, config, actual_located, flow_started_at=None, **kwargs):
            self.config = config
            self.located = actual_located
            self.flow_started_at = flow_started_at
            self.kwargs = kwargs

        def start(self):
            pass

        def submit_field(self, row_index, field, business):
            task_id = f"field-{len(calls['field'])}"
            calls["field"].append(
                {
                    "task_id": task_id,
                    "row_index": row_index,
                    "name": field["name"],
                    "value": business[field["value_key"]],
                }
            )
            return task_id

        def submit_snapshot(
            self, label, max_rows=5, timeout=1.2, interval=0.08, min_matches=0
        ):
            task_id = f"snapshot-{len(calls['snapshot'])}"
            calls["snapshot"].append(
                {
                    "task_id": task_id,
                    "label": label,
                    "max_rows": max_rows,
                    "min_matches": min_matches,
                }
            )
            return task_id

        def submit_row_count(self, expected_rows, timeout=1.1, interval=0.06):
            task_id = f"rows-{len(calls['rows'])}"
            calls["rows"].append(expected_rows)
            return task_id

        def wait(self, task_ids, timeout=2.0):
            ids = list(task_ids)
            calls["wait"].append((ids, timeout))
            if len(calls["wait"]) == 1:
                return {
                    "ok": False,
                    "submitted": ["field-0", "rows-0"],
                    "done": 1,
                    "pending": 1,
                    "failed": [],
                    "results": {
                        "rows-0": {
                            "ok": True,
                            "type": "row_count",
                            "expected_rows": 1,
                            "actual_rows": 1,
                        }
                    },
                }
            result = pipeline_wait_ok_with_cny_snapshot(["field-0", "rows-0", "field-1"])
            result.update({
                "ok": True,
                "submitted": ["field-0", "rows-0", "field-1"],
                "done": 3,
                "pending": 0,
                "failed": [],
                "results": {
                    "rows-0": {"ok": True, "type": "row_count"},
                    "field-1": {
                        "ok": True,
                        "type": "field",
                        "name": "收款银行账户",
                    },
                },
            })
            return result

        def snapshot(self):
            return {"ok": len(calls["wait"]) >= 2}

        def close(self, timeout=1.0):
            calls["closed"] = timeout

    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: {
            "ok": True,
            "entry_state": {
                "hits": [
                    {
                        "window": {
                            "hwnd": 2002,
                            "class_name": "SunAwtCanvas",
                            "visible": True,
                        },
                        "control": {
                            "path": "0.0.1.0.0.0.0.5.0.0",
                            "dynamic_index": 5,
                        },
                    }
                ]
            },
        },
    )
    patch_all(monkeypatch, "JABOperator", FakeJAB)
    patch_all(monkeypatch, "fill_header",
        lambda *_args, **_kwargs: [
            {"ok": True, "label": "客户", "accepted_text": "ACME LTD"}
        ],
    )

    def fake_locate(*_args, **_kwargs):
        calls["locate"] += 1
        return located

    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        fake_locate,
    )
    patch_all(monkeypatch, "read_body_table",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("修复成功后不应整表读 fallback")
        ),
    )

    def fake_write_detail(_jab, business, actual_located, after_field=None, **_kwargs):
        assert actual_located["best"]["path"] == located["best"]["path"]
        assert actual_located["best"]["window"] == located["best"]["window"]
        field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}
        assert after_field is not None
        return [
            {
                "ok": True,
                "input_ok": True,
                "name": field["name"],
                "async_verify_task": after_field(
                    0,
                    field,
                    business,
                    {"ok": True, "name": field["name"]},
                ),
            }
        ]

    def fake_write_field_once(
        _jab,
        actual_located,
        table_window,
        row_index,
        field,
        business,
        attempt_no,
        current_col=None,
        recover_after_failure=None,
    ):
        calls["repair"].append(
            {
                "path": actual_located["best"]["path"],
                "table_window": table_window,
                "row_index": row_index,
                "field": field["name"],
                "value": business[field["value_key"]],
                "attempt_no": attempt_no,
                "current_col": current_col,
                "has_recover_hook": recover_after_failure is not None,
            }
        )
        return {
            "ok": True,
            "input_ok": True,
            "commit_ok": True,
            "commit_col": field["col"],
            "target": {"row": row_index, "col": field["col"]},
        }

    patch_all(monkeypatch, "write_detail_line_by_screen",
        fake_write_detail,
    )
    patch_all(monkeypatch, "write_field_once",
        fake_write_field_once,
    )
    patch_all(monkeypatch, "delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    patch_all(monkeypatch, "wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    patch_all(monkeypatch, "DetailPipelineVerifier", RepairingVerifier
    )
    patch_all(monkeypatch, "recover_cancelable_modal_now",
        lambda *_args, **_kwargs: {"ok": False, "attempted": False},
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True, (
        report.get("failed_step"),
        report.get("reason"),
        report.get("exception"),
        report.get("detail_pipeline_repair"),
        report.get("detail_pipeline_verify_after_repair"),
    )
    assert calls["locate"] == 1
    assert calls["field"] == [
        {
            "task_id": "field-0",
            "row_index": 0,
            "name": "收款银行账户",
            "value": "FTE1219165931831",
        },
        {
            "task_id": "field-1",
            "row_index": 0,
            "name": "收款银行账户",
            "value": "FTE1219165931831",
        },
    ]
    assert calls["snapshot"] == [
        {
            "task_id": "snapshot-0",
            "label": "after-main-line",
            "max_rows": 3,
            "min_matches": 1,
        },
        {
            "task_id": "snapshot-1",
            "label": "after-detail-repair",
            "max_rows": 3,
            "min_matches": 1,
        },
    ]
    assert calls["rows"] == [1]
    assert calls["wait"] == [
        (["field-0", "rows-0"], 2.0),
        (["field-1"], 2.0),
    ]
    assert calls["repair"] == [
        {
            "path": "0.1",
            "table_window": {"hwnd": 2002},
            "row_index": 0,
            "field": "收款银行账户",
            "value": "FTE1219165931831",
            "attempt_no": 2,
            "current_col": None,
            "has_recover_hook": True,
        }
    ]
    assert report["detail_pipeline_repair"]["ok"] is True
    assert report["detail_pipeline_verify_after_repair"]["ok"] is True
    assert report["after_table"]["skipped"] is True


def test_extract_entry_dynamic_index_from_entry_button_path():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0",
                    }
                }
            ]
        }
    }

    assert extract_entry_dynamic_index(report) == 5


def test_extract_entry_dynamic_index_from_anchor_hit():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.5.0.0.0.1",
                        "dynamic_index": 7,
                    }
                }
            ]
        }
    }

    assert extract_entry_dynamic_index(report) == 7


def test_extract_entry_dynamic_index_prefers_customer_corrected_anchor_index():
    report = {
        "entry_state": {
            "hits": [
                {
                    "control": {
                        "path": "0.0.1.0.0.0.0.3.0.0.0.1",
                        "dynamic_index": 5,
                        "dynamic_prefix": "0.0.1.0.0.0.0.5",
                    }
                }
            ]
        },
        "anchor": {
            "mode": "current-canvas-anchor-corrected-by-customer",
            "initial_dynamic_index": 3,
            "dynamic_index": 5,
        },
    }

    assert extract_entry_dynamic_index(report) == 5


def test_extract_entry_scope_hwnd_from_windows_after_choose():
    report = {
        "parsed": {
            "windows_after_choose": [
                {
                    "is_java": True,
                    "visible": True,
                    "hwnd": 24680,
                    "class_name": "SunAwtCanvas",
                },
            ]
        }
    }

    assert extract_entry_scope_hwnd(report) == 24680


def test_extract_entry_scope_hwnd_prefers_self_made_canvas():
    frame_hwnd = 1001
    canvas_hwnd = 2002
    report = {
        "entry_state": {
            "hits": [
                {
                    "window": {
                        "hwnd": frame_hwnd,
                        "class_name": "SunAwtFrame",
                        "visible": True,
                    },
                    "control": {"path": "0.0.0.0.1.0.0.0.0.3.0.0"},
                },
                {
                    "window": {
                        "hwnd": canvas_hwnd,
                        "class_name": "SunAwtCanvas",
                        "visible": True,
                    },
                    "control": {"path": "0.0.1.0.0.0.0.3.0.0"},
                },
            ]
        }
    }

    assert extract_entry_scope_hwnd(report) == canvas_hwnd


def test_run_one_row_recovers_modal_only_after_save_failure(monkeypatch):
    calls = {"save": 0, "recover": 0}

    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class LocalFakeVerifier:
        def __init__(self, *_args, **_kwargs):
            pass

        def start(self):
            pass

        def submit_snapshot(self, *args, **kwargs):
            return "snapshot-0"

        def submit_row_count(self, *args, **kwargs):
            return "rows-0"

        def wait(self, task_ids=(), *args, **kwargs):
            return pipeline_wait_ok_with_cny_snapshot(task_ids)

        def close(self, timeout=1.0):
            pass

    patch_all(monkeypatch, "JABOperator", LocalFakeJAB)
    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    patch_all(monkeypatch, "fill_header",
        lambda *_args, **_kwargs: [
            {"ok": True, "label": "客户", "accepted_text": "ACME LTD"}
        ],
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    patch_all(monkeypatch, "write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    patch_all(monkeypatch, "delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    patch_all(monkeypatch, "wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    patch_all(monkeypatch, "verify_and_repair_header_targets",
        lambda *_args, **_kwargs: {"ok": True, "reads": [], "missing": []},
    )
    patch_all(monkeypatch, "DetailPipelineVerifier", LocalFakeVerifier
    )

    def fake_save(*_args, **_kwargs):
        calls["save"] += 1
        if calls["save"] == 1:
            return {"ok": False, "reason": "前台窗口不是目标 NC 窗口"}
        return {"ok": True, "triggered": True}

    patch_all(monkeypatch, "save_receipt_by_ctrl_s", fake_save
    )
    patch_all(monkeypatch, "recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (
            calls.__setitem__("recover", calls["recover"] + 1)
            or {"ok": True, "attempted": True}
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=True)

    assert report["ok"] is True
    assert calls == {"save": 2, "recover": 1}
    assert report["save"]["retried_after_modal_recovery"] is True


def test_run_one_row_stops_when_customer_name_readback_is_empty(monkeypatch):
    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class LocalFakeVerifier:
        def __init__(self, config, located, flow_started_at=None, **_kwargs):
            self.config = config
            self.located = located
            self.flow_started_at = flow_started_at

        def start(self):
            pass

        def submit_snapshot(self, *args, **kwargs):
            return "snapshot-0"

        def submit_row_count(self, *args, **kwargs):
            return "rows-0"

        def wait(self, *args, **kwargs):
            return {"ok": True}

        def close(self, timeout=1.0):
            pass

    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    patch_all(monkeypatch, "JABOperator", LocalFakeJAB)
    patch_all(monkeypatch, "fill_header",
        lambda _jab, _business, **_kwargs: [
            {"ok": True, "label": "客户", "value": "YW03574"}
        ],
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    patch_all(monkeypatch, "read_body_table",
        lambda _jab, step: {"ok": True, "step": step, "rows": []},
    )
    patch_all(monkeypatch, "write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    patch_all(monkeypatch, "delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    patch_all(monkeypatch, "wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    patch_all(monkeypatch, "DetailPipelineVerifier", LocalFakeVerifier
    )
    patch_all(monkeypatch, "recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("no-save 正常路径不应主动检查弹窗")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is False
    assert report["failed_step"] == "header-customer-name"
    assert report["nc_customer_name"] == ""
    assert "客户名称未确认" in report["reason"]


def test_run_one_row_continues_when_header_account_readback_is_empty(monkeypatch):
    account_readback_timeouts = []

    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    class LocalFakeVerifier:
        def __init__(self, config, located, flow_started_at=None, **_kwargs):
            self.config = config
            self.located = located
            self.flow_started_at = flow_started_at

        def start(self):
            pass

        def submit_snapshot(self, *args, **kwargs):
            return "snapshot-0"

        def submit_row_count(self, *args, **kwargs):
            return "rows-0"

        def wait(self, task_ids=(), *args, **kwargs):
            return pipeline_wait_ok_with_cny_snapshot(task_ids)

        def close(self, timeout=1.0):
            pass

    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    patch_all(monkeypatch, "JABOperator", LocalFakeJAB)
    patch_all(monkeypatch, "fill_header",
        lambda _jab, _business, **_kwargs: [
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "accepted_text": "ACME LTD",
            }
        ],
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    patch_all(monkeypatch, "read_body_table",
        lambda _jab, step: {"ok": True, "step": step, "rows": []},
    )
    patch_all(monkeypatch, "write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    patch_all(monkeypatch, "delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    patch_all(monkeypatch, "wait_header_account_description",
        lambda _jab, timeout=5.0, **_kwargs: (
            account_readback_timeouts.append(timeout)
            or {"accepted": False, "description": "", "text": ""}
        ),
    )
    patch_all(monkeypatch, "verify_and_repair_header_targets",
        lambda *_args, **_kwargs: {"ok": True, "reads": [], "missing": []},
    )
    patch_all(monkeypatch, "DetailPipelineVerifier", LocalFakeVerifier
    )
    patch_all(monkeypatch, "recover_cancelable_modal_now",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("no-save 正常路径不应主动检查弹窗")
        ),
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True
    assert "header_account_readback_warning" in report
    assert account_readback_timeouts == [0.0]


def test_pause_after_customer_diagnoses_cleared_header_and_stops(monkeypatch):
    class LocalFakeInfo:
        name = ""
        description = ""

    class LocalFakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

        def get_context_info(self, vm_id, context):
            return LocalFakeInfo()

        def get_text_context_value(self, vm_id, context):
            return ""

        def release_contexts(self, vm_id, contexts):
            return None

    patch_all(monkeypatch, "open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    patch_all(monkeypatch, "JABOperator", LocalFakeJAB)
    monkeypatch.setattr("builtins.input", lambda _prompt: "")
    patch_all(monkeypatch, "find_receipt_header_field_by_dynamic_path",
        lambda _jab, label, dynamic_index, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [object()],
            "path": f"path-{label}",
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
        },
    )
    patch_all(monkeypatch, "locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("表头诊断失败后不应继续定位明细表")
        ),
    )

    def fake_fill_header(_jab, _business, after_field=None, **_kwargs):
        assert after_field is not None
        steps = [
            {
                "ok": True,
                "label": "财务组织",
                "value": "A001",
                "dynamic_index": 5,
                "dynamic_prefix": "0.0.1.0.0.0.0.5",
                "path": "path-finance",
            },
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "dynamic_index": 5,
                "dynamic_prefix": "0.0.1.0.0.0.0.5",
                "path": "path-customer",
            },
        ]
        after_field("财务组织", "A001", steps[0])
        callback = after_field("客户", "YW03574", steps[1])
        steps[1]["after_field_callback"] = callback
        if callback and not callback.get("ok", True):
            steps.append(
                {
                    "step": "blocked",
                    "ok": False,
                    "label": "客户",
                    "reason": callback["reason"],
                }
            )
        return steps

    patch_all(monkeypatch, "fill_header", fake_fill_header)

    report = run_one_row(
        {},
        plan_row(10),
        save_enabled=False,
        pause_after_header_field="客户",
        diagnose_header_after_pause=True,
    )

    assert report["ok"] is False
    assert report["failed_step"] == "header-fill"
    diagnostics = report["header_pause_diagnostics"][0]
    assert diagnostics["ok"] is False
    assert [item["label"] for item in diagnostics["header_readback"]] == [
        "财务组织",
        "客户",
    ]
    assert all(item["present"] is False for item in diagnostics["header_readback"])
