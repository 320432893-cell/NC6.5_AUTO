# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程入口的计划行选择、业务值映射和保存安全确认
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_entry.py

from datetime import date
from decimal import Decimal

import pytest

from core.receipt_models import ReceiptPlanIssue, ReceiptPlanRow
from tools.receipt_full_flow_entry import (
    build_console_report_lines,
    business_from_plan_row,
    confirm_save,
    extract_entry_anchor_path,
    extract_entry_dynamic_index,
    extract_entry_scope_hwnd,
    extract_header_accepted_text,
    open_self_made_entry,
    parse_args,
    post_query_failure_reasons,
    read_customer_name_after_header,
    run_one_row,
    save_receipt_by_ctrl_s,
    select_plan_rows,
    wait_receipt_header_anchor_in_current_canvas,
)
from tools.receipt_post_save_query import target_to_match_row


def plan_row(row, fee=Decimal("0.00")):
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


class FakeInfo:
    role = "text"
    role_en_US = "text"
    states = "enabled,visible,showing,editable"
    states_en_US = "enabled,visible,showing,editable"

    def __init__(self, name="", description=""):
        self.name = name
        self.description = description


class Args:
    excel_row: int | None = None
    excel_rows: str | None = None
    limit: int = 1


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


def test_read_customer_name_after_header_uses_customer_description(monkeypatch):
    class FakeJAB:
        def get_context_info(self, _vm_id, _context):
            return FakeInfo(description="INDUSTRIAS METALURGICAS PESCARMONA")

        def get_text_context_value(self, _vm_id, _context):
            return ""

        def release_contexts(self, _vm_id, _contexts):
            pass

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.find_receipt_header_field_by_dynamic_path",
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
                "dynamic_index": 4,
                "path": "customer.path",
            }
        ],
        4,
        197550,
    )

    assert result["ok"] is True
    assert result["value"] == "INDUSTRIAS METALURGICAS PESCARMONA"
    assert result["source"] == "path-readback"


def test_select_plan_rows_skips_issue_rows_and_defaults_limit_one():
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

    selected = select_plan_rows(rows, issues, Args())

    assert [row.row for row in selected] == [3]


def test_select_plan_rows_can_target_specific_excel_row():
    args = Args()
    args.excel_row = 4
    args.limit = 10

    selected = select_plan_rows([plan_row(3), plan_row(4)], [], args)

    assert [row.row for row in selected] == [4]


def test_select_plan_rows_can_target_multiple_excel_rows_in_order():
    args = Args()
    args.excel_rows = "4,2,4,3"
    args.limit = 10

    selected = select_plan_rows([plan_row(2), plan_row(3), plan_row(4)], [], args)

    assert [row.row for row in selected] == [4, 2, 3]


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


def test_main_external_stop_finishes_aborted_with_exit_code_3(monkeypatch, tmp_path):
    # 外部停止经 JAB 原语 check_abort 抛 SystemExit；按 ENGINE_CONTRACT 应收尾为
    # aborted、退出码 3，而不是让 run_state 卡在 running 或被当成崩溃。
    import json
    import types

    from core.paths import logs_dir
    import tools.receipt_full_flow_entry as entry

    monkeypatch.setenv("NC_RUNTIME_DIR", str(tmp_path))

    class FakeWorkbook:
        def __init__(self, config, excel_path=None):
            pass

        def build_local_plan(self, write_sheet=False):
            return [], [], {}

    def _raise_stop(*args, **kwargs):
        raise SystemExit("外部停止")

    monkeypatch.setattr(entry, "load_config", lambda path: {"_config_path": path})
    monkeypatch.setattr(entry, "ReceiptEntryWorkbook", FakeWorkbook)
    monkeypatch.setattr(
        entry,
        "select_plan_rows",
        lambda plan_rows, issues, args: [types.SimpleNamespace(row=2)],
    )
    monkeypatch.setattr(entry, "run_one_row", _raise_stop)
    monkeypatch.setattr(entry, "write_last_report", lambda report: None)
    monkeypatch.setattr(entry, "print_report", lambda report, args: None)

    exit_code = entry.main([])

    assert exit_code == 3
    state = json.loads((logs_dir() / "run_state.json").read_text(encoding="utf-8"))
    assert state["status"] == "aborted"


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


def test_header_anchor_wait_polls_current_canvas_every_point_two(monkeypatch):
    calls = {"anchor": [], "sleep": []}

    def fake_anchor(_jab, hwnd, timeout=0.05):
        calls["anchor"].append((hwnd, timeout))
        if len(calls["anchor"]) == 1:
            return {"ok": False, "reason": "not ready"}
        return {
            "ok": True,
            "scope_hwnd": hwnd,
            "dynamic_index": 5,
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
        }

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.resolve_receipt_header_anchor_in_canvas",
        fake_anchor,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.time.sleep",
        lambda seconds: calls["sleep"].append(seconds),
    )

    result = wait_receipt_header_anchor_in_current_canvas(
        object(),
        919586,
        timeout=1.2,
        interval=0.2,
    )

    assert result["ok"] is True
    assert result["dynamic_index"] == 5
    assert len(calls["anchor"]) == 2
    assert calls["anchor"][0][0] == 919586
    assert calls["sleep"] == [0.2]
    assert result["poll_interval"] == 0.2


def test_open_self_made_entry_always_runs_new_probe(monkeypatch):
    calls = []

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.run_receipt_new_probe",
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

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.run_receipt_new_probe",
        lambda: (_ for _ in ()).throw(AssertionError("不应起子进程开单")),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.run_receipt_new_probe_with_jab",
        lambda actual_jab: (
            calls.append(actual_jab) or {"ok": True, "mode": "in-process"}
        ),
    )

    assert open_self_made_entry({"receipt_entry": {}}, jab) == {
        "ok": True,
        "mode": "in-process",
    }
    assert calls == [jab]


def test_save_receipt_uses_keyboard_hotkey_not_jab_button_or_sendinput(monkeypatch):
    calls = {"hotkey": [], "states": 0}

    class FakeJAB:
        def click_save(self, timeout=None):
            raise AssertionError("收款单保存不能调用凭证/制单保存按钮查找")

        def wait_save_success(self, timeout=None):
            raise AssertionError("收款单保存不等保存成功提示作为触发闭包")

        def press_hotkey(self, *keys, wait=None):
            calls["hotkey"].append((keys, wait))

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.probe_receipt_entry_page",
        lambda _jab: {
            "ok": True,
            "scope": {"scope_hwnd": 12345},
            "method": "header-scope",
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda window: {"ok": True, "target_window": window, "foreground": {}},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 12345}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_receipt_parent_new_ready",
        lambda _windows: {
            "ok": True,
            "usable_new_button_count": 1,
            "usable_new_buttons": [{"control": {"description": "新增(Ctrl+N)"}}],
        },
    )

    def fake_detect(_windows):
        calls["states"] += 1
        return {"ok": False, "reason": "已回到新增态"}

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_self_made_entry_state",
        fake_detect,
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.5)

    assert result["ok"] is True
    assert result["triggered"] is True
    assert calls["hotkey"] == [(("ctrl", "s"), 0)]
    assert result["hotkey"]["mode"] == "jab.press_hotkey"
    assert result["oracle"]["name"] == "receipt_parent_new_ready_after_save"
    assert result["oracle"]["parent_new_state"]["ok"] is True
    assert calls["states"] == 1


def test_save_receipt_stops_before_oracle_when_foreground_guard_fails(monkeypatch):
    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            raise AssertionError("前台保护失败时不应触发 Ctrl+S")

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.probe_receipt_entry_page",
        lambda _jab: {"ok": True, "scope": {"scope_hwnd": 12345}},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda _window: {"ok": False, "reason": "当前前台窗口不是目标 NC 窗口"},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: (_ for _ in ()).throw(
            AssertionError("Ctrl+S 未发出时不应继续等保存 oracle")
        ),
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.5)

    assert result["ok"] is False
    assert result["triggered"] is False
    assert "当前前台窗口不是目标 NC 窗口" in result["reason"]


def test_save_receipt_uses_entry_state_hwnd_without_header_scope_probe(monkeypatch):
    calls = {"guard": [], "hotkey": []}

    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            calls["hotkey"].append((keys, wait))

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 24680, "controls": []}],
    )

    state_calls = {"count": 0}

    def fake_detect(_windows):
        state_calls["count"] += 1
        if state_calls["count"] == 1:
            return {
                "ok": True,
                "hits": [{"window": {"hwnd": 24680, "class_name": "SunAwtCanvas"}}],
            }
        return {"ok": False, "reason": "已回到新增态"}

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_self_made_entry_state",
        fake_detect,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda window: calls["guard"].append(window) or {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_receipt_parent_new_ready",
        lambda _windows: {"ok": True, "usable_new_button_count": 1},
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.5)

    assert result["ok"] is True
    assert calls["guard"] == [{"hwnd": 24680}]
    assert calls["hotkey"] == [(("ctrl", "s"), 0)]


def test_save_receipt_does_not_treat_missing_entry_buttons_as_success_without_new_button(
    monkeypatch,
):
    class FakeJAB:
        def press_hotkey(self, *keys, wait=None):
            return None

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.probe_receipt_entry_page",
        lambda _jab: {"ok": True, "scope": {"scope_hwnd": 12345}},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.foreground_matches_window",
        lambda _window: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.collect_receipt_new_windows",
        lambda _jab: [{"hwnd": 12345}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_self_made_entry_state",
        lambda _windows: {"ok": False, "reason": "三按钮读不到"},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.detect_receipt_parent_new_ready",
        lambda _windows: {"ok": False, "usable_new_button_count": 0},
    )

    result = save_receipt_by_ctrl_s(FakeJAB(), timeout=0.01)

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
            return {"ok": True, "submitted": list(task_ids), "done": len(task_ids)}

        def close(self, timeout=1.0):
            calls["closed"] = timeout

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
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
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
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
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
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

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        fail_sync_read_before,
    )

    def fake_write_detail(_jab, business, _located, after_field=None, **_kwargs):
        field = {"col": 4, "name": "收款银行账户", "value_key": "bank_account"}
        step = {"ok": True, "input_ok": True, "name": field["name"]}
        assert after_field is not None
        step["async_verify_task"] = after_field(0, field, business, step)
        return [step]

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        fake_write_detail,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda _jab, _timeout=0.0, scope=None: (
            calls["account_scope"].append(scope) or {"accepted": True}
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", FakeVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
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
    assert calls["wait"] == [(["field-0", "rows-0"], 2.0)]
    assert calls["fill_header_kwargs"][0]["scope_hwnd"] == 2002
    assert calls["fill_header_kwargs"][0]["dynamic_index"] == 5
    assert (
        calls["fill_header_kwargs"][0]["anchor_path"]
        == "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.1.1.1.0"
    )
    assert calls["body_locate_kwargs"][0]["scope_hwnd"] == 2002
    assert (
        calls["body_locate_kwargs"][0]["cached"]["best"]["path"]
        == "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.0.1.0.2.1.0.0.0.0.0"
    )
    assert calls["body_locate_kwargs"][0]["cached"]["best"]["window"] == {
        "hwnd": 2002,
        "class_name": "SunAwtCanvas",
    }
    assert calls["account_scope"][0]["scope_hwnd"] == 2002
    assert calls["account_scope"][0]["dynamic_index"] == 5
    assert report["before_table"]["skipped"] is True
    assert report["after_table"]["skipped"] is True
    assert calls["closed"] == 0.2


def test_run_one_row_retries_current_canvas_header_anchor(monkeypatch):
    calls = {
        "anchor_retry": [],
        "fill_header_kwargs": [],
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

        def wait(self, *_args, **_kwargs):
            return {"ok": True}

        def close(self, timeout=1.0):
            pass

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
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
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_receipt_header_anchor_in_current_canvas",
        lambda _jab, hwnd, timeout=1.2, interval=0.2: (
            calls["anchor_retry"].append((hwnd, timeout, interval))
            or {
                "ok": True,
                "scope_hwnd": hwnd,
                "dynamic_index": 5,
                "dynamic_prefix": "0.0.1.0.0.0.0.5",
            }
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business, **kwargs: (
            calls["fill_header_kwargs"].append(kwargs)
            or [{"ok": True, "label": "客户", "accepted_text": "ACME LTD"}]
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
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
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", FakeVerifier
    )

    report = run_one_row({}, plan_row(10), save_enabled=False)

    assert report["ok"] is True, report
    assert calls["anchor_retry"] == [(919586, 1.2, 0.2)]
    assert report["entry_dynamic_index"] == 5
    assert report["entry_header_anchor_retry"]["ok"] is True
    assert calls["fill_header_kwargs"][0]["scope_hwnd"] == 919586
    assert calls["fill_header_kwargs"][0]["dynamic_index"] == 5
    assert calls["body_locate_kwargs"][0]["scope_hwnd"] == 919586
    assert (
        calls["body_locate_kwargs"][0]["cached"]["best"]["path"]
        == "0.0.1.0.0.0.0.5.0.0.0.1.1.0.0.0.0.1.0.2.1.0.0.0.0.0"
    )


def test_run_one_row_stops_when_current_canvas_header_anchor_missing(monkeypatch):
    class FakeJAB:
        def __init__(self, config):
            self.config = config

        def ensure_started(self):
            return True

        def close(self):
            return None

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
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
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_receipt_header_anchor_in_current_canvas",
        lambda *_args, **_kwargs: {
            "ok": False,
            "reason": "当前 canvas 未找到财务组织(O) 锚点",
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("锚点失败后不应进入表头写入")
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
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
            return {
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
            }

        def snapshot(self):
            return {"ok": len(calls["wait"]) >= 2}

        def close(self, timeout=1.0):
            calls["closed"] = timeout

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
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
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", FakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda *_args, **_kwargs: [
            {"ok": True, "label": "客户", "accepted_text": "ACME LTD"}
        ],
    )

    def fake_locate(*_args, **_kwargs):
        calls["locate"] += 1
        return located

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        fake_locate,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
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
        row_count,
        field,
        next_col,
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
                "row_count": row_count,
                "field": field["name"],
                "next_col": next_col,
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

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        fake_write_detail,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_field_once",
        fake_write_field_once,
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", RepairingVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
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
            "row_count": 1,
            "field": "收款银行账户",
            "next_col": 4,
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

        def wait(self, *args, **kwargs):
            return {"ok": True}

        def close(self, timeout=1.0):
            pass

    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda *_args, **_kwargs: [
            {"ok": True, "label": "客户", "accepted_text": "ACME LTD"}
        ],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", LocalFakeVerifier
    )

    def fake_save(*_args, **_kwargs):
        calls["save"] += 1
        if calls["save"] == 1:
            return {"ok": False, "reason": "前台窗口不是目标 NC 窗口"}
        return {"ok": True, "triggered": True}

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.save_receipt_by_ctrl_s", fake_save
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
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

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business, **_kwargs: [
            {"ok": True, "label": "客户", "value": "YW03574"}
        ],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        lambda _jab, step: {"ok": True, "step": step, "rows": []},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda *_args, **_kwargs: {"accepted": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", LocalFakeVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
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

        def wait(self, *args, **kwargs):
            return {"ok": True}

        def close(self, timeout=1.0):
            pass

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.fill_header",
        lambda _jab, _business, **_kwargs: [
            {
                "ok": True,
                "label": "客户",
                "value": "YW03574",
                "accepted_text": "ACME LTD",
            }
        ],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
        lambda _jab, max_rows=5, **_kwargs: {
            "best": {"path": "0.1", "row_count": 1, "col_count": 25, "window": {}},
            "candidates": [],
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.read_body_table",
        lambda _jab, step: {"ok": True, "step": step, "rows": []},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.write_detail_line_by_screen",
        lambda *_args, **_kwargs: [{"ok": True}],
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True},
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.wait_header_account_description",
        lambda _jab, timeout=5.0, **_kwargs: (
            account_readback_timeouts.append(timeout)
            or {"accepted": False, "description": "", "text": ""}
        ),
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.DetailPipelineVerifier", LocalFakeVerifier
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.recover_cancelable_modal_now",
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

    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.open_self_made_entry",
        lambda _config, _jab=None: open_report_with_header_anchor(),
    )
    monkeypatch.setattr("tools.receipt_full_flow_entry.JABOperator", LocalFakeJAB)
    monkeypatch.setattr("builtins.input", lambda _prompt: "")
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.find_receipt_header_field_by_dynamic_path",
        lambda _jab, label, dynamic_index, **_kwargs: {
            "ok": True,
            "context": object(),
            "vm_id": 1,
            "owned_contexts": [object()],
            "path": f"path-{label}",
            "dynamic_prefix": "0.0.1.0.0.0.0.5",
        },
    )
    monkeypatch.setattr(
        "tools.receipt_full_flow_entry.locate_receipt_body_table_cached",
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

    monkeypatch.setattr("tools.receipt_full_flow_entry.fill_header", fake_fill_header)

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
