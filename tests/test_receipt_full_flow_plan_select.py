# 生命周期：持久维护
# 覆盖的业务场景：收款单完整流程的计划行选择/校验、业务值映射与保存安全确认、CLI 入口
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB（使用 Fake 替身与 monkeypatch）
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_plan_select.py


from tests._receipt_full_flow_helpers import (
    Args,
    Decimal,
    ReceiptPlanIssue,
    build_console_report_lines,
    business_from_plan_row,
    confirm_save,
    parse_args,
    plan_row,
    post_query_failure_reasons,
    pytest,
    select_plan_rows,
    target_to_match_row,
)


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
