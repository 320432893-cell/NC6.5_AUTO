# 生命周期：长期
# 覆盖场景：收款单保存后批次查询的纯匹配口径
# 依赖环境：pytest；不依赖 NC/JAB/Excel
# 运行方式：python -m pytest tests/test_receipt_post_save_query.py

from datetime import date
from decimal import Decimal

from core.receipt_models import ReceiptMatchIssue, ReceiptNCIndexedRow, ReceiptPlanRow
from core.receipt_post_save_query import (
    BatchQueryTarget,
    document_no_sort_number,
    group_targets_by_org,
    match_snapshot_to_result,
    run_post_save_batch_query,
)


def plan_row(row, org_code="A001", receipt_date=date(2026, 6, 15)):
    return ReceiptPlanRow(
        row=row,
        receipt_date=receipt_date,
        payer_name=f"PAYER-{row}",
        raw_amount=Decimal("100.00"),
        bank="PayPal",
        currency="USD",
        customer_code="YW001",
        fee=Decimal("0.00"),
        organization_code=org_code,
        organization_name=f"主体-{org_code}",
        organization_short_name=org_code,
        account_id="paypal",
        account_label="PayPal",
        account_no="paypal-account",
        header_currency_code="USD",
        duplicate_key=(str(row),),
    )


def nc_row(document_no, name, amount="100.00", receipt_date=date(2026, 6, 15)):
    return ReceiptNCIndexedRow(
        row_index=1,
        table_index=0,
        document_no=document_no,
        document_date=receipt_date,
        original_amount=Decimal(amount),
        name=name,
    )


def test_match_snapshot_result_uses_incremental_match_success():
    target = BatchQueryTarget(
        row=plan_row(811),
        row_report={"nc_customer_name": "上海移为通信技术股份有限公司"},
    )
    matched_row = nc_row("SK-OK", "上海移为通信技术股份有限公司")

    result = match_snapshot_to_result(
        [target],
        {
            "matched": {811: matched_row},
            "match_issues": [],
        },
    )

    assert result == {
        "matched": {811: "SK-OK"},
        "issues": {},
    }


def test_match_snapshot_result_is_the_single_post_query_source():
    target = BatchQueryTarget(
        row=plan_row(828),
        row_report={"nc_customer_name": "CalAmp Wireless Networks Corporation"},
    )
    result = match_snapshot_to_result(
        [target],
        {
            "matched": {},
            "match_issues": [
                ReceiptMatchIssue(
                    excel_row=828,
                    reason="名称匹配但金额不匹配",
                    nc_rows=[0, 2],
                )
            ],
        },
    )

    assert result == {
        "matched": {},
        "issues": {828: "名称匹配但金额不匹配"},
    }


def test_match_snapshot_resolves_duplicate_by_largest_document_no():
    target = BatchQueryTarget(
        row=plan_row(1956),
        row_report={"nc_customer_name": "Leader Products Co Pty Ltd"},
    )

    result = match_snapshot_to_result(
        [target],
        {
            "matched": {},
            "nc_rows": [
                nc_row("D22026062200027101", "Leader Products Co Pty Ltd"),
                nc_row("D22026062200027134", "Leader Products Co Pty Ltd"),
            ],
            "match_issues": [
                ReceiptMatchIssue(
                    excel_row=1956,
                    reason="重复2条：名称和金额相同，需人工确认",
                    nc_rows=[0, 1],
                )
            ],
        },
    )

    assert result == {
        "matched": {1956: "D22026062200027134"},
        "issues": {},
    }


def test_match_snapshot_keeps_duplicate_issue_when_document_no_not_sortable():
    target = BatchQueryTarget(
        row=plan_row(1956),
        row_report={"nc_customer_name": "Leader Products Co Pty Ltd"},
    )

    result = match_snapshot_to_result(
        [target],
        {
            "matched": {},
            "nc_rows": [
                nc_row("TEMP-A", "Leader Products Co Pty Ltd"),
                nc_row("TEMP-B", "Leader Products Co Pty Ltd"),
            ],
            "match_issues": [
                ReceiptMatchIssue(
                    excel_row=1956,
                    reason="重复2条：名称和金额相同，需人工确认",
                    nc_rows=[0, 1],
                )
            ],
        },
    )

    assert result == {
        "matched": {},
        "issues": {1956: "重复2条：名称和金额相同，需人工确认"},
    }


def test_document_no_sort_number_uses_all_digits():
    assert document_no_sort_number("D22026062200027134") == 22026062200027134
    assert document_no_sort_number("") is None


def test_group_targets_by_org_groups_and_sorts_by_date_then_row():
    targets = [
        BatchQueryTarget(row=plan_row(3, "A001", date(2026, 6, 2)), row_report={}),
        BatchQueryTarget(row=plan_row(2, "A006", date(2026, 6, 1)), row_report={}),
        BatchQueryTarget(row=plan_row(1, "A001", date(2026, 6, 1)), row_report={}),
    ]

    grouped = group_targets_by_org(targets)

    assert list(grouped) == ["A001", "A006"]
    assert [target.row.row for target in grouped["A001"]] == [1, 3]


def test_run_post_save_batch_query_reuses_external_jab_without_closing(monkeypatch):
    calls = {"closed": 0, "ensured": 0}

    class FakeJAB:
        def ensure_started(self):
            calls["ensured"] += 1

        def close(self):
            calls["closed"] += 1

    monkeypatch.setattr(
        "core.receipt_post_save_query.query_one_org",
        lambda _jab, _config, _query_cfg, _jab_cfg, _extractor, org_code, targets, _timings: {
            "ok": True,
            "organization_code": org_code,
            "target_rows": [target.row.row for target in targets],
            "match": {"matched": {}, "issues": {}},
        },
    )
    monkeypatch.setattr(
        "core.receipt_post_save_query.apply_group_match_results",
        lambda *_args, **_kwargs: None,
    )
    monkeypatch.setattr(
        "core.receipt_post_save_query.ReceiptNCResultExtractor",
        lambda _config: object(),
    )

    jab = FakeJAB()
    config = {"receipt_entry": {"query": {"jab": {}}}}
    results, report = run_post_save_batch_query(
        config,
        [plan_row(1)],
        [{"excel_row": 1, "ok": True, "nc_customer_name": "PAYER-1"}],
        jab=jab,
    )

    assert report["ok"] is True
    assert [item.plan_row.row for item in results] == [1]
    assert calls["ensured"] == 1
    assert calls["closed"] == 0


def test_post_save_query_preserves_business_skipped_rows(monkeypatch):
    class FakeJAB:
        def ensure_started(self):
            pass

        def close(self):
            pass

    monkeypatch.setattr(
        "core.receipt_post_save_query.query_one_org",
        lambda _jab, _config, _query_cfg, _jab_cfg, _extractor, org_code, targets, _timings: {
            "ok": True,
            "organization_code": org_code,
            "target_rows": [target.row.row for target in targets],
            "match": {"matched": {1: "D001"}, "issues": {}},
        },
    )
    monkeypatch.setattr(
        "core.receipt_post_save_query.apply_group_match_results",
        lambda *_args, **_kwargs: None,
    )
    monkeypatch.setattr(
        "core.receipt_post_save_query.ReceiptNCResultExtractor",
        lambda _config: object(),
    )

    results, report = run_post_save_batch_query(
        {"receipt_entry": {"query": {"jab": {}}}},
        [plan_row(1), plan_row(2)],
        [
            {"excel_row": 1, "ok": True, "nc_customer_name": "PAYER-1"},
            {
                "excel_row": 2,
                "ok": False,
                "business_exception_skipped": True,
                "sheet2_exception_reason": "客户名称不匹配；已取消当前单据并跳过",
                "nc_customer_name": "WRONG NAME",
            },
        ],
        jab=FakeJAB(),
    )

    assert report["ok"] is True
    assert results[0].local_status == "通过"
    assert results[1].local_status == "跳过"
    assert results[1].exception_reason == "客户名称不匹配；已取消当前单据并跳过"
    assert results[1].nc_customer_name == "WRONG NAME"
