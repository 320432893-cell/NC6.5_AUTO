# 生命周期：长期
# 覆盖场景：收款单保存后批次查询的纯匹配口径
# 依赖环境：pytest；不依赖 NC/JAB/Excel
# 运行方式：python -m pytest tests/test_receipt_post_save_query.py

from datetime import date
from decimal import Decimal

from core.receipt_models import ReceiptNCIndexedRow, ReceiptPlanRow
from tools.receipt_post_save_query import (
    BatchQueryTarget,
    group_targets_by_org,
    match_targets,
    name_similarity,
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


def test_name_similarity_requires_normalized_high_score():
    assert (
        name_similarity("上海移为通信技术股份有限公司", "上海移为通信技术股份有限公司")
        == 100
    )
    assert name_similarity("上海移为通信技术股份有限公司", "完全不同客户") < 90


def test_match_targets_uses_nc_customer_name_and_exact_amount():
    target = BatchQueryTarget(
        row=plan_row(811),
        row_report={"nc_customer_name": "上海移为通信技术股份有限公司"},
    )

    result = match_targets(
        [target],
        [
            nc_row("SK-OTHER", "完全不同客户"),
            nc_row("SK-OK", "上海移为通信技术股份有限公司"),
        ],
    )

    assert result["matched"] == {811: "SK-OK"}
    assert result["issues"] == {}


def test_match_targets_reports_amount_hit_but_name_mismatch():
    target = BatchQueryTarget(
        row=plan_row(839),
        row_report={"nc_customer_name": "上海移为通信技术股份有限公司"},
    )

    result = match_targets([target], [nc_row("SK-NO", "完全不同客户")])

    assert result["matched"] == {}
    assert "名称相似度" in result["issues"][839]


def test_group_targets_by_org_groups_and_sorts_by_date_then_row():
    targets = [
        BatchQueryTarget(row=plan_row(3, "A001", date(2026, 6, 2)), row_report={}),
        BatchQueryTarget(row=plan_row(2, "A006", date(2026, 6, 1)), row_report={}),
        BatchQueryTarget(row=plan_row(1, "A001", date(2026, 6, 1)), row_report={}),
    ]

    grouped = group_targets_by_org(targets)

    assert list(grouped) == ["A001", "A006"]
    assert [target.row.row for target in grouped["A001"]] == [1, 3]
