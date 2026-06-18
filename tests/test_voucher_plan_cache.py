from decimal import Decimal

import pytest

from core.errors import TableMatchError
from core.models import ExcelVoucherItem, PendingMatch
from core.voucher_plan_cache import (
    matches_from_plan_cache,
    validate_voucher_plan_cache,
    write_voucher_plan_cache,
)


def make_item(row):
    return ExcelVoucherItem(
        row=row,
        raw_key="",
        raw_amount="1.00",
        raw_partner="深圳公司",
        amount=Decimal("1.00"),
        partner="深圳公司",
        voucher="",
        source="split_ab",
        parse_error="",
    )


def test_voucher_plan_cache_round_trip(monkeypatch, tmp_path):
    monkeypatch.setattr("core.voucher_plan_cache.logs_dir", lambda: tmp_path)
    item = make_item(3)
    config = {
        "excel_path": "x.xlsx",
        "sheet_my": "Sheet1",
        "has_header": False,
    }

    path = write_voucher_plan_cache(
        config=config,
        limit=3,
        start_row=2,
        end_row=None,
        matches=[PendingMatch(item=item, nc_row=8, row_data={})],
    )
    cache = path.read_text(encoding="utf-8")

    assert "x.xlsx" in cache
    loaded = {
        "excel_path": "x.xlsx",
        "sheet": "Sheet1",
        "has_header": False,
        "limit": 3,
        "start_row": 2,
        "end_row": None,
        "rows": [{"excel_row": 3, "nc_row": 8}],
    }
    validate_voucher_plan_cache(
        cache=loaded,
        config=config,
        limit=3,
        start_row=2,
        end_row=None,
        pending=[item],
    )
    matches = matches_from_plan_cache(loaded, [item])
    assert matches[0].nc_row == 8
    assert matches[0].item.row == 3


def test_voucher_plan_cache_rejects_parameter_mismatch():
    item = make_item(3)

    with pytest.raises(TableMatchError, match="参数不一致"):
        validate_voucher_plan_cache(
            cache={
                "excel_path": "old.xlsx",
                "sheet": "Sheet1",
                "has_header": True,
                "limit": None,
                "start_row": None,
                "end_row": None,
                "rows": [{"excel_row": 3, "nc_row": 8}],
            },
            config={"excel_path": "new.xlsx", "sheet_my": "Sheet1", "has_header": True},
            limit=None,
            start_row=None,
            end_row=None,
            pending=[item],
        )
