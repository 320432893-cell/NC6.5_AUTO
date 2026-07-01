from decimal import Decimal

import openpyxl
import pytest

from core.data_handler import DataHandler
from core.errors import ExcelLockedError, ExcelPreflightError
from core.models import ExcelVoucherItem


@pytest.fixture
def handler():
    return DataHandler({"excel_path": "unused.xlsx", "jab_batch": {}})


def test_parse_jab_concat_key(handler):
    amount, partner = handler.parse_jab_concat_key(
        "141,688.50 深圳市科贸电子科技有限公司"
    )

    assert amount == Decimal("141688.50")
    assert partner == "深圳市科贸电子科技有限公司"


def test_parse_jab_concat_key_requires_amount_prefix(handler):
    with pytest.raises(ValueError, match="需要以金额开头"):
        handler.parse_jab_concat_key("深圳市科贸电子科技有限公司141688.50")


def test_parse_split_row(handler):
    amount, partner, source = handler.parse_jab_row(
        raw_key="not a concat key",
        raw_amount="1,000",
        raw_partner=" 深圳 市 ",
    )

    assert amount == Decimal("1000.00")
    assert partner == "深圳市"
    assert source == "split_ab"


def test_select_jab_concat_candidate(handler):
    selected = handler.select_jab_concat_candidate("100A / 200B", 2)

    assert selected == "200B"


def test_looks_like_voucher(handler):
    assert handler._looks_like_voucher(123)
    assert handler._looks_like_voucher(123.0)
    assert handler._looks_like_voucher("00123")
    assert not handler._looks_like_voucher("已生成待回填")
    assert not handler._looks_like_voucher(0)


def test_backfill_load_skips_only_numeric_voucher_status(tmp_path):
    excel = tmp_path / "voucher.xlsx"
    write_workbook(
        excel,
        [
            [Decimal("100.00"), "客户A", "已生成未取到凭证号"],
            [Decimal("200.00"), "客户B", "回填未找到"],
            [Decimal("300.00"), "客户C", ""],
            [Decimal("400.00"), "客户D", "00123"],
        ],
    )
    local_handler = DataHandler(
        {
            "excel_path": str(excel),
            "sheet_my": "Sheet1",
            "has_header": False,
            "jab_batch": {},
        }
    )

    items = local_handler.load_jab_batch_data(
        skip_filled=True,
        skip_any_status=False,
    )

    assert [item.row for item in items] == [1, 2, 3]
    assert [item.partner for item in items] == ["客户A", "客户B", "客户C"]


def test_save_workbook_wraps_permission_error(handler):
    class LockedWorkbook:
        closed = False

        def save(self, path):
            raise PermissionError(path)

        def close(self):
            self.closed = True

    wb = LockedWorkbook()

    with pytest.raises(ExcelLockedError, match="Excel 文件无法写入"):
        handler._save_workbook(wb, "写入测试")

    assert wb.closed


def write_workbook(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


def test_preflight_rejects_header_true_when_first_row_looks_like_data(tmp_path):
    excel = tmp_path / "voucher.xlsx"
    write_workbook(
        excel,
        [
            [Decimal("100.00"), "客户A", ""],
            [Decimal("200.00"), "客户B", ""],
        ],
    )
    local_handler = DataHandler(
        {
            "excel_path": str(excel),
            "sheet_my": "Sheet1",
            "has_header": True,
            "jab_batch": {},
        }
    )

    items = local_handler.load_jab_batch_data(skip_any_status=True)

    with pytest.raises(ExcelPreflightError, match="has_header=true"):
        local_handler.preflight_jab_items(
            items,
            start_row=2,
            end_row=4,
            skip_any_status=True,
            context="plan",
        )


def test_preflight_accepts_header_false_for_first_row_data(tmp_path):
    excel = tmp_path / "voucher.xlsx"
    write_workbook(
        excel,
        [
            [Decimal("100.00"), "客户A", ""],
            [Decimal("200.00"), "客户B", ""],
            [Decimal("300.00"), "客户C", ""],
        ],
    )
    local_handler = DataHandler(
        {
            "excel_path": str(excel),
            "sheet_my": "Sheet1",
            "has_header": False,
            "jab_batch": {},
        }
    )

    items = local_handler.load_jab_batch_data(skip_any_status=True)
    report = local_handler.preflight_jab_items(
        items,
        start_row=1,
        end_row=3,
        skip_any_status=True,
        context="plan",
    )

    assert report["rows"] == 3
    assert report["errors"] == []


def test_preflight_rejects_parse_errors_before_jab(handler, monkeypatch):
    monkeypatch.setattr(handler, "detect_header_mismatch", lambda: "")
    items = [
        ExcelVoucherItem(
            row=2,
            raw_key="bad",
            raw_amount="bad",
            raw_partner="",
            amount=None,
            partner="",
            voucher="",
            source="",
            parse_error="A/B拆分列不完整",
        )
    ]

    with pytest.raises(ExcelPreflightError, match="存在格式错误"):
        handler.preflight_jab_items(items, start_row=2, end_row=2)
