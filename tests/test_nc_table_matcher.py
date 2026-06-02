from decimal import Decimal

from core.models import ExcelVoucherItem, PendingMatch
from core.nc_table_matcher import NCTableMatcher


class FakeJAB:
    def normalize_amount(self, value):
        return Decimal(str(value)).quantize(Decimal("0.01"))


class FakeProcessor:
    max_batch_size = 2
    generated_date_col = 18
    generated_date_value = "2026-05-31"
    jab = FakeJAB()


def test_build_increasing_batches_splits_on_descending_row_and_max_size():
    matcher = NCTableMatcher(FakeProcessor())
    item = ExcelVoucherItem(
        row=1,
        raw_key="",
        raw_amount="",
        raw_partner="",
        amount=Decimal("1.00"),
        partner="A",
        voucher="",
        source="split_ab",
        parse_error="",
    )
    matches: list[PendingMatch] = [
        PendingMatch(item=item, nc_row=1, row_data={}),
        PendingMatch(item=item, nc_row=2, row_data={}),
        PendingMatch(item=item, nc_row=5, row_data={}),
        PendingMatch(item=item, nc_row=4, row_data={}),
    ]

    batches = matcher.build_increasing_batches(matches)

    assert batches == [
        [
            PendingMatch(item=item, nc_row=1, row_data={}),
            PendingMatch(item=item, nc_row=2, row_data={}),
        ],
        [PendingMatch(item=item, nc_row=5, row_data={})],
        [PendingMatch(item=item, nc_row=4, row_data={})],
    ]


def test_filter_generated_date_rows():
    matcher = NCTableMatcher(FakeProcessor())
    rows = [
        {"row_index": 1, "extra_text": {18: "2026-05-30"}},
        {"row_index": 2, "extra_text": {18: "2026-05-31"}},
    ]

    assert matcher.filter_generated_date_rows(rows) == [
        {"row_index": 2, "extra_text": {18: "2026-05-31"}}
    ]


def test_as_decimal_delegates_to_jab_normalization():
    matcher = NCTableMatcher(FakeProcessor())

    assert matcher._as_decimal("1.2") == Decimal("1.20")
