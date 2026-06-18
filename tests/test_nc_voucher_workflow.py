from decimal import Decimal

import pytest

from core.errors import ContractViolation
from core.models import ExcelVoucherItem, VoucherSaveMatch
from core.nc_voucher_workflow import NCVoucherWorkflow


class FakeSpan:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False


class FakePerf:
    def __init__(self):
        self.spans = []

    def span(self, *args, **kwargs):
        self.spans.append((args, kwargs))
        return FakeSpan()


class FakeJAB:
    def __init__(self, *, window_exists=True, table_counts=None):
        self._window_exists = window_exists
        self._table_counts = table_counts or []
        self.read_table_calls = []
        self.closed_windows = []

    def window_exists(self, *args, **kwargs):
        return self._window_exists

    def read_window_table_counts(self, *args, **kwargs):
        return list(self._table_counts)

    def read_window_table_cells(self, *args, **kwargs):
        self.read_table_calls.append((args, kwargs))
        if kwargs.get("max_rows") is not None:
            return [
                {
                    "row_count": kwargs["max_rows"],
                    "col_count": kwargs.get("max_cols", 13),
                    "rows": [],
                }
            ]
        raise AssertionError("保存验证不应读取整张制单表")

    def close_window_by_title(self, *args, **kwargs):
        self.closed_windows.append((args, kwargs))
        self._window_exists = False


class FakeProcessor:
    def __init__(self, jab):
        self.jab = jab
        self.perf = FakePerf()
        self.voucher_record_timeout = 0.1
        self.voucher_window_title = "制单"
        self.voucher_window_class = "SunAwtDialog"
        self.batch_cfg = {
            "state_wait_timeout": 0.1,
            "state_wait_interval": 0.01,
            "voucher_table_read_min_rows": 5,
            "voucher_table_read_row_buffer": 2,
        }


def make_match(row, *, table_index=2, table_rows=5, voucher_row=0):
    item = ExcelVoucherItem(
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
    return VoucherSaveMatch(
        item=item,
        nc_row=row,
        row_data={},
        table_index=table_index,
        table_rows=table_rows,
        voucher_row=voucher_row,
        voucher_cells=["1.00", "深圳公司"],
    )


def test_verify_voucher_batch_removed_passes_on_exact_table_count_decrease():
    jab = FakeJAB(
        table_counts=[
            {"table_index": 2, "row_count": 3, "col_count": 13},
        ]
    )
    workflow = NCVoucherWorkflow(FakeProcessor(jab))
    batch = [
        make_match(10, table_rows=5, voucher_row=0),
        make_match(11, table_rows=5, voucher_row=1),
    ]

    assert workflow.verify_voucher_batch_removed(batch, before_count=5) is True


def test_verify_voucher_batch_removed_treats_empty_table_as_queue_done():
    jab = FakeJAB(
        table_counts=[
            {"table_index": 2, "row_count": 0, "col_count": 13},
        ]
    )
    workflow = NCVoucherWorkflow(FakeProcessor(jab))

    assert (
        workflow.verify_voucher_batch_removed(
            [make_match(10, table_rows=1)], before_count=1
        )
        == "empty_window"
    )


def test_verify_voucher_batch_removed_rejects_closed_window():
    workflow = NCVoucherWorkflow(FakeProcessor(FakeJAB(window_exists=False)))

    with pytest.raises(ContractViolation, match="制单窗口已关闭"):
        workflow.verify_voucher_batch_removed(
            [make_match(10, table_rows=1)], before_count=1
        )


def test_verify_voucher_batch_removed_rejects_over_removed_count():
    jab = FakeJAB(
        table_counts=[
            {"table_index": 2, "row_count": 2, "col_count": 13},
        ]
    )
    workflow = NCVoucherWorkflow(FakeProcessor(jab))
    batch = [
        make_match(10, table_rows=5, voucher_row=0),
        make_match(11, table_rows=5, voucher_row=1),
    ]

    with pytest.raises(ContractViolation, match="减少超过预期"):
        workflow.verify_voucher_batch_removed(batch, before_count=5)


def test_read_voucher_tables_limits_rows_to_pending_count_plus_buffer():
    jab = FakeJAB()
    processor = FakeProcessor(jab)
    workflow = NCVoucherWorkflow(processor)

    workflow.read_voucher_tables(pending_count=3)

    assert jab.read_table_calls[-1][1]["max_rows"] == 7


def test_wait_for_voucher_tables_returns_first_available_snapshot():
    jab = FakeJAB()
    processor = FakeProcessor(jab)
    workflow = NCVoucherWorkflow(processor)

    tables = workflow.wait_for_voucher_tables(pending_count=3, timeout=0)

    assert tables
    assert len(jab.read_table_calls) == 1
    assert jab.read_table_calls[-1][1]["max_rows"] == 7


def test_close_voucher_window_after_save_polls_without_action_wait():
    jab = FakeJAB(window_exists=True)
    workflow = NCVoucherWorkflow(FakeProcessor(jab))

    assert workflow.close_voucher_window_after_save([make_match(10)]) is True

    assert jab.closed_windows
    assert jab.closed_windows[-1][1]["wait"] == 0
