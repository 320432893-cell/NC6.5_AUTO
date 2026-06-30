from decimal import Decimal
from typing import Optional

import pytest

from core.errors import ExcelPreflightError, TableMatchError
from core.models import ExcelVoucherItem, MatchIssue, PendingMatch
from core.nc_pending_workflow import NCPendingWorkflow


class FakeSpan:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False


class FakePerf:
    path = None

    def span(self, *args, **kwargs):
        return FakeSpan()

    def event(self, *args, **kwargs):
        pass


class FakeRunState:
    def __init__(self):
        self.events = []
        self.stages = []
        self.counts = {}

    def set_stage(self, stage, **fields):
        self.stages.append((stage, fields))

    def update_counts(self, **counts):
        self.counts.update(counts)

    def event(self, name, **fields):
        self.events.append((name, fields))


class FakeDataHandler:
    def __init__(self, items):
        self.items = items
        self.split_saved = []
        self.result_updates = []
        self.preflight_error: Optional[ExcelPreflightError] = None

    def load_jab_batch_data(self, **kwargs):
        return self.items

    def preflight_jab_items(self, items, **kwargs):
        if self.preflight_error:
            raise self.preflight_error
        return {"rows": len(items), "errors": [], "warnings": [], "parse_errors": 0}

    def save_jab_split_columns(self, items):
        self.split_saved.append(list(items))

    def save_jab_results(self, updates):
        self.result_updates.append(dict(updates))


class FakeTableMatcher:
    def __init__(self, matches, issues):
        self.matches = matches
        self.issues = issues

    def match_current_table(self, items):
        return self.matches, self.issues


class FakeProcessor:
    def __init__(self, items, matches, issues, duplicate_policy="stop"):
        self.data_handler = FakeDataHandler(items)
        self.table_matcher = FakeTableMatcher(matches, issues)
        self.perf = FakePerf()
        self.run_state = FakeRunState()
        self.duplicate_match_policy = duplicate_policy
        self.page_state_calls = []

    def require_page_state(self, expected, items=None, command=""):
        self.page_state_calls.append((expected, command))
        return None

    def record_event(self, name, **kwargs):
        self.run_state.event(name, **kwargs)


def make_item(row, amount="1.00", partner="深圳公司"):
    return ExcelVoucherItem(
        row=row,
        raw_key="",
        raw_amount=amount,
        raw_partner=partner,
        amount=Decimal(amount),
        partner=partner,
        voucher="",
        source="split_ab",
        parse_error="",
    )


def test_plan_stops_before_page_guard_when_excel_preflight_fails():
    item = make_item(2)
    processor = FakeProcessor([item], matches=[], issues=[])
    processor.data_handler.preflight_error = ExcelPreflightError("has_header=true")
    workflow = NCPendingWorkflow(processor)

    with pytest.raises(ExcelPreflightError, match="has_header=true"):
        workflow.dry_run(start_row=2, end_row=2)

    assert processor.page_state_calls == []


def test_generate_stops_before_page_guard_when_excel_preflight_fails(monkeypatch):
    monkeypatch.setattr("core.nc_pending_workflow.check_abort", lambda: None)
    item = make_item(2)
    processor = FakeProcessor([item], matches=[], issues=[])
    processor.data_handler.preflight_error = ExcelPreflightError("has_header=true")
    workflow = NCPendingWorkflow(processor)

    with pytest.raises(ExcelPreflightError, match="has_header=true"):
        workflow.generate_and_save(start_row=2, end_row=2)

    assert processor.page_state_calls == []
    assert processor.data_handler.split_saved == []


# ---- 直测「重复/未全量 → 停手」规则(规则已迁到纯方法,不再经旧 generate 现场匹配)----
def _pending_wf(duplicate_policy="stop"):
    return NCPendingWorkflow(
        FakeProcessor([], matches=[], issues=[], duplicate_policy=duplicate_policy)
    )


def test_ensure_full_pending_match_raises_on_issue():
    item = make_item(2)
    issue = MatchIssue(item=item, reason="重复2条", rows=[1, 17])
    with pytest.raises(TableMatchError, match="未全量匹配"):
        _pending_wf().ensure_full_pending_match([item], [], [issue])


def test_ensure_full_pending_match_raises_on_missing_row():
    item = make_item(2)
    with pytest.raises(TableMatchError, match="未全量匹配"):
        _pending_wf().ensure_full_pending_match([item], [], [])


def test_ensure_full_pending_match_passes_when_all_matched():
    item = make_item(2)
    match = PendingMatch(item=item, nc_row=7, row_data={})
    _pending_wf().ensure_full_pending_match([item], [match], [])  # 无异常即通过


def test_ensure_full_pending_match_skip_policy_bypasses_stop():
    item = make_item(2)
    issue = MatchIssue(item=item, reason="重复2条", rows=[1, 17])
    _pending_wf("skip").ensure_full_pending_match([item], [], [issue])  # skip 策略不停


def test_ensure_full_voucher_match_stops_when_count_short():
    item = make_item(2)
    pending = [PendingMatch(item=item, nc_row=7, row_data={})]
    with pytest.raises(TableMatchError, match="未全量匹配"):
        _pending_wf().ensure_full_voucher_match(pending, [])
