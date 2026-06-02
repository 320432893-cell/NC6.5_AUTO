from decimal import Decimal

import pytest

from core.errors import TableMatchError
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

    def load_jab_batch_data(self, **kwargs):
        return self.items

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

    def require_page_state(self, expected, items=None, command=""):
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


def test_generate_stops_before_nc_clicks_on_duplicate_match(monkeypatch):
    monkeypatch.setattr("core.nc_pending_workflow.check_abort", lambda: None)
    duplicate_item = make_item(2)
    unique_item = make_item(3, amount="2.00", partner="上海公司")
    match = PendingMatch(item=unique_item, nc_row=7, row_data={})
    duplicate = MatchIssue(item=duplicate_item, reason="重复2条", rows=[1, 17])
    processor = FakeProcessor(
        [duplicate_item, unique_item],
        matches=[match],
        issues=[duplicate],
    )
    workflow = NCPendingWorkflow(processor)
    processed_matches = []

    def fake_process_full_selection(matches, max_save_batches=None):
        processed_matches.extend(matches)
        return matches, 1

    monkeypatch.setattr(workflow, "process_full_selection", fake_process_full_selection)

    with pytest.raises(TableMatchError, match="匹配不唯一"):
        workflow.generate_and_save()

    assert processed_matches == []
    assert processor.data_handler.result_updates == []
    assert processor.run_state.events[-1] == (
        "duplicate_match_issues",
        {
            "policy": "stop",
            "count": 1,
            "issues": [
                {
                    "excel_row": 2,
                    "amount": "1.00",
                    "partner": "深圳公司",
                    "nc_rows": [1, 17],
                },
            ],
        },
    )


def test_generate_can_skip_duplicate_match_and_process_unique_matches(monkeypatch):
    monkeypatch.setattr("core.nc_pending_workflow.check_abort", lambda: None)
    duplicate_item = make_item(2)
    unique_item = make_item(3, amount="2.00", partner="上海公司")
    match = PendingMatch(item=unique_item, nc_row=7, row_data={})
    duplicate = MatchIssue(item=duplicate_item, reason="重复2条", rows=[1, 17])
    processor = FakeProcessor(
        [duplicate_item, unique_item],
        matches=[match],
        issues=[duplicate],
        duplicate_policy="skip",
    )
    workflow = NCPendingWorkflow(processor)
    processed_matches = []

    def fake_process_full_selection(matches, max_save_batches=None):
        processed_matches.extend(matches)
        return matches, 1

    monkeypatch.setattr(workflow, "process_full_selection", fake_process_full_selection)

    assert workflow.generate_and_save() == 1

    assert processed_matches == [match]
    assert processor.data_handler.result_updates == [
        {2: "重复2条-NC行1,17"},
    ]
    assert processor.run_state.events[-1] == (
        "duplicate_match_issues",
        {
            "policy": "skip",
            "count": 1,
            "issues": [
                {
                    "excel_row": 2,
                    "amount": "1.00",
                    "partner": "深圳公司",
                    "nc_rows": [1, 17],
                },
            ],
        },
    )
