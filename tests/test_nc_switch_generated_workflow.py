from core.nc_switch_generated_workflow import NCSwitchGeneratedWorkflow


class FakeSpan:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False


class FakePerf:
    def span(self, *args, **kwargs):
        return FakeSpan()


class FakeJAB:
    def __init__(self):
        self.keys = []
        self.activated = []

    def activate_window_by_title(self, title, class_name=None, timeout=5):
        self.activated.append((title, class_name, timeout))

    def maximize_window_by_title(self, title, class_name=None, timeout=5):
        self.activated.append((title, class_name, timeout))
        return True

    def press_key(self, key, wait=0):
        self.keys.append((key, wait))


class FakeProcessor:
    def __init__(self):
        self.jab = FakeJAB()
        self.events = []
        self.perf = FakePerf()
        self.batch_cfg = {"state_wait_timeout": 0.3}

    def record_event(self, name, **kwargs):
        self.events.append((name, kwargs))


def test_open_query_hotkey_retries_until_query_window(monkeypatch):
    processor = FakeProcessor()
    workflow = NCSwitchGeneratedWorkflow(processor)
    current_time = {"value": 0.0}
    seen_timeouts = []

    def fake_monotonic():
        return current_time["value"]

    def fake_find_query_window(open_query, timeout):
        seen_timeouts.append(timeout)
        current_time["value"] += timeout
        if len(processor.jab.keys) >= 2:
            return 12345
        return None

    monkeypatch.setattr(
        "core.nc_switch_generated_workflow.time.monotonic", fake_monotonic
    )
    monkeypatch.setattr(workflow, "find_query_window", fake_find_query_window)

    query_hwnd = workflow.open_query_with_hotkey_until_window(
        {
            "main_title": "Yonyou UClient",
            "main_class": "YonyouUWnd",
            "key": "f3",
            "window_poll_interval": 0.1,
            "open_key_retry_interval": 0.2,
            "key_wait": 0.0,
        },
        timeout=1.0,
    )

    assert query_hwnd == 12345
    assert processor.jab.keys == [("f3", 0.0), ("f3", 0.0)]
    assert seen_timeouts[:3] == [0.1, 0.1, 0.1]
    assert processor.events[-1][0] == "event_query_window_detected"
    assert processor.events[-1][1]["attempts"] == 2


def test_open_query_hotkey_timeout_is_capped_at_two_seconds(monkeypatch):
    processor = FakeProcessor()
    workflow = NCSwitchGeneratedWorkflow(processor)
    current_time = {"value": 0.0}
    elapsed_wait = {"value": 0.0}

    def fake_monotonic():
        return current_time["value"]

    def fake_find_query_window(open_query, timeout):
        elapsed_wait["value"] += timeout
        current_time["value"] += timeout
        return None

    monkeypatch.setattr(
        "core.nc_switch_generated_workflow.time.monotonic", fake_monotonic
    )
    monkeypatch.setattr(workflow, "find_query_window", fake_find_query_window)

    query_hwnd = workflow.open_query_with_hotkey_until_window(
        {
            "main_title": "Yonyou UClient",
            "main_class": "YonyouUWnd",
            "key": "f3",
            "window_poll_interval": 0.1,
            "open_key_retry_interval": 0.2,
        },
        timeout=5.0,
    )

    assert query_hwnd is None
    assert round(elapsed_wait["value"], 6) == 2.0
    assert processor.jab.activated == [("Yonyou UClient", "YonyouUWnd", 1.0)]


def test_confirm_step_does_not_use_configured_wait_as_sleep(monkeypatch):
    processor = FakeProcessor()
    workflow = NCSwitchGeneratedWorkflow(processor)
    sleeps = []
    actions = []

    monkeypatch.setattr(
        "core.nc_switch_generated_workflow.time.sleep",
        lambda seconds: sleeps.append(seconds),
    )
    monkeypatch.setattr(
        workflow,
        "run_query_window_step",
        lambda step, dialog_title, dialog_class: actions.append(step),
    )

    wait_timeout = workflow.run_switch_generated_steps(
        open_query={},
        steps=[
            {
                "path": "0.0",
                "name": "确定",
                "role": "push button",
                "wait": 1.0,
            }
        ],
        query_method="hotkey",
    )

    assert wait_timeout == 0.3
    assert actions[0]["wait"] == 0.0
    assert sleeps == []
