from tools import receipt_runtime_monitor as monitor


def snapshot(status, last_error):
    return {
        "ts": "2026-07-01 12:10:20",
        "run_state": {
            "present": True,
            "run_id": "20260701_120237",
            "status": status,
            "stage": "后验查询",
            "step_index": 26,
            "total_steps": 26,
            "current": {"excel_row": 2040},
            "last_error": last_error,
            "events": [],
        },
        "stdout": {
            "new_bytes": 0,
            "size": 0,
        },
        "stdout_since_start": {
            "jab_load_count": 0,
            "fatal_hit_count": 0,
            "thread_alive_true_count": 0,
            "async_timeout_hint_count": 0,
        },
    }


def test_success_snapshot_shows_last_error_as_historical():
    line = monitor.compact_line(snapshot("success", "历史异常"), previous=None)

    assert "lastErr=历史异常" in line
    assert "err=历史异常" not in line


def test_running_snapshot_shows_current_error():
    line = monitor.compact_line(snapshot("running", "当前异常"), previous=None)

    assert "err=当前异常" in line
    assert "lastErr=当前异常" not in line
