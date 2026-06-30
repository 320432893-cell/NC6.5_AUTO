# 覆盖的业务场景：现场全流程测试 wrapper 一个文件选择保存/不保存/故障恢复/verify 审查，不复制正式业务逻辑
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_full_flow_test_wrappers.py

from tools import receipt_full_flow_save_query_write_test as save_query_write
from core.receipt_full_flow_test_prompts import (
    DEFAULT_EXCEL_NAME,
    build_interactive_args,
    with_default_excel_path,
)


def test_save_query_write_wrapper_appends_save_query_and_sheet2(monkeypatch):
    calls = []

    def fake_run_full_flow(argv):
        calls.append(argv)
        return 1

    monkeypatch.setattr(save_query_write, "run_full_flow", fake_run_full_flow)
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main(["--start-row", "811", "--limit", "3"]) == 1
    assert calls == [
        [
            "--config",
            str(save_query_write.ROOT / "config.json"),
            "--start-row",
            "811",
            "--limit",
            "3",
            "--save",
            "--query-after-save",
            "--write-selected-plan-sheet",
        ]
    ]


def test_wrapper_rejects_controlled_flags(monkeypatch):
    monkeypatch.setattr(save_query_write, "run_full_flow", lambda _argv: 99)

    assert save_query_write.main(["--write-selected-plan-sheet"]) == 2


def test_wrapper_rejects_diagnostic_flags(monkeypatch):
    monkeypatch.setattr(save_query_write, "run_full_flow", lambda _argv: 99)

    assert save_query_write.main(["--pause-after-header-field"]) == 2


def test_wrapper_rejects_json_flag(monkeypatch):
    monkeypatch.setattr(save_query_write, "run_full_flow", lambda _argv: 99)

    assert save_query_write.main(["--json"]) == 2


def test_wrapper_rejects_detail_repair_diagnostic_flag(monkeypatch):
    monkeypatch.setattr(save_query_write, "run_full_flow", lambda _argv: 99)

    assert save_query_write.main(["--diagnose-detail-repair"]) == 2


def test_wrapper_keeps_explicit_config(monkeypatch):
    calls = []
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main(["--config", "custom.json", "--start-row", "811", "--limit", "1"]) == 0
    assert calls == [
        [
            "--config",
            "custom.json",
            "--start-row",
            "811",
            "--limit",
            "1",
            "--save",
            "--query-after-save",
            "--write-selected-plan-sheet",
        ]
    ]


def test_wrapper_runs_from_project_root(monkeypatch, tmp_path):
    calls = []
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main(["--start-row", "811", "--limit", "1"]) == 0
    assert calls
    assert str(save_query_write.ROOT) == str(tmp_path.cwd())


def test_interactive_args_accept_start_row_and_count(monkeypatch):
    answers = iter(["811", "3", "5"])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))

    assert build_interactive_args("test") == [
        "--start-row",
        "811",
        "--limit",
        "3",
        "--start-delay",
        "5",
    ]


def test_query_write_wrapper_prompts_with_three_row_default(monkeypatch):
    calls = []
    answers = iter(["", "", "", ""])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main([]) == 0
    assert calls == [
        [
            "--config",
            str(save_query_write.ROOT / "config.json"),
            "--start-row",
            "811",
            "--limit",
            "3",
            "--start-delay",
            "2",
            "--save",
            "--query-after-save",
            "--write-selected-plan-sheet",
        ]
    ]


def test_wrapper_prompts_no_save_mode(monkeypatch):
    calls = []
    answers = iter(["2", "", "", ""])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main([]) == 0
    assert calls == [
        [
            "--config",
            str(save_query_write.ROOT / "config.json"),
            "--start-row",
            "811",
            "--limit",
            "3",
            "--start-delay",
            "2",
        ]
    ]


def test_wrapper_prompts_recovery_mode(monkeypatch):
    calls = []
    answers = iter(["3", "", "", ""])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main([]) == 0
    assert calls[0][-3:] == [
        "--pause-after-header-field",
        "客户",
        "--diagnose-header-after-pause",
    ]


def test_wrapper_prompts_verify_audit_mode(monkeypatch):
    calls = []
    answers = iter(["4", "", "", ""])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main([]) == 0
    assert "--json" not in calls[0]
    assert calls[0] == [
        "--config",
        str(save_query_write.ROOT / "config.json"),
        "--start-row",
        "811",
        "--limit",
        "3",
        "--start-delay",
        "2",
    ]


def test_wrapper_prompts_modal_alt_c_mode(monkeypatch):
    calls = []
    answers = iter(["5", "", "", ""])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda _argv: (_ for _ in ()).throw(AssertionError("不应跑完整流程")),
    )
    monkeypatch.setattr(
        save_query_write,
        "run_modal_alt_c_test",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main([]) == 0
    assert calls == [
        [
            "--config",
            str(save_query_write.ROOT / "config.json"),
            "--start-row",
            "811",
            "--limit",
            "3",
            "--start-delay",
            "2",
        ]
    ]


def test_wrapper_prompts_detail_repair_drill_mode(monkeypatch):
    calls = []
    answers = iter(["6", "", "", ""])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))
    monkeypatch.setattr(
        save_query_write,
        "run_full_flow",
        lambda argv: calls.append(argv) or 0,
    )
    monkeypatch.setattr(
        save_query_write,
        "with_default_excel_path",
        lambda argv, _root: argv,
    )
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": True},
    )

    assert save_query_write.main([]) == 0
    assert calls == [
        [
            "--config",
            str(save_query_write.ROOT / "config.json"),
            "--start-row",
            "811",
            "--limit",
            "3",
            "--start-delay",
            "2",
            "--diagnose-detail-repair",
        ]
    ]


def test_with_default_excel_path_finds_project_file(tmp_path, monkeypatch):
    monkeypatch.setattr(
        "core.receipt_full_flow_test_prompts.Path.home", lambda: tmp_path
    )
    excel = tmp_path / "project" / DEFAULT_EXCEL_NAME
    excel.parent.mkdir()
    excel.write_bytes(b"placeholder")

    assert with_default_excel_path(["--start-row", "811", "--limit", "1"], excel.parent) == [
        "--start-row",
        "811",
        "--limit",
        "1",
        "--excel-path",
        str(excel),
    ]


def test_with_default_excel_path_keeps_explicit_path(tmp_path):
    assert with_default_excel_path(["--excel-path", "manual.xlsx"], tmp_path) == [
        "--excel-path",
        "manual.xlsx",
    ]


def test_wrapper_reports_jab_environment_failure(monkeypatch):
    monkeypatch.setattr(
        save_query_write,
        "prepare_java_access_bridge",
        lambda: {"ok": False, "reason": "missing jab"},
    )
    monkeypatch.setattr(save_query_write, "run_full_flow", lambda _argv: 99)

    assert save_query_write.main(["--start-row", "811", "--limit", "1"]) == 4
