# 生命周期：持久维护
# 覆盖的业务场景：引擎模式(NC_ENGINE_MODE=1)下入口层抑制面向操作员的旁白/确认/交互，
#   同时机器结果信封(末段 JSON)在任何模式下照常输出，且不调用 input()
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB；run_state 重定向到 tmp_path
# 运行方式：.venv/bin/python -m pytest -q tests/test_engine_mode_suppression.py

import json
import types

import tools.jab_batch as jab_batch
import tools.receipt_detail_entry as detail_entry
import tools.receipt_full_flow_entry as full_flow


def _no_input(*_args, **_kwargs):
    raise AssertionError("引擎模式下不应调用 input()（子进程无 TTY 会挂）")


# --------------------------------------------------------------------------
# jab_batch.py：plan 旁白 / generate 确认
# --------------------------------------------------------------------------


class _FakeProcessor:
    def __init__(self, *_args, **_kwargs):
        self.run_state = types.SimpleNamespace(set_stage=lambda *_a, **_k: None)

    def dry_run(self, *_args, **_kwargs):
        return {
            "matches": [],
            "issues": [],
            "batches": [],
            "parse_errors": [],
        }

    def finish_run_state(self, *_args, **_kwargs):
        return None

    def close(self):
        return None


def _patch_jab_batch(monkeypatch, tmp_path):
    monkeypatch.setenv("NC_RUNTIME_DIR", str(tmp_path))
    monkeypatch.setattr(jab_batch, "load_config", lambda _path: {})
    monkeypatch.setattr(jab_batch, "JABBatchProcessor", _FakeProcessor)
    monkeypatch.setattr("builtins.input", _no_input)


def test_jab_batch_plan_suppresses_narration_but_keeps_envelope_in_engine_mode(
    monkeypatch, tmp_path, capsys
):
    monkeypatch.setenv("NC_ENGINE_MODE", "1")
    _patch_jab_batch(monkeypatch, tmp_path)
    monkeypatch.setattr(jab_batch.sys, "argv", ["jab_batch.py", "plan", "--json"])

    jab_batch.main()

    out = capsys.readouterr().out
    # ① 旁白被抑制
    assert "JAB 批量计划" not in out
    # ② 机器信封仍照常输出（末段 JSON 可解析）
    envelope = json.loads(out.strip().splitlines()[-1])
    assert envelope["command"] == "plan"
    assert envelope["ok"] is True
    assert envelope["summary"]["total"] == 0


def test_jab_batch_plan_keeps_narration_when_not_engine_mode(
    monkeypatch, tmp_path, capsys
):
    monkeypatch.delenv("NC_ENGINE_MODE", raising=False)
    _patch_jab_batch(monkeypatch, tmp_path)
    monkeypatch.setattr(jab_batch.sys, "argv", ["jab_batch.py", "plan", "--json"])

    jab_batch.main()

    out = capsys.readouterr().out
    # CLI 直跑：旁白照常打印，信封也照常输出
    assert "JAB 批量计划" in out
    envelope = json.loads(out.strip().splitlines()[-1])
    assert envelope["command"] == "plan"


def test_jab_batch_generate_skips_confirm_input_in_engine_mode(
    monkeypatch, tmp_path, capsys
):
    # generate 未给 --yes：引擎模式必须跳过 input()（否则子进程无 TTY 挂起）。
    monkeypatch.setenv("NC_ENGINE_MODE", "1")
    monkeypatch.setenv("NC_RUNTIME_DIR", str(tmp_path))
    monkeypatch.setattr(jab_batch, "load_config", lambda _path: {})
    monkeypatch.setattr("builtins.input", _no_input)

    saved_calls = {"count": 0}

    class _GenProcessor(_FakeProcessor):
        def generate_and_save(self, *_args, **_kwargs):
            saved_calls["count"] += 1
            return 0

    monkeypatch.setattr(jab_batch, "JABBatchProcessor", _GenProcessor)
    monkeypatch.setattr(jab_batch.sys, "argv", ["jab_batch.py", "generate", "--json"])

    jab_batch.main()

    out = capsys.readouterr().out
    # ③ input 未被调用（_no_input 会断言失败）；旁白被抑制；业务照常执行
    assert "即将真实点击 NC" not in out
    assert saved_calls["count"] == 1
    envelope = json.loads(out.strip().splitlines()[-1])
    assert envelope["command"] == "generate"


# --------------------------------------------------------------------------
# receipt_full_flow_entry.py：confirm_save 高风险旁白
# --------------------------------------------------------------------------


def test_confirm_save_suppresses_narration_in_engine_mode(monkeypatch, capsys):
    monkeypatch.setenv("NC_ENGINE_MODE", "1")
    monkeypatch.setattr("builtins.input", _no_input)

    class SaveArgs:
        yes_i_understand = True

    assert full_flow.confirm_save(SaveArgs()) is None
    out = capsys.readouterr().out
    assert out == ""  # 高风险旁白被抑制


def test_confirm_save_prints_narration_when_not_engine_mode(monkeypatch, capsys):
    monkeypatch.delenv("NC_ENGINE_MODE", raising=False)

    class SaveArgs:
        yes_i_understand = True

    full_flow.confirm_save(SaveArgs())
    out = capsys.readouterr().out
    assert "高风险" in out  # CLI 直跑仍打印操作员旁白


# --------------------------------------------------------------------------
# receipt_detail_entry.py：操作员说明 / 切窗提示 / 摘要 / 等待回车
# --------------------------------------------------------------------------


def _patch_detail_entry(monkeypatch, tmp_path):
    monkeypatch.setenv("NC_RUNTIME_DIR", str(tmp_path))
    monkeypatch.setattr(detail_entry, "load_config", lambda _path: {})
    monkeypatch.setattr(
        detail_entry,
        "get_test_account",
        lambda _config, _bank: types.SimpleNamespace(account_no="FTE-TEST-001"),
    )

    def fake_trial(_config, _account, _args, report, _timings, _recorder=None):
        report["ok"] = True
        report["mode"] = "main-line"
        report["fill_steps"] = [{"name": "收款银行账户", "ok": True}]
        return 0

    monkeypatch.setattr(detail_entry, "run_detail_trial", fake_trial)
    monkeypatch.setattr("builtins.input", _no_input)


def test_detail_entry_suppresses_operator_narration_in_engine_mode(
    monkeypatch, tmp_path, capsys
):
    monkeypatch.setenv("NC_ENGINE_MODE", "1")
    _patch_detail_entry(monkeypatch, tmp_path)

    result = detail_entry.main(["--start-delay", "0", "--json"])

    out = capsys.readouterr().out
    # ① 操作员说明 / 切窗提示 / 开始测试 / 人类摘要 均不出现
    assert "测试功能" not in out  # print_header 大段说明
    assert "切到 NC" not in out
    assert "开始测试" not in out
    assert "测试结果" not in out  # print_summary 人类摘要
    # ② 机器信封仍照常输出
    envelope = json.loads(out.strip().splitlines()[-1])
    assert envelope["command"] == "receipt-detail"
    assert envelope["ok"] is True
    # ③ input 未被调用（_no_input 否则会抛 AssertionError）
    assert result == 0


def test_detail_entry_keeps_narration_when_not_engine_mode(
    monkeypatch, tmp_path, capsys
):
    monkeypatch.delenv("NC_ENGINE_MODE", raising=False)
    _patch_detail_entry(monkeypatch, tmp_path)

    # 非引擎模式仍给 --no-wait，避免 CLI 路径触发 input("按回车退出...")。
    result = detail_entry.main(["--start-delay", "0", "--json", "--no-wait"])

    out = capsys.readouterr().out
    # CLI 直跑：操作员说明与人类摘要照常打印，信封也照常输出
    assert "测试功能" in out
    assert "开始测试" in out
    assert "测试结果" in out
    envelope = json.loads(out.strip().splitlines()[-1])
    assert envelope["command"] == "receipt-detail"
    assert result == 0
