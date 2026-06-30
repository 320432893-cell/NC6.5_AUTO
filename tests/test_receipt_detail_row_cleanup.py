# 生命周期：持久维护
# 覆盖的业务场景：手续费行账户清空前必须读取真实账户列，不能用不完整写入结果误判为空
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_detail_row_cleanup.py

from core import receipt_detail_row_cleanup as cleanup


def test_clear_fee_account_reads_row_when_known_cells_miss_account(monkeypatch):
    calls = []

    def fake_read_row_cells(jab, row_index, located):
        calls.append((jab, row_index, located))
        return {"ok": True}, {"4": ""}

    monkeypatch.setattr(cleanup, "read_row_cells", fake_read_row_cells)

    result = cleanup.clear_fee_account_if_filled(
        object(),
        {"best": {}},
        1,
        known_cells={"1": "手续费", "5": "660305\\财务费用"},
    )

    assert calls
    assert result["ok"] is True
    assert result["skipped"] is True
    assert result["source"] == "known_cells"
