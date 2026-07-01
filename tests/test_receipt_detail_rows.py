from decimal import Decimal

from core import receipt_detail_rows as rows


def test_fee_business_type_failure_rewrites_once(monkeypatch):
    calls = []

    monkeypatch.setattr(
        rows,
        "read_fee_prepare_row_count",
        lambda *_args, **_kwargs: {"ok": True, "row_count": 2},
    )
    monkeypatch.setattr(
        rows,
        "delete_extra_row_if_present",
        lambda *_args, **_kwargs: {"ok": True, "skipped": True},
    )

    def fake_write(_jab, _business, _located, fields=None, row_index=0, **_kwargs):
        calls.append([field["name"] for field in fields])
        if len(calls) == 1:
            return [
                {
                    "ok": False,
                    "name": "收款业务类型",
                    "target": {"row": row_index, "col": 1},
                    "reason": "即时校验未匹配：字段=收款业务类型，期望='手续费'，实际='货款'",
                    "actual": "货款",
                }
            ]
        return [
            {
                "ok": True,
                "name": "收款业务类型",
                "target": {"row": row_index, "col": 1},
                "actual": "手续费",
            }
        ]

    monkeypatch.setattr(rows, "write_detail_line_by_screen", fake_write)

    add_row, steps, clear_account, delete_extra = rows.run_fee_only(
        object(),
        {"best": {"row_count": 2, "col_count": 25, "window": {"hwnd": 1}}},
        Decimal("10.00"),
    )

    assert add_row["ok"] is True
    assert clear_account["ok"] is True
    assert delete_extra["ok"] is True
    assert len(calls) == 2
    assert calls[1] == ["收款业务类型"]
    assert steps[0]["ok"] is True
    assert steps[0]["repair_of"] == "fee_business_type"
