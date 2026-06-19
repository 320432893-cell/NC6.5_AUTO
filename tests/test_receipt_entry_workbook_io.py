# 生命周期：持久维护
# 覆盖的业务场景：收款单录入的 Workbook 读写：输出列/科目补齐、NC 状态回写、候选行筛选
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_entry_workbook_io.py


from tests._receipt_entry_helpers import (
    ReceiptEntryWorkbook,
    Workbook,
    date,
    load_workbook,
    receipt_config,
)


def test_ensure_output_columns_and_subjects(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行"])
    ws.append([date(2026, 1, 16), "lamine Mohamed", 225.68, "Paypal"])
    ws.append([date(2025, 12, 31), "old", 1, "Paypal"])
    wb.save(path)
    wb.close()

    rows, candidates, issues = ReceiptEntryWorkbook(
        receipt_config(path)
    ).ensure_output_columns_and_subjects(today=date(2026, 1, 20))

    assert issues == []
    assert len(rows) == 1
    assert candidates == rows
    assert rows[0].organization_code == "A001"

    saved = load_workbook(path)
    ws = saved["💸Payments来款通知"]
    headers = [ws.cell(1, column).value for column in range(1, ws.max_column + 1)]
    assert headers == [
        "到款日期",
        "🟪银行来款名",
        "🟪原始金额",
        "银行",
        "主体名称",
        "是否NC已做过",
    ]
    assert ws.cell(2, 5).value == "上海移为通信技术股份有限公司"
    assert ws.cell(3, 5).value is None
    saved.close()


def test_write_nc_done_statuses_creates_status_column(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行"])
    ws.append([date(2026, 1, 16), "matched", 225.68, "Paypal"])
    ws.append([date(2026, 1, 17), "missing", 100, "Paypal"])
    wb.save(path)
    wb.close()

    result = ReceiptEntryWorkbook(receipt_config(path)).write_nc_done_statuses(
        {2: "已做过", 3: "未做过"}
    )

    assert result == {"updated": 2, "rows": [2, 3]}
    saved = load_workbook(path)
    ws = saved["💸Payments来款通知"]
    assert ws.cell(1, 5).value == "是否NC已做过"
    assert ws.cell(2, 5).value == "已做过"
    assert ws.cell(3, 5).value == "未做过"
    saved.close()


def test_candidate_rows_use_recent_months_and_blank_status(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行", "是否NC已做过"])
    ws.append([date(2026, 3, 31), "old recent excluded", 100, "Paypal", None])
    ws.append([date(2026, 4, 2), "already done", 200, "Paypal", "已做过"])
    ws.append([date(2026, 4, 2), "candidate", 300, "Paypal", None])
    wb.save(path)
    wb.close()

    rows, candidates, issues = ReceiptEntryWorkbook(receipt_config(path)).preview_rows(
        today=date(2026, 6, 2)
    )

    assert issues == []
    assert len(rows) == 3
    assert [row.payer_name for row in candidates] == ["candidate"]


def test_candidate_from_date_overrides_recent_months(tmp_path):
    path = tmp_path / "payments.xlsx"
    config = receipt_config(path)
    config["receipt_entry"]["candidate_check"]["from_date"] = "2026-05-01"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行"])
    ws.append([date(2026, 4, 30), "old", 100, "Paypal"])
    ws.append([date(2026, 5, 1), "candidate", 200, "Paypal"])
    wb.save(path)
    wb.close()

    rows, candidates, issues = ReceiptEntryWorkbook(config).preview_rows(
        today=date(2026, 6, 2)
    )

    assert issues == []
    assert len(rows) == 2
    assert [row.payer_name for row in candidates] == ["candidate"]
