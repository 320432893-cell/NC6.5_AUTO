# 生命周期：持久维护
# 覆盖的业务场景：收款单录入的计划构建：本地计划/sheet2 机器表/批次结果表
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_entry_plan_build.py


from tests._receipt_entry_helpers import (
    ReceiptBatchResultRow,
    ReceiptEntryWorkbook,
    Workbook,
    date,
    load_workbook,
    receipt_config,
)


def test_build_local_plan_reports_precise_local_issues(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(
        [
            "到款日期",
            "🟪银行来款名",
            "🟪原始金额",
            "银行",
            "币种",
            "客户编码",
            "手续费",
        ]
    )
    ws.append([date(2026, 6, 1), "DUP INC", 100, "Paypal", "USD", "YW001", None])
    ws.append([date(2026, 6, 1), "DUP INC", 100, "Paypal", "USD", "YW001", None])
    ws.append([date(2026, 6, 2), "UNKNOWN", 200, "不存在银行", "USD", "YW002", None])
    ws.append([date(2026, 6, 3), "NO CUSTOMER", 300, "Paypal", "USD", "", None])
    wb.save(path)
    wb.close()

    rows, issues, summary = ReceiptEntryWorkbook(
        receipt_config(path)
    ).build_local_plan()

    assert [row.row for row in rows] == [2, 3]
    assert summary["runnable_rows"] == 0
    assert summary["duplicate_rows"] == [2, 3]
    issue_types = [(issue.excel_row, issue.issue_type) for issue in issues]
    assert (2, "DUPLICATE_EXCEL_ROWS") in issue_types
    assert (3, "DUPLICATE_EXCEL_ROWS") in issue_types
    assert (4, "BANK_ACCOUNT_NOT_CONFIGURED") in issue_types
    assert (5, "CUSTOMER_CODE_EMPTY") in issue_types
    bank_issue = next(
        issue for issue in issues if issue.issue_type == "BANK_ACCOUNT_NOT_CONFIGURED"
    )
    assert bank_issue.stage == "配置识别"
    assert bank_issue.field == "银行"
    assert bank_issue.raw_value == "不存在银行"
    assert (
        bank_issue.config_node
        == "receipt_entry.accounts[*].account_label/aliases/excel_bank_aliases"
    )
    assert "未匹配任何账户配置" in bank_issue.message


def test_build_local_plan_writes_machine_sheet2(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(
        [
            "到款日期",
            "🟪银行来款名",
            "🟪原始金额",
            "银行",
            "币种",
            "客户编码",
            "手续费",
        ]
    )
    ws.append([date(2026, 6, 1), "OK INC", 100, "Paypal", "USD", "YW001", 0])
    wb.save(path)
    wb.close()

    rows, issues, summary = ReceiptEntryWorkbook(receipt_config(path)).build_local_plan(
        write_sheet=True
    )

    assert issues == []
    assert summary["runnable_rows"] == 1
    saved = load_workbook(path)
    ws = saved["收款单自动化结果"]
    headers = [ws.cell(1, column).value for column in range(1, ws.max_column + 1)]
    assert headers[:15] == [
        "原Sheet1行号",
        "执行主体名称",
        "到款日期",
        "🟪银行来款名",
        "客户编码",
        "NC客户名称",
        "🟪原始金额",
        "手续费",
        "🟪到账金额",
        "币种",
        "银行",
        "收款银行账户",
        "本地预检状态",
        "后验核对状态",
        "异常原因",
    ]
    assert ws.cell(2, 1).value == rows[0].row
    assert ws.cell(2, 2).value == "上海移为通信技术股份有限公司"
    assert ws.cell(2, 13).value == "通过"
    assert ws.cell(2, 14).value in (None, "")
    assert ws.cell(2, 15).value in (None, "")
    saved.close()


def test_build_local_plan_rewrites_sheet2_without_duplicate_appends(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(
        [
            "到款日期",
            "🟪银行来款名",
            "🟪原始金额",
            "银行",
            "币种",
            "客户编码",
            "手续费",
        ]
    )
    ws.append([date(2026, 6, 1), "OK INC", 100, "Paypal", "USD", "YW001", 0])
    wb.save(path)
    wb.close()

    workbook = ReceiptEntryWorkbook(receipt_config(path))
    workbook.build_local_plan(write_sheet=True)
    workbook.build_local_plan(write_sheet=True)

    saved = load_workbook(path)
    ws = saved["收款单自动化结果"]
    columns = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
    assert ws.max_row == 2
    assert ws.cell(2, 1).value == 2
    assert ws.cell(2, columns["本地预检状态"]).value == "通过"
    saved.close()


def test_write_batch_result_sheet_uses_business_columns_and_sorting(tmp_path):
    path = tmp_path / "payments.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(
        [
            "到款日期",
            "🟪银行来款名",
            "🟪原始金额",
            "银行",
            "币种",
            "客户编码",
            "手续费",
        ]
    )
    ws.append([date(2026, 6, 2), "LATE", 100, "Paypal", "USD", "YW001", 0])
    ws.append([date(2026, 6, 1), "EARLY", 200, "香港花旗", "USD", "YW002", 3])
    wb.save(path)
    wb.close()

    rows, _issues, _summary = ReceiptEntryWorkbook(
        receipt_config(path)
    ).build_local_plan()
    by_row = {row.row: row for row in rows}
    ReceiptEntryWorkbook(receipt_config(path)).write_batch_result_sheet(
        [
            ReceiptBatchResultRow(
                plan_row=by_row[2],
                local_status="通过",
                nc_customer_name="NC LATE",
                nc_document_no="SK2",
            ),
            ReceiptBatchResultRow(
                plan_row=by_row[3],
                local_status="通过",
                nc_customer_name="NC EARLY",
                nc_document_no="SK1",
            ),
        ]
    )

    saved = load_workbook(path)
    ws = saved["收款单自动化结果"]
    columns = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
    assert ws.cell(2, 1).value == "主体：上海移为通信技术股份有限公司"
    assert ws.cell(2, 1).fill.fgColor.rgb == "00D9EAF7"
    assert ws.cell(3, 1).value == 2
    assert ws.cell(3, 4).value == "LATE"
    assert ws.cell(3, columns["后验核对状态"]).value == "后验通过"
    assert ws.cell(4, 1).value == "主体：上海移为通信技术（香港）有限公司"
    assert ws.cell(5, 1).value == 3
    assert ws.cell(5, 4).value == "EARLY"
    assert ws.cell(5, 6).value == "NC EARLY"
    assert ws.cell(5, 7).value == "203.00"
    assert ws.cell(5, 8).value == "3.00"
    assert ws.cell(5, 9).value == "200.00"
    assert ws.cell(5, columns["后验核对状态"]).value == "后验通过"
    saved.close()


def test_build_local_plan_writes_multiple_global_issues_to_sheet2(tmp_path):
    path = tmp_path / "payments.xlsx"
    config = receipt_config(path)
    config["receipt_entry"]["excel"]["start_row"] = 1
    wb = Workbook()
    ws = wb.active
    ws.title = "💸Payments来款通知"
    ws.append(["到款日期", "🟪银行来款名", "🟪原始金额", "银行"])
    ws.append([date(2026, 6, 1)])
    wb.save(path)
    wb.close()

    _rows, issues, summary = ReceiptEntryWorkbook(config).build_local_plan(
        write_sheet=True
    )

    assert summary["issues"] >= 2
    assert {issue.issue_type for issue in issues} == {"EXCEL_REQUIRED_COLUMN_MISSING"}
    saved = load_workbook(path)
    ws = saved["收款单自动化结果"]
    columns = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
    assert ws.max_row == 3
    assert [ws.cell(row, columns["本地预检状态"]).value for row in range(2, 4)] == [
        "异常"
    ] * 2
    assert all(ws.cell(row, columns["异常原因"]).value for row in range(2, 4))
    saved.close()
