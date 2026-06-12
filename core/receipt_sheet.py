# 职责：提供收款单 Sheet2 结果表的表头维护、行值生成、排序规则和结果区重写
# 不做什么：不打开/保存 Excel 文件，不查询 NC，不执行 JAB/GUI 动作
# 允许依赖层：收款单数据模型对象和 openpyxl worksheet 风格接口
# 谁不应该 import：底层 JAB/NC workflow 模块不应 import

RESULT_SHEET_HEADERS = [
    "原Sheet1行号",
    "执行主体名称",
    "到款日期",
    "客户编码",
    "币种",
    "银行来款名",
    "实收金额",
    "手续费",
    "总金额",
    "收款银行账户",
    "本地预检状态",
    "异常原因",
]

DEPRECATED_RESULT_SHEET_HEADERS = {
    "执行主体编码",
    "执行主体简称",
    "银行",
    "账户配置ID",
    "异常阶段",
    "异常类型",
    "异常字段",
    "原始值",
    "配置节点",
    "异常说明",
    "处理动作",
    "录入结果",
    "保存结果",
    "后验查询结果",
}


def plan_sheet_row(row, issue, status):
    if row is None:
        base = [""] * 10
    else:
        total_amount = row.raw_amount + row.fee
        base = [
            row.row,
            row.organization_name,
            row.receipt_date.isoformat(),
            row.customer_code,
            row.currency,
            row.payer_name,
            str(row.raw_amount),
            str(row.fee),
            str(total_amount),
            row.account_no,
        ]
    issue_reason = format_plan_issue_reason(issue)
    return [
        *base,
        status,
        issue_reason,
    ]


def format_plan_issue_reason(issue):
    if issue is None:
        return ""
    return f"本地预检：{plan_issue_summary(issue)}"


def plan_issue_summary(issue):
    summaries = {
        "EXCEL_REQUIRED_COLUMN_MISSING": "缺少必需列",
        "EXCEL_START_ROW_INVALID": "起始行配置错误",
        "DATE_INVALID": "到款日期格式错误",
        "PAYER_NAME_EMPTY": "银行来款名为空",
        "AMOUNT_ZERO_OR_NEGATIVE": "实收金额必须大于0",
        "AMOUNT_INVALID": "实收金额格式错误",
        "BANK_EMPTY": "银行为空",
        "BANK_ACCOUNT_NOT_CONFIGURED": "银行未配置",
        "BANK_ACCOUNT_DISABLED": "银行账户已禁用",
        "ORG_NOT_CONFIGURED": "执行主体未配置",
        "CURRENCY_EMPTY": "币种为空",
        "CURRENCY_UNSUPPORTED": "币种不支持",
        "CUSTOMER_CODE_EMPTY": "客户编码为空",
        "FEE_NEGATIVE": "手续费不能小于0",
        "FEE_INVALID": "手续费格式错误",
        "DETAIL_ACCOUNT_CANDIDATE_MISSING": "收款银行账户候选缺失",
        "DUPLICATE_EXCEL_ROWS": "本批存在重复行",
    }
    return summaries.get(issue.issue_type) or str(issue.message or "预检失败")


def ensure_result_sheet_headers(ws, header_row):
    for column in range(ws.max_column, 0, -1):
        value = ws.cell(header_row, column).value
        text = str(value or "").strip()
        if text in DEPRECATED_RESULT_SHEET_HEADERS:
            ws.delete_cols(column)

    columns = {}
    for column in range(1, ws.max_column + 1):
        value = ws.cell(header_row, column).value
        text = str(value or "").strip()
        if text:
            columns[text] = column
    next_column = max(columns.values(), default=0) + 1
    for header in RESULT_SHEET_HEADERS:
        if header in columns:
            continue
        ws.cell(row=header_row, column=next_column, value=header)
        columns[header] = next_column
        next_column += 1
    return columns


def append_plan_sheet_row(ws, columns, row_number, values):
    row_by_header = dict(zip(RESULT_SHEET_HEADERS, values, strict=True))
    for header, value in row_by_header.items():
        ws.cell(row=row_number, column=columns[header], value=value)
    return row_number + 1


def rewrite_plan_sheet(wb, sheet_name, header_row, rows, issues):
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
    else:
        ws = wb[sheet_name]
    columns = ensure_result_sheet_headers(ws, header_row)
    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)
    append_start_row = header_row + 1
    issues_by_row = {}
    global_issues = []
    for issue in issues:
        if issue.excel_row is None:
            global_issues.append(issue)
        else:
            issues_by_row.setdefault(issue.excel_row, []).append(issue)
    rows_by_number = {row.row: row for row in rows}
    emitted_rows = set()
    for row in sorted(rows, key=plan_sheet_sort_key):
        emitted_rows.add(row.row)
        row_issues = issues_by_row.get(row.row, [])
        if row_issues:
            for issue in row_issues:
                append_start_row = append_plan_sheet_row(
                    ws,
                    columns,
                    append_start_row,
                    plan_sheet_row(row, issue, "异常"),
                )
        else:
            append_start_row = append_plan_sheet_row(
                ws,
                columns,
                append_start_row,
                plan_sheet_row(row, None, "通过"),
            )
    orphan_issue_items = [
        (row_number, row_issues)
        for row_number, row_issues in issues_by_row.items()
        if row_number not in emitted_rows
    ]
    orphan_issue_items.sort(
        key=lambda item: orphan_issue_sort_key(item, rows_by_number)
    )
    for row_number, row_issues in orphan_issue_items:
        if row_number in emitted_rows:
            continue
        for issue in row_issues:
            append_start_row = append_plan_sheet_row(
                ws,
                columns,
                append_start_row,
                plan_sheet_row(rows_by_number.get(row_number), issue, "异常"),
            )
    for issue in global_issues:
        append_start_row = append_plan_sheet_row(
            ws,
            columns,
            append_start_row,
            plan_sheet_row(None, issue, "异常"),
        )


def plan_sheet_sort_key(row):
    return (
        str(row.organization_code or ""),
        str(row.organization_name or ""),
        int(row.row or 0),
    )


def orphan_issue_sort_key(item, rows_by_number):
    row_number, _row_issues = item
    row = rows_by_number.get(row_number)
    if row is not None:
        return plan_sheet_sort_key(row)
    return ("ZZZ", "", int(row_number or 0))
