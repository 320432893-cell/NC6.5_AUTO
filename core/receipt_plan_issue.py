# 职责：提供收款单本地预检问题构造、重复检测、计划摘要和 worksheet 单元格读取小工具
# 不做什么：不打开/保存 Excel，不写 Sheet2，不查询 NC，不执行 JAB/GUI 动作
# 允许依赖层：收款单数据模型和解析/格式化纯函数
# 谁不应该 import：底层 JAB/NC workflow 模块不应 import

from core.receipt_models import ReceiptPlanIssue
from core.receipt_parsing import format_receipt_value


def detect_duplicate_rows(rows):
    grouped = {}
    for row in rows:
        grouped.setdefault(row.duplicate_key, []).append(row)
    issues = []
    for key, group in grouped.items():
        if len(group) <= 1:
            continue
        row_numbers = [row.row for row in group]
        key_text = " + ".join(key)
        for row in group:
            issues.append(
                ReceiptPlanIssue(
                    excel_row=row.row,
                    stage="本地重复校验",
                    issue_type="DUPLICATE_EXCEL_ROWS",
                    field="重复键",
                    raw_value=key_text,
                    config_node="local.duplicate_key",
                    message=(
                        f"本批 Sheet1 存在重复行；重复键={key_text}；"
                        f"重复原行号={row_numbers}；为避免重复制单，整组未录入。"
                    ),
                    action="跳过重复组",
                )
            )
    return issues


def summarize_plan(rows, issues, validation_policy):
    duplicate_rows = {
        issue.excel_row
        for issue in issues
        if issue.issue_type == "DUPLICATE_EXCEL_ROWS" and issue.excel_row
    }
    runnable = [row for row in rows if row.row not in duplicate_rows]
    grouped = {}
    for row in runnable:
        grouped.setdefault(row.organization_code, []).append(row.row)
    return {
        "rows": len(rows),
        "issues": len(issues),
        "runnable_rows": len(runnable),
        "duplicate_rows": sorted(duplicate_rows),
        "organizations": {key: value for key, value in sorted(grouped.items())},
        "validation_policy": validation_policy,
        "can_run": not issues or validation_policy == "skip_invalid_rows",
    }


def read_optional_cell(ws, row, column):
    if not column:
        return ""
    value = ws.cell(row, column).value
    return "" if value is None else str(value).strip()


def plan_issue(
    excel_row,
    stage,
    issue_type,
    field,
    raw_value,
    config_node,
    message,
    action="跳过本行",
):
    return ReceiptPlanIssue(
        excel_row=excel_row,
        stage=stage,
        issue_type=issue_type,
        field=field,
        raw_value=format_receipt_value(raw_value),
        config_node=config_node,
        message=message,
        action=action,
    )
