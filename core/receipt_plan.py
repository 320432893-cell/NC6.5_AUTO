# 职责：构建收款单本地预检计划行
# 不做什么：不打开/保存 Excel，不写 Sheet2，不做重复检测/摘要，不查询 NC，不执行 JAB/GUI 动作
# 允许依赖层：收款单配置对象、数据模型、解析/格式化纯函数和 worksheet 读接口
# 谁不应该 import：底层 JAB 操作模块不应 import

from decimal import Decimal

from core.errors import WorkflowStateError
from core.receipt_models import ReceiptPlanIssue, ReceiptPlanRow
from core.receipt_matching import normalize_counterparty
from core.receipt_parsing import (
    make_receipt_duplicate_key,
    normalize_lookup_key,
    normalize_receipt_currency,
    parse_amount,
    parse_date,
)
from core.receipt_plan_issue import plan_issue, read_optional_cell


def build_plan_rows(config, ws, columns):
    issues = []
    rows = []
    required = [
        config.date_column,
        config.payer_name_column,
        config.raw_amount_column,
        config.bank_column,
        config.currency_column,
        config.customer_code_column,
    ]
    missing = [name for name in required if name not in columns]
    for name in missing:
        issues.append(
            ReceiptPlanIssue(
                excel_row=None,
                stage="配置识别",
                issue_type="EXCEL_REQUIRED_COLUMN_MISSING",
                field=name,
                raw_value="",
                config_node="receipt_entry.excel",
                message=f"Sheet1 缺少必需列 {name!r}",
                action="停止整批",
            )
        )
    if missing:
        return [], issues
    if config.start_row <= config.header_row:
        issues.append(
            ReceiptPlanIssue(
                excel_row=None,
                stage="配置识别",
                issue_type="EXCEL_START_ROW_INVALID",
                field="start_row",
                raw_value=str(config.start_row),
                config_node="receipt_entry.excel.start_row",
                message=(
                    "receipt_entry.excel.start_row 必须大于 "
                    f"header_row={config.header_row}"
                ),
                action="停止整批",
            )
        )
        return [], issues

    for row_index in range(config.start_row, ws.max_row + 1):
        raw_date = ws.cell(row_index, columns[config.date_column]).value
        if raw_date in (None, ""):
            continue
        row, row_issues = build_plan_row(config, ws, columns, row_index)
        issues.extend(row_issues)
        if row is not None:
            rows.append(row)
    return rows, issues


def build_plan_row(config, ws, columns, row_index):
    row_issues = []
    raw_date = ws.cell(row_index, columns[config.date_column]).value
    payer_name = read_optional_cell(
        ws, row_index, columns.get(config.payer_name_column)
    )
    raw_amount = ws.cell(row_index, columns[config.raw_amount_column]).value
    bank = read_optional_cell(ws, row_index, columns.get(config.bank_column))
    currency = read_optional_cell(ws, row_index, columns.get(config.currency_column))
    customer_code = read_optional_cell(
        ws, row_index, columns.get(config.customer_code_column)
    )
    fee_raw = (
        ws.cell(row_index, columns[config.fee_column]).value
        if config.fee_column in columns
        else None
    )

    receipt_date = None
    amount = None
    fee = Decimal("0.00")
    account = None
    organization = None

    try:
        receipt_date = parse_date(raw_date)
    except ValueError:
        row_issues.append(
            plan_issue(
                row_index,
                "本地数据校验",
                "DATE_INVALID",
                config.date_column,
                raw_date,
                "receipt_entry.excel.date_column",
                f"到款日期格式无法识别: {raw_date!r}",
            )
        )

    if not payer_name:
        row_issues.append(
            plan_issue(
                row_index,
                "本地数据校验",
                "PAYER_NAME_EMPTY",
                config.payer_name_column,
                payer_name,
                "receipt_entry.excel.payer_name_column",
                "银行来款名为空，无法用于后验匹配",
            )
        )

    try:
        amount = parse_amount(raw_amount)
        if amount <= 0:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "AMOUNT_ZERO_OR_NEGATIVE",
                    config.raw_amount_column,
                    raw_amount,
                    "receipt_entry.excel.raw_amount_column",
                    f"原始金额必须大于 0，当前为 {amount}",
                )
            )
    except ValueError as exc:
        row_issues.append(
            plan_issue(
                row_index,
                "本地数据校验",
                "AMOUNT_INVALID",
                config.raw_amount_column,
                raw_amount,
                "receipt_entry.excel.raw_amount_column",
                f"Excel第{row_index}行 金额列：{exc}",
            )
        )

    if not bank:
        row_issues.append(
            plan_issue(
                row_index,
                "配置识别",
                "BANK_EMPTY",
                config.bank_column,
                bank,
                "receipt_entry.excel.bank_column",
                "银行为空，无法匹配 receipt_entry.accounts",
            )
        )
    else:
        account = config.account_for_bank(bank)
        if not account:
            row_issues.append(
                plan_issue(
                    row_index,
                    "配置识别",
                    "BANK_ACCOUNT_NOT_CONFIGURED",
                    config.bank_column,
                    bank,
                    "receipt_entry.accounts[*].account_label/aliases/excel_bank_aliases",
                    (
                        f"Sheet1 银行={bank!r} 未匹配任何账户配置；"
                        f"可用配置值={config.account_lookup_labels()}"
                    ),
                )
            )
        elif not account.enabled:
            row_issues.append(
                plan_issue(
                    row_index,
                    "配置识别",
                    "BANK_ACCOUNT_DISABLED",
                    config.bank_column,
                    bank,
                    f"receipt_entry.accounts[{account.id}].enabled",
                    f"银行={bank!r} 匹配到账户 {account.id!r}，但账户已禁用",
                )
            )
        else:
            organization = config.organizations.get(account.organization_code)
            if not organization:
                row_issues.append(
                    plan_issue(
                        row_index,
                        "配置识别",
                        "ORG_NOT_CONFIGURED",
                        "organization_code",
                        account.organization_code,
                        "receipt_entry.finance_organizations",
                        (
                            f"账户 {account.id!r} 的 organization_code="
                            f"{account.organization_code!r} 不存在"
                        ),
                    )
                )

    currency_name = normalize_receipt_currency(currency)
    if not currency:
        row_issues.append(
            plan_issue(
                row_index,
                "本地数据校验",
                "CURRENCY_EMPTY",
                config.currency_column,
                currency,
                "receipt_entry.excel.currency_column",
                "币种为空，无法选择 NC 明细币种和账号候选",
            )
        )
    elif not currency_name:
        row_issues.append(
            plan_issue(
                row_index,
                "本地数据校验",
                "CURRENCY_UNSUPPORTED",
                config.currency_column,
                currency,
                "receipt_entry.excel.currency_column",
                f"币种={currency!r} 不在支持列表 USD/RMB/CNY/美元/人民币",
            )
        )

    if not customer_code:
        row_issues.append(
            plan_issue(
                row_index,
                "本地数据校验",
                "CUSTOMER_CODE_EMPTY",
                config.customer_code_column,
                customer_code,
                "receipt_entry.excel.customer_code_column",
                "客户编码为空，不能写收款单表头客户字段",
            )
        )

    if fee_raw not in (None, ""):
        try:
            fee = parse_amount(fee_raw)
            if fee < 0:
                row_issues.append(
                    plan_issue(
                        row_index,
                        "本地数据校验",
                        "FEE_NEGATIVE",
                        config.fee_column,
                        fee_raw,
                        "receipt_entry.excel.fee_column",
                        f"手续费不能小于 0，当前为 {fee}",
                    )
                )
        except ValueError as exc:
            row_issues.append(
                plan_issue(
                    row_index,
                    "本地数据校验",
                    "FEE_INVALID",
                    config.fee_column,
                    fee_raw,
                    "receipt_entry.excel.fee_column",
                    str(exc).replace("原始金额", "手续费"),
                )
            )

    if account and currency_name and not account.nc_candidates(currency_name):
        row_issues.append(
            plan_issue(
                row_index,
                "配置识别",
                "DETAIL_ACCOUNT_CANDIDATE_MISSING",
                config.bank_column,
                bank,
                f"receipt_entry.accounts[{account.id}].nc_candidates_by_currency",
                (
                    f"账户 {account.id!r} 在币种 {currency_name!r} 下没有可用 "
                    "NC 账号候选"
                ),
            )
        )

    if row_issues:
        return None, row_issues
    if (
        receipt_date is None
        or amount is None
        or account is None
        or organization is None
        or currency_name is None
    ):
        raise WorkflowStateError("收款单本地预检内部状态不完整，无法生成运行计划")
    duplicate_key = make_receipt_duplicate_key(
        organization.code,
        receipt_date,
        bank,
        currency_name,
        customer_code,
        payer_name,
        amount,
        normalize_lookup_key,
        normalize_counterparty,
    )
    return (
        ReceiptPlanRow(
            row=row_index,
            receipt_date=receipt_date,
            payer_name=payer_name,
            raw_amount=amount,
            bank=bank,
            currency=currency_name,
            customer_code=customer_code,
            fee=fee,
            organization_code=organization.code,
            organization_name=organization.name,
            organization_short_name=organization.short_name,
            account_id=account.id,
            account_label=account.account_label,
            account_no=account.account_no,
            header_currency_code=account.header_currency_code,
            duplicate_key=duplicate_key,
        ),
        [],
    )
