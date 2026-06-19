# 生命周期：持久维护
# 覆盖的业务场景：收款单录入测试的共享配置工厂与 import
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_entry_*.py

from datetime import date
from decimal import Decimal

from openpyxl import Workbook, load_workbook

from core.receipt_config import ReceiptEntryConfig
from core.receipt_entry import ReceiptEntryWorkbook
from core.receipt_matching import (
    ReceiptEntryDryRunMatcher,
    ReceiptEntryMatcher,
    format_receipt_amount_name_mismatch_reason,
    format_receipt_duplicate_reason,
    format_receipt_name_amount_mismatch_reason,
    format_receipt_not_found_reason,
    names_match,
    normalize_counterparty,
)
from core.receipt_models import (
    ReceiptBatchResultRow,
    ReceiptExcelRow,
    ReceiptNCIndexedRow,
    ReceiptNCRow,
)
from core.receipt_nc_extract import ReceiptNCResultExtractor, extract_receipt_nc_rows
from core.receipt_parsing import parse_amount
from core.receipt_parsing import parse_amount as parse_amount_from_new_module
from core.receipt_sheet import RESULT_SHEET_HEADERS


def receipt_config(path="unused.xlsx"):
    return {
        "receipt_entry": {
            "state_label": "收款单录入",
            "excel": {
                "path": str(path),
                "sheet_name": "💸Payments来款通知",
                "header_row": 1,
                "start_row": 2,
                "result_sheet_name": "收款单自动化结果",
                "start_date": "2026-01-01",
                "date_column": "到款日期",
                "payer_name_column": "🟪银行来款名",
                "raw_amount_column": "🟪原始金额",
                "bank_column": "银行",
                "currency_column": "币种",
                "customer_code_column": "客户编码",
                "fee_column": "手续费",
                "organization_column": "主体名称",
                "nc_done_column": "是否NC已做过",
            },
            "validation_policy": {
                "mode": "strict",
                "skip_invalid_rows": False,
            },
            "candidate_check": {
                "recent_months": 2,
                "from_date": None,
                "only_blank_status": True,
            },
            "finance_organizations": [
                {
                    "code": "A001",
                    "name": "上海移为通信技术股份有限公司",
                    "short_name": "移为",
                },
                {
                    "code": "A006",
                    "name": "上海移为通信技术（香港）有限公司",
                    "short_name": "移为香港",
                },
            ],
            "accounts": [
                {
                    "organization_code": "A001",
                    "organization_short_name": "移为",
                    "account_label": "PayPal",
                    "account_no": "paypal",
                },
                {
                    "organization_code": "A006",
                    "organization_short_name": "移为香港",
                    "account_label": "香港花旗",
                    "account_no": "1778667904",
                },
            ],
        }
    }


__all__ = [
    'Decimal',
    'RESULT_SHEET_HEADERS',
    'ReceiptBatchResultRow',
    'ReceiptEntryConfig',
    'ReceiptEntryDryRunMatcher',
    'ReceiptEntryMatcher',
    'ReceiptEntryWorkbook',
    'ReceiptExcelRow',
    'ReceiptNCIndexedRow',
    'ReceiptNCResultExtractor',
    'ReceiptNCRow',
    'Workbook',
    'date',
    'extract_receipt_nc_rows',
    'format_receipt_amount_name_mismatch_reason',
    'format_receipt_duplicate_reason',
    'format_receipt_name_amount_mismatch_reason',
    'format_receipt_not_found_reason',
    'load_workbook',
    'names_match',
    'normalize_counterparty',
    'parse_amount',
    'parse_amount_from_new_module',
    'receipt_config',
]
