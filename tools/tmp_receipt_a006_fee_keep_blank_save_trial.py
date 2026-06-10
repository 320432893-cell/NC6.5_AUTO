# 生命周期：T0 一次性（删除条件：A006 手续费保留空白行保存试验完成后删除）
# 覆盖的业务阶段：收款单自制录入-A006 手续费行-保留空白行真实保存试验
# 依赖的服务/环境：Windows Python、NC 收款单录入页、Java Access Bridge、收款单Excel
# 运行方式：python tools/tmp_receipt_a006_fee_keep_blank_save_trial.py

import os
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ["RECEIPT_SKIP_FEE_EXTRA_ROW_DELETE"] = "1"

from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools import tmp_receipt_two_case_save_run as base_run  # noqa: E402


def choose_a006_fee_case(config):
    workbook = ReceiptEntryWorkbook(config)
    rows, issues, _summary = workbook.build_local_plan(write_sheet=False)
    issue_rows = {issue.excel_row for issue in issues if issue.excel_row is not None}
    candidates = [
        row
        for row in rows
        if row.row not in issue_rows
        and row.organization_code == "A006"
        and row.header_currency_code == "USD"
    ]
    candidates.sort(key=lambda row: (row.receipt_date, row.row), reverse=True)
    if not candidates:
        raise RuntimeError("找不到 A006/USD 有效测试行")
    row = candidates[0]
    return base_run.TestCase(
        name="A006手续费保留空白行保存试验",
        excel_row=row.row,
        document_date=row.receipt_date.isoformat(),
        customer_code=row.customer_code,
        payer_name=row.payer_name,
        source_bank=row.bank,
        bank_label=row.account_label,
        bank_account_no=row.account_no,
        currency=row.currency,
        amount=str(row.raw_amount),
        fee="20.00",
    )


def main():
    base_run.SAVE_ENABLED = True
    base_run.TEST_BANK_ACCOUNT_NO = ""
    base_run.ALLOW_EXISTING_ENTRY_FOR_FIRST_CASE = False
    config = load_config(str(ROOT / "config.json"))
    case = choose_a006_fee_case(config)
    print("A006 手续费保留空白行保存试验：")
    print(
        f"  Sheet1行={case.excel_row} | 日期={case.document_date} | "
        f"客户={case.customer_code} | 银行={case.bank_label} | "
        f"账号={case.bank_account_no} | 金额={case.amount} | 手续费={case.fee}"
    )
    print("  本轮会真实保存；手续费后不删除额外空白行；不查询、不写 Sheet2。")
    print(
        f"请在 {base_run.START_DELAY_SECONDS} 秒内切到 NC【收款单录入】且能看到【新增】的页面..."
    )
    time.sleep(base_run.START_DELAY_SECONDS)
    report = base_run.run_one_case(config, case, allow_existing_entry=False)
    base_run.print_case_summary(report)
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:
        traceback.print_exc()
        raise
