# 生命周期：临时，三笔真实保存验收完成后删除
# 覆盖的业务场景：从正式收款计划中挑选手续费、无手续费人民币、香港移为三类测试行
# 依赖的服务/环境：本地 Python + openpyxl，不依赖 NC/GUI/JAB
# 运行方式：py -3.11 tools\tmp_select_receipt_three_rows.py

import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402
from core.utils import load_config  # noqa: E402


def main():
    config = load_config("config.json")
    rows, issues, summary = ReceiptEntryWorkbook(config).build_local_plan(False)
    issue_rows = {issue.excel_row for issue in issues if issue.excel_row is not None}
    runnable = [row for row in rows if row.row not in issue_rows]

    picked = []
    picked_numbers = set()

    def pick(label, predicate):
        for row in runnable:
            if row.row in picked_numbers:
                continue
            if predicate(row):
                picked.append((label, row))
                picked_numbers.add(row.row)
                return

    pick("fee", lambda row: row.fee > 0)
    pick(
        "rmb_no_fee",
        lambda row: (
            row.fee == 0
            and row.currency in {"人民币", "CNY"}
            and row.organization_code != "A006"
        ),
    )
    pick("hongkong", lambda row: row.organization_code == "A006")

    result = {
        "summary": summary,
        "issue_count": len(issues),
        "picked": [
            {
                "label": label,
                "row": row.row,
                "org": row.organization_short_name,
                "date": row.receipt_date.isoformat(),
                "payer": row.payer_name,
                "amount": str(row.raw_amount),
                "fee": str(row.fee),
                "currency": row.currency,
                "bank": row.bank,
                "customer": row.customer_code,
                "account": row.account_no,
            }
            for label, row in picked
        ],
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    return 0 if len(picked) == 3 else 2


if __name__ == "__main__":
    raise SystemExit(main())
