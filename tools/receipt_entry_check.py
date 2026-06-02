import argparse
from collections import Counter
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402


def load_json(path):
    with Path(path).open(encoding="utf-8") as file:
        return json.load(file)


def main():
    parser = argparse.ArgumentParser(description="Prepare receipt-entry Excel columns")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--excel-path", default=None)
    parser.add_argument(
        "--write",
        action="store_true",
        help="write output columns and organization names to the workbook",
    )
    args = parser.parse_args()

    config = load_json(args.config)
    workbook = ReceiptEntryWorkbook(config, excel_path=args.excel_path)
    if args.write:
        rows, issues = workbook.ensure_output_columns_and_subjects()
    else:
        rows, issues = workbook.preview_rows()

    by_org = Counter(row.organization_short_name for row in rows)
    by_bank = Counter(row.bank for row in rows)
    print(f"receipt rows from start date: {len(rows)}")
    print(f"issues: {len(issues)}")
    print(f"organizations: {dict(sorted(by_org.items()))}")
    print(f"banks: {dict(sorted(by_bank.items()))}")
    for issue in issues[:20]:
        print(
            f"issue row={issue.excel_row} reason={issue.reason} nc_rows={issue.nc_rows}"
        )


if __name__ == "__main__":
    main()
