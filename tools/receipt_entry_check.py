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
        "--recent-months",
        type=int,
        default=None,
        help="override receipt_entry.candidate_check.recent_months",
    )
    parser.add_argument(
        "--from-date",
        default=None,
        help="override receipt_entry.candidate_check.from_date, format YYYY-MM-DD",
    )
    parser.add_argument(
        "--write",
        action="store_true",
        help="write output columns and organization names to the workbook",
    )
    args = parser.parse_args()

    config = load_json(args.config)
    candidate_cfg = config.setdefault("receipt_entry", {}).setdefault(
        "candidate_check", {}
    )
    if args.recent_months is not None:
        candidate_cfg["recent_months"] = args.recent_months
        candidate_cfg["from_date"] = None
    if args.from_date is not None:
        candidate_cfg["from_date"] = args.from_date
    workbook = ReceiptEntryWorkbook(config, excel_path=args.excel_path)
    if args.write:
        rows, candidates, issues = workbook.ensure_output_columns_and_subjects()
    else:
        rows, candidates, issues = workbook.preview_rows()

    by_org = Counter(row.organization_short_name for row in rows)
    by_bank = Counter(row.bank for row in rows)
    candidate_by_org = Counter(row.organization_short_name for row in candidates)
    candidate_by_bank = Counter(row.bank for row in candidates)
    print(f"receipt rows from start date: {len(rows)}")
    print(f"candidate rows: {len(candidates)}")
    print(f"issues: {len(issues)}")
    print(f"organizations: {dict(sorted(by_org.items()))}")
    print(f"banks: {dict(sorted(by_bank.items()))}")
    print(f"candidate organizations: {dict(sorted(candidate_by_org.items()))}")
    print(f"candidate banks: {dict(sorted(candidate_by_bank.items()))}")
    for issue in issues[:20]:
        print(
            f"issue row={issue.excel_row} reason={issue.reason} nc_rows={issue.nc_rows}"
        )


if __name__ == "__main__":
    main()
