import argparse
from collections import Counter
import json
from pathlib import Path
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.receipt_entry import ReceiptEntryWorkbook  # noqa: E402


def load_json(path):
    with Path(path).open(encoding="utf-8") as file:
        return json.load(file)


def main():
    parser = argparse.ArgumentParser(description="Build receipt-entry local run plan")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--excel-path", default=None)
    parser.add_argument(
        "--recent-months",
        type=int,
        default=None,
        help="legacy preview only: override receipt_entry.candidate_check.recent_months",
    )
    parser.add_argument(
        "--from-date",
        default=None,
        help="legacy preview only: override receipt_entry.candidate_check.from_date",
    )
    parser.add_argument(
        "--write",
        action="store_true",
        help="write the machine-generated Sheet2 local validation result",
    )
    parser.add_argument(
        "--legacy-candidates",
        action="store_true",
        help="use the old recent-month/blank-status candidate preview",
    )
    parser.add_argument(
        "--validation-mode",
        choices=("strict", "skip_invalid_rows"),
        default=None,
        help="override receipt_entry.validation_policy.mode for this run",
    )
    args = parser.parse_args()

    config = load_json(args.config)
    if args.validation_mode:
        policy = config.setdefault("receipt_entry", {}).setdefault(
            "validation_policy", {}
        )
        policy["mode"] = args.validation_mode
        policy["skip_invalid_rows"] = args.validation_mode == "skip_invalid_rows"

    candidate_cfg = config.setdefault("receipt_entry", {}).setdefault(
        "candidate_check", {}
    )
    if args.recent_months is not None:
        candidate_cfg["recent_months"] = args.recent_months
        candidate_cfg["from_date"] = None
    if args.from_date is not None:
        candidate_cfg["from_date"] = args.from_date
    workbook = ReceiptEntryWorkbook(config, excel_path=args.excel_path)
    if args.legacy_candidates:
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
                f"issue row={issue.excel_row} reason={issue.reason} "
                f"nc_rows={issue.nc_rows}"
            )
        return 0

    rows, issues, summary = workbook.build_local_plan(write_sheet=args.write)
    by_org = Counter(row.organization_short_name for row in rows)
    by_bank = Counter(row.bank for row in rows)
    print(f"receipt local plan rows: {len(rows)}")
    print(f"runnable rows: {summary['runnable_rows']}")
    print(f"issues: {len(issues)}")
    print(f"validation_policy: {summary['validation_policy']}")
    print(f"organizations: {dict(sorted(by_org.items()))}")
    print(f"banks: {dict(sorted(by_bank.items()))}")
    print(f"grouped row numbers: {summary['organizations']}")
    if args.write:
        print("machine result sheet written")
    for issue in issues[:50]:
        print(
            "issue "
            f"row={issue.excel_row} "
            f"stage={issue.stage} "
            f"type={issue.issue_type} "
            f"field={issue.field} "
            f"raw={issue.raw_value} "
            f"config={issue.config_node} "
            f"action={issue.action} "
            f"message={issue.message}"
        )
    if issues and summary["validation_policy"] == "strict":
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
