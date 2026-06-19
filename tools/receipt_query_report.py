# 职责：生成收款单查询 dry-run 匹配报告、写回计划报告和匹配输入诊断。
# 不做什么：不操作 NC/JAB 窗口，不读取分页，不解析 CLI 参数。
# 允许依赖层：core.receipt_entry、core.receipt_matching、core.receipt_nc_extract 暴露的模型和匹配能力。
# 谁不应该 import：JAB 底层适配器、临时探针脚本、与收款单查询无关的业务流程。

from collections import Counter

from core.receipt_entry import ReceiptEntryWorkbook
from core.receipt_matching import (
    RECEIPT_NOT_FOUND_MARKER,
    ReceiptEntryDryRunMatcher,
    format_receipt_amount_name_mismatch_reason,
    format_receipt_duplicate_reason,
    format_receipt_name_amount_mismatch_reason,
    names_match,
)
from tools.receipt_query_result_tables import (
    RECEIPT_RESULT_AMOUNT_CANDIDATE_COLUMNS,
    RECEIPT_RESULT_NAME_CANDIDATE_COLUMNS,
)


def build_dry_run_match_report(
    config,
    extractor,
    tables,
    org_code,
    business_date,
    write_back=False,
):
    rows, candidates, excel_issues = ReceiptEntryWorkbook(config).preview_rows(
        today=business_date
    )
    return build_dry_run_match_report_from_preview(
        config,
        extractor,
        tables,
        org_code,
        business_date,
        rows,
        candidates,
        excel_issues,
        write_back=write_back,
    )


def build_dry_run_match_report_from_preview(
    config,
    extractor,
    tables,
    org_code,
    business_date,
    rows,
    candidates,
    excel_issues,
    write_back=False,
    target_rows=None,
    configured_match_snapshot=None,
):
    org_candidates = [row for row in candidates if row.organization_code == org_code]
    match_candidates = target_rows or org_candidates
    report = {
        "business_date": business_date.isoformat(),
        "excel_rows": len(rows),
        "excel_candidates": len(candidates),
        "org_candidates": len(org_candidates),
        "match_candidates": len(match_candidates),
        "excel_issues": len(excel_issues),
        "candidate_banks": dict(
            sorted(Counter(row.bank for row in match_candidates).items())
        ),
        "write_back": {"enabled": bool(write_back), "updated": 0, "rows": []},
        "variants": [],
    }
    matcher = ReceiptEntryDryRunMatcher()
    configured_amount_column = extractor.config.result_column_indexes["original_amount"]
    configured_name_column = extractor.config.result_column_indexes["payer_name"]
    dry_run_all_variants = bool(
        (config.get("receipt_entry") or {})
        .get("query", {})
        .get("dry_run_all_variants", False)
    )
    if dry_run_all_variants:
        amount_columns = unique_ordered(
            [configured_amount_column, *RECEIPT_RESULT_AMOUNT_CANDIDATE_COLUMNS]
        )
        name_columns = unique_ordered(
            [configured_name_column, *RECEIPT_RESULT_NAME_CANDIDATE_COLUMNS]
        )
    else:
        amount_columns = [configured_amount_column]
        name_columns = [configured_name_column]
    for column in name_columns:
        for amount_column in amount_columns:
            variant_name = f"name_col{column}_amount_col{amount_column}"
            is_configured_variant = (
                column == configured_name_column
                and amount_column == configured_amount_column
            )
            if is_configured_variant and configured_match_snapshot:
                nc_rows = configured_match_snapshot["nc_rows"]
                extract_issues = configured_match_snapshot["extract_issues"]
                matched = configured_match_snapshot["matched"]
                match_issues = configured_match_snapshot["match_issues"]
                source = "incremental"
            else:
                nc_rows, extract_issues = extractor.extract_by_indexes(
                    tables,
                    column,
                    amount_column=amount_column,
                )
                matched, match_issues = matcher.match(match_candidates, nc_rows)
                source = "computed"
            report["variants"].append(
                {
                    "name": variant_name,
                    "name_column": column,
                    "amount_column": amount_column,
                    "source": source,
                    "nc_rows": len(nc_rows),
                    "nc_summary": summarize_nc_rows(nc_rows),
                    "match_diagnostics": diagnose_match_inputs(
                        match_candidates, nc_rows
                    ),
                    "extract_issues": len(extract_issues),
                    "matches": len(matched),
                    "match_issues": len(match_issues),
                    "matched_excel_rows": sorted(matched.keys())[:20],
                    "issue_samples": [
                        {
                            "excel_row": issue.excel_row,
                            "reason": issue.reason,
                            "nc_rows": issue.nc_rows,
                        }
                        for issue in match_issues[:20]
                    ],
                    "extract_issue_samples": [
                        {
                            "table_index": issue.table_index,
                            "row_index": issue.row_index,
                            "reason": issue.reason,
                        }
                        for issue in extract_issues[:20]
                    ],
                }
            )
            if is_configured_variant:
                report["write_back"] = build_receipt_write_back_report(
                    config,
                    match_candidates,
                    matched,
                    match_issues,
                    enabled=write_back,
                )
    return report


def unique_ordered(values):
    result = []
    seen = set()
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        result.append(value)
    return result


def build_receipt_write_back_report(
    config,
    excel_rows,
    matched,
    match_issues,
    enabled=False,
):
    issue_by_row = {issue.excel_row: issue for issue in match_issues}
    statuses = {}
    duplicate_rows = []
    exception_rows = []
    for excel_row in excel_rows:
        if excel_row.row in matched:
            statuses[excel_row.row] = "已做过"
            continue
        issue = issue_by_row.get(excel_row.row)
        if issue and issue.reason.startswith(RECEIPT_NOT_FOUND_MARKER):
            statuses[excel_row.row] = "未做过"
        elif issue:
            statuses[excel_row.row] = issue.reason
            exception_rows.append(excel_row.row)
            if issue.reason.startswith("重复"):
                duplicate_rows.append(excel_row.row)

    report = {
        "enabled": bool(enabled),
        "planned": len(statuses),
        "matched_rows": sorted(matched),
        "not_found_rows": sorted(
            row for row, status in statuses.items() if status == "未做过"
        ),
        "duplicate_rows": sorted(duplicate_rows),
        "exception_rows": sorted(exception_rows),
        "skipped_duplicate_rows": sorted(duplicate_rows),
        "updated": 0,
        "rows": [],
    }
    if enabled:
        write_result = ReceiptEntryWorkbook(config).write_nc_done_statuses(statuses)
        report["updated"] = write_result["updated"]
        report["rows"] = write_result["rows"]
    return report


def diagnose_match_inputs(excel_rows, nc_rows):
    nc_by_amount = {}
    for nc_row in nc_rows:
        nc_by_amount.setdefault(nc_row.original_amount, []).append(nc_row)

    amount_only_hits = 0
    name_amount_hits = 0
    duplicate_hits = 0
    name_only_hits = 0
    name_amount_samples = []
    name_mismatch_samples = []
    amount_mismatch_samples = []
    no_amount_samples = []
    for excel_row in excel_rows:
        amount_candidates = nc_by_amount.get(excel_row.raw_amount, [])
        if not amount_candidates:
            name_candidates = [
                nc_row
                for nc_row in nc_rows
                if names_match(excel_row.payer_name, nc_row.name)
            ]
            if name_candidates:
                name_only_hits += 1
                if len(amount_mismatch_samples) < 10:
                    amount_mismatch_samples.append(
                        {
                            "excel_row": excel_row.row,
                            "excel_amount": str(excel_row.raw_amount),
                            "excel_name": excel_row.payer_name,
                            "reason": format_receipt_name_amount_mismatch_reason(
                                excel_amount=excel_row.raw_amount,
                                excel_name=excel_row.payer_name,
                                nc_amounts=[
                                    row.original_amount for row in name_candidates
                                ],
                            ),
                            "nc_amounts": [
                                str(row.original_amount) for row in name_candidates[:5]
                            ],
                            "nc_rows": [row.row_index for row in name_candidates[:5]],
                        }
                    )
                continue
            if len(no_amount_samples) < 10:
                no_amount_samples.append(
                    {
                        "excel_row": excel_row.row,
                        "excel_amount": str(excel_row.raw_amount),
                        "excel_name": excel_row.payer_name,
                    }
                )
            continue

        amount_only_hits += 1
        matched_names = [
            nc_row
            for nc_row in amount_candidates
            if names_match(excel_row.payer_name, nc_row.name)
        ]
        if len(matched_names) == 1:
            name_amount_hits += 1
            if len(name_amount_samples) < 10:
                name_amount_samples.append(
                    {
                        "excel_row": excel_row.row,
                        "excel_amount": str(excel_row.raw_amount),
                        "excel_name": excel_row.payer_name,
                        "nc_name": matched_names[0].name,
                        "nc_row": matched_names[0].row_index,
                    }
                )
        elif len(matched_names) > 1:
            duplicate_hits += 1
            if len(name_mismatch_samples) < 10:
                name_mismatch_samples.append(
                    {
                        "excel_row": excel_row.row,
                        "excel_amount": str(excel_row.raw_amount),
                        "excel_name": excel_row.payer_name,
                        "reason": format_receipt_duplicate_reason(len(matched_names)),
                        "nc_names": [row.name for row in matched_names[:5]],
                        "nc_rows": [row.row_index for row in matched_names[:5]],
                    }
                )
        elif len(name_mismatch_samples) < 10:
            name_mismatch_samples.append(
                {
                    "excel_row": excel_row.row,
                    "excel_amount": str(excel_row.raw_amount),
                    "excel_name": excel_row.payer_name,
                    "reason": format_receipt_amount_name_mismatch_reason(
                        excel_amount=excel_row.raw_amount,
                        excel_name=excel_row.payer_name,
                        nc_names=[row.name for row in amount_candidates],
                    ),
                    "nc_names": [row.name for row in amount_candidates[:5]],
                    "nc_rows": [row.row_index for row in amount_candidates[:5]],
                }
            )
    return {
        "amount_only_hits": amount_only_hits,
        "name_amount_hits": name_amount_hits,
        "duplicate_hits": duplicate_hits,
        "name_only_hits": name_only_hits,
        "name_amount_samples": name_amount_samples,
        "name_mismatch_samples": name_mismatch_samples,
        "amount_mismatch_samples": amount_mismatch_samples,
        "no_amount_samples": no_amount_samples,
    }


def summarize_nc_rows(nc_rows):
    if not nc_rows:
        return {
            "amount_min": None,
            "amount_max": None,
            "name_samples": [],
        }
    amounts = [row.original_amount for row in nc_rows]
    names = []
    seen = set()
    for row in nc_rows:
        name = row.name.strip()
        if not name or name in seen:
            continue
        seen.add(name)
        names.append(name)
        if len(names) >= 10:
            break
    return {
        "amount_min": str(min(amounts)),
        "amount_max": str(max(amounts)),
        "name_samples": names,
    }
