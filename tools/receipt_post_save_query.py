# 职责：对本次已保存的收款单批次做后验查询、匹配 NC 单据号并生成 Sheet2 结果
# 不做什么：不录入收款单、不写 Sheet1 状态列、不解析 CLI、不保存 NC 单据
# 允许依赖层：core 收款模型/匹配纯函数、tools 查询填充和分页读表组件
# 谁不应该 import：core 层模块不应 import；查询历史写回入口不应反向 import

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from difflib import SequenceMatcher
import time

from core.jab_operator import JABOperator
from core.receipt_models import ReceiptBatchResultRow, ReceiptPlanRow
from core.receipt_nc_extract import ReceiptNCResultExtractor
from core.receipt_parsing import normalize_lookup_key
from tools.receipt_query_fill import (
    ensure_query_window,
)
from tools.receipt_query_dynamic_fields import (
    find_query_condition_scope,
    set_query_dynamic_text,
)
from tools.receipt_query_guard import (
    guard_receipt_parent_page,
    guard_receipt_result_tables,
)
from tools.receipt_query_match_reader import read_receipt_result_pages_incremental
from tools.receipt_query_pagination import wait_after_query_confirm
from tools.receipt_query_result_tables import receipt_result_read_columns


@dataclass(frozen=True)
class BatchQueryTarget:
    row: ReceiptPlanRow
    row_report: dict


class TimingRecorder:
    def __init__(self):
        self.items = []

    def measure(self, name, func, *args, **kwargs):
        started = time.perf_counter()
        result = func(*args, **kwargs)
        self.items.append(
            {"name": name, "seconds": round(time.perf_counter() - started, 3)}
        )
        return result


def run_post_save_batch_query(config, selected_rows, row_reports):
    reports_by_row = {int(report.get("excel_row")): report for report in row_reports}
    targets = [
        BatchQueryTarget(row=row, row_report=reports_by_row.get(row.row) or {})
        for row in selected_rows
        if (reports_by_row.get(row.row) or {}).get("ok")
    ]
    results_by_row = {
        row.row: ReceiptBatchResultRow(
            plan_row=row,
            local_status=(
                "通过" if (reports_by_row.get(row.row) or {}).get("ok") else "异常"
            ),
            exception_reason=(
                ""
                if (reports_by_row.get(row.row) or {}).get("ok")
                else format_entry_failure(reports_by_row.get(row.row) or {})
            ),
            nc_customer_name=str(
                (reports_by_row.get(row.row) or {}).get("nc_customer_name") or ""
            ).strip(),
        )
        for row in selected_rows
    }
    report = {"ok": True, "groups": [], "timings": []}
    if not targets:
        report["ok"] = False
        report["reason"] = "没有保存成功的行可后验查询"
        return list(results_by_row.values()), report

    grouped = group_targets_by_org(targets)
    if len(grouped) > 4:
        report["ok"] = False
        report["reason"] = f"查询失败-本批主体数量超过4个:{sorted(grouped)}"
        for target in targets:
            current = results_by_row[target.row.row]
            results_by_row[target.row.row] = ReceiptBatchResultRow(
                plan_row=current.plan_row,
                local_status=current.local_status,
                exception_reason=report["reason"],
                nc_customer_name=current.nc_customer_name,
                nc_document_no="",
            )
        return [results_by_row[row.row] for row in selected_rows], report

    query_cfg = config["receipt_entry"]["query"]
    jab_cfg = query_cfg["jab"]
    extractor = ReceiptNCResultExtractor(config)
    timings = TimingRecorder()
    jab = JABOperator(config)
    try:
        timings.measure("jab.ensure-started", jab.ensure_started)
        for org_code, group_targets in grouped.items():
            group_report = query_one_org(
                jab,
                config,
                query_cfg,
                jab_cfg,
                extractor,
                org_code,
                group_targets,
                timings,
            )
            report["groups"].append(group_report)
            apply_group_match_results(results_by_row, group_targets, group_report)
            if not group_report.get("ok"):
                report["ok"] = False
    finally:
        jab.close()
    report["timings"] = timings.items
    return [results_by_row[row.row] for row in selected_rows], report


def group_targets_by_org(targets):
    grouped = defaultdict(list)
    for target in sorted(
        targets,
        key=lambda item: (
            item.row.organization_code,
            item.row.receipt_date,
            item.row.row,
        ),
    ):
        grouped[target.row.organization_code].append(target)
    return dict(grouped)


def query_one_org(
    jab,
    config,
    query_cfg,
    jab_cfg,
    extractor,
    org_code,
    targets,
    timings,
):
    date_from = min(target.row.receipt_date for target in targets).isoformat()
    date_to = max(target.row.receipt_date for target in targets).isoformat()
    group_report = {
        "organization_code": org_code,
        "date_from": date_from,
        "date_to": date_to,
        "target_rows": [target.row.row for target in targets],
    }
    try:
        timings.measure("page_guard", guard_receipt_parent_page, jab, config, query_cfg)
        opened = timings.measure(
            "ensure_query_window",
            ensure_query_window,
            jab,
            config,
            query_cfg,
            jab_cfg,
            False,
        )
        if not opened:
            raise RuntimeError("查询窗口未打开")
        query_scope = timings.measure(
            "query.dynamic-scope",
            find_query_condition_scope,
            jab,
            jab_cfg,
        )
        if not query_scope.get("ok"):
            raise RuntimeError(
                f"查询条件动态 path 定位失败:{query_scope.get('reason')}"
            )
        org_ok = timings.measure(
            "set_finance_org",
            set_query_dynamic_text,
            jab,
            jab_cfg,
            query_scope,
            "finance_org",
            org_code,
        )
        if not org_ok.get("ok"):
            raise RuntimeError(f"主体写入失败:{org_code}")
        for name, value in [
            ("document_date_from", date_from),
            ("document_date_to", date_to),
        ]:
            written = timings.measure(
                name,
                set_query_dynamic_text,
                jab,
                jab_cfg,
                query_scope,
                name,
                value,
            )
            if not written.get("ok"):
                raise RuntimeError(f"日期条件写入失败:{name}={value}")
        confirmed = timings.measure(
            "confirm_query",
            jab.do_action_by_path,
            jab_cfg["confirm_button_path"],
            title=jab_cfg["dialog_title"],
            class_name=jab_cfg["dialog_class"],
            role="push button",
            action_name="单击",
            wait=float(jab_cfg.get("confirm_wait", 0.0)),
            timeout=float(jab_cfg.get("confirm_timeout", 1.0)),
            require_showing=True,
        )
        if not confirmed:
            raise RuntimeError("确定查询失败")
        wait_after_query_confirm(jab, query_cfg)
        tables, page_report, match_snapshot = timings.measure(
            "read_receipt_result_pages_incremental",
            read_receipt_result_pages_incremental,
            jab,
            query_cfg,
            extractor,
            [target_to_match_row(target) for target in targets],
            max_rows=500,
            max_cols=80,
            read_columns=receipt_result_read_columns(
                query_cfg,
                include_amount_candidates=False,
            ),
        )
        timings.measure(
            "guard_receipt_result_tables",
            guard_receipt_result_tables,
            tables,
            query_cfg,
        )
        match = match_targets(targets, (match_snapshot or {}).get("nc_rows") or [])
        group_report.update(
            {
                "ok": True,
                "page_report": page_report,
                "match": match,
            }
        )
    except Exception as exc:
        group_report.update(
            {
                "ok": False,
                "reason": f"查询失败-{type(exc).__name__}:{exc}",
                "match": {
                    "matched": {},
                    "issues": {target.row.row: f"查询失败-{exc}" for target in targets},
                },
            }
        )
    return group_report


def target_to_match_row(target):
    return BatchMatchRow(
        row=target.row.row,
        receipt_date=target.row.receipt_date,
        raw_amount=target.row.raw_amount,
        payer_name=str(target.row_report.get("nc_customer_name") or "").strip(),
    )


@dataclass(frozen=True)
class BatchMatchRow:
    row: int
    receipt_date: object
    raw_amount: Decimal
    payer_name: str


def match_targets(targets, nc_rows):
    matched = {}
    issues = {}
    for target in targets:
        expected_name = str(target.row_report.get("nc_customer_name") or "").strip()
        amount_candidates = [
            nc_row
            for nc_row in nc_rows
            if nc_row.original_amount == target.row.raw_amount
        ]
        exact_date = [
            nc_row
            for nc_row in amount_candidates
            if nc_row.document_date == target.row.receipt_date
        ] or amount_candidates
        scored = [
            (name_similarity(expected_name, nc_row.name), nc_row)
            for nc_row in exact_date
            if expected_name
        ]
        candidates = [(score, nc_row) for score, nc_row in scored if score >= 90]
        if len(candidates) == 1:
            matched[target.row.row] = candidates[0][1]
        elif len(candidates) > 1:
            issues[target.row.row] = "后验重复匹配-金额和名称命中多条"
        elif exact_date:
            best = max(scored, default=(0, None), key=lambda item: item[0])
            if best[1] is not None:
                issues[target.row.row] = (
                    f"后验未匹配-金额相同但名称相似度{best[0]:.0f}低于90:"
                    f"{expected_name} vs {best[1].name}"
                )
            else:
                issues[target.row.row] = (
                    f"后验未匹配-名称{expected_name}无对应，但有相同金额"
                )
        else:
            name_candidates = [
                nc_row
                for nc_row in nc_rows
                if name_similarity(expected_name, nc_row.name) >= 90
            ]
            if name_candidates:
                amounts = ",".join(
                    str(row.original_amount) for row in name_candidates[:3]
                )
                issues[target.row.row] = (
                    f"后验未匹配-名称匹配但金额不一致:NC金额={amounts}"
                )
            else:
                issues[target.row.row] = f"后验未匹配-金额{target.row.raw_amount}无对应"
    return {
        "matched": {row: nc_row.document_no for row, nc_row in matched.items()},
        "issues": issues,
    }


def name_similarity(left, right):
    left_key = normalize_lookup_key(left)
    right_key = normalize_lookup_key(right)
    if not left_key or not right_key:
        return 0.0
    if left_key == right_key:
        return 100.0
    return SequenceMatcher(None, left_key, right_key).ratio() * 100


def apply_group_match_results(results_by_row, targets, group_report):
    match = group_report.get("match") or {}
    matched = match.get("matched") or {}
    issues = match.get("issues") or {}
    for target in targets:
        current = results_by_row[target.row.row]
        if target.row.row in matched:
            results_by_row[target.row.row] = ReceiptBatchResultRow(
                plan_row=current.plan_row,
                local_status=current.local_status,
                exception_reason="",
                nc_customer_name=current.nc_customer_name,
                nc_document_no=str(matched[target.row.row] or ""),
            )
        else:
            results_by_row[target.row.row] = ReceiptBatchResultRow(
                plan_row=current.plan_row,
                local_status=current.local_status,
                exception_reason=(
                    issues.get(target.row.row)
                    or group_report.get("reason")
                    or "后验未匹配"
                ),
                nc_customer_name=current.nc_customer_name,
                nc_document_no="",
            )


def format_entry_failure(report):
    failed_step = str(report.get("failed_step") or "").strip()
    reason = str(report.get("reason") or "").strip()
    if failed_step.startswith("save"):
        return f"保存失败-{reason or failed_step}"
    if failed_step:
        return (
            f"录入失败-{failed_step}:{reason}" if reason else f"录入失败-{failed_step}"
        )
    return reason or "录入失败"
