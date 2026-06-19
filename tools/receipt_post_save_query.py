# 职责：对本次已保存的收款单批次做后验查询、匹配 NC 单据号并生成 Sheet2 结果
# 不做什么：不录入收款单、不写 Sheet1 状态列、不解析 CLI、不保存 NC 单据
# 允许依赖层：core 收款模型/匹配纯函数、tools 查询填充和分页读表组件
# 谁不应该 import：core 层模块不应 import；查询历史写回入口不应反向 import

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
import time

from core.jab_operator import JABOperator
from core.receipt_amounts import receipt_nc_amount
from core.receipt_models import ReceiptBatchResultRow, ReceiptPlanRow
from core.receipt_nc_extract import ReceiptNCResultExtractor
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


class PostSaveQueryError(Exception):
    """后验查询失败：message 是给人看的业务原因；detail 仅作开发诊断，不进 Sheet2。"""

    def __init__(self, message, detail=""):
        super().__init__(message)
        self.detail = detail


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
            raise PostSaveQueryError(
                f"后验查询失败：主体 {org_code} 的查询条件窗口未能打开"
                "（期望按 F3 弹出【查询条件】窗口，实际未出现）；"
                "请确认 NC 已停在收款单录入页且前台未被其它窗口占用后重试。"
            )
        query_scope = timings.measure(
            "query.dynamic-scope",
            find_query_condition_scope,
            jab,
            jab_cfg,
        )
        if not query_scope.get("ok"):
            raise PostSaveQueryError(
                f"后验查询失败：主体 {org_code} 的查询条件窗口已打开，"
                "但未能定位收款财务组织/单据日期输入区；请人工在 NC 中手动查询核对。",
                detail=f"query_scope:{query_scope.get('reason')}",
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
            raise PostSaveQueryError(
                f"后验查询失败：收款财务组织条件写入失败，期望写入主体编码 {org_code}，"
                "NC 未接受该值；请人工在查询条件中确认主体后重试。",
                detail=f"set_finance_org:{org_ok.get('reason')}",
            )
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
                field_label = (
                    "单据日期起" if name == "document_date_from" else "单据日期止"
                )
                raise PostSaveQueryError(
                    f"后验查询失败：{field_label}条件写入失败，期望写入 {value}，"
                    "NC 未接受该值；请人工在查询条件中确认日期区间后重试。",
                    detail=f"{name}={value} reason={written.get('reason')}",
                )
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
            raise PostSaveQueryError(
                f"后验查询失败：主体 {org_code} 的查询条件已填好，但点击【确定】未生效；"
                "请人工在 NC 中手动确认查询后核对结果。",
            )
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
        match = match_snapshot_to_result(
            targets,
            match_snapshot,
        )
        group_report.update(
            {
                "ok": True,
                "page_report": page_report,
                "match": match,
            }
        )
    except PostSaveQueryError as exc:
        user_reason = str(exc)
        group_report.update(
            {
                "ok": False,
                "reason": user_reason,
                "error_detail": exc.detail,
                "match": {
                    "matched": {},
                    "issues": {target.row.row: user_reason for target in targets},
                },
            }
        )
    except Exception as exc:
        # 给人看的原因保持业务可读；异常类型/原文只留作开发诊断字段，不进 Sheet2。
        user_reason = (
            f"后验查询失败：主体 {org_code} 在查询/读结果阶段中断；"
            "请人工在 NC 中手动查询该主体本批日期区间后核对单据号。"
        )
        group_report.update(
            {
                "ok": False,
                "reason": user_reason,
                "error_detail": f"{type(exc).__name__}: {exc}",
                "match": {
                    "matched": {},
                    "issues": {target.row.row: user_reason for target in targets},
                },
            }
        )
    return group_report


def match_snapshot_to_result(targets, match_snapshot):
    snapshot = match_snapshot or {}
    matched = snapshot.get("matched") or {}
    match_issues = snapshot.get("match_issues") or []
    result_issues = {
        int(issue.excel_row): str(issue.reason or "后验未匹配")
        for issue in match_issues
    }
    for target in targets:
        if target.row.row not in matched:
            result_issues.setdefault(target.row.row, "后验未匹配")
    return {
        "matched": {
            int(row): getattr(nc_row, "document_no", "")
            for row, nc_row in matched.items()
        },
        "issues": result_issues,
    }


def target_to_match_row(target):
    return BatchMatchRow(
        row=target.row.row,
        receipt_date=target.row.receipt_date,
        raw_amount=receipt_nc_amount(target.row),
        payer_name=str(target.row_report.get("nc_customer_name") or "").strip(),
    )


@dataclass(frozen=True)
class BatchMatchRow:
    row: int
    receipt_date: object
    raw_amount: Decimal
    payer_name: str


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
