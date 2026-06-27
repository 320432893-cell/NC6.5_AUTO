# 职责：收款单批次结果汇总/控制台摘要/落盘报告/失败信息格式化
# 不做什么：不驱动 NC、不做业务判定
# 允许依赖层：标准库;谁不应 import：core 层

from decimal import Decimal
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.receipt_models import ReceiptBatchResultRow  # noqa: E402


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

SLOW_STEP_THRESHOLD_SECONDS = 0.5

def build_batch_results(selected_rows, row_reports):
    reports_by_row = {int(report.get("excel_row")): report for report in row_reports}
    results = []
    for row in selected_rows:
        report = reports_by_row.get(row.row) or {}
        ok = bool(report.get("ok"))
        reason = "" if ok else format_row_failure_reason(report)
        results.append(
            ReceiptBatchResultRow(
                plan_row=row,
                local_status="通过" if ok else "异常",
                exception_reason=reason,
                nc_customer_name=str(report.get("nc_customer_name") or "").strip(),
                nc_document_no=str(report.get("nc_document_no") or "").strip(),
            )
        )
    return results

def post_query_failure_reasons(post_query):
    if not post_query:
        return {"*": "后验查询失败"}
    issues = {}
    for group in post_query.get("groups") or []:
        match = group.get("match") or {}
        for row, reason in (match.get("issues") or {}).items():
            issues[str(row)] = reason or "后验未匹配"
        if not group.get("ok"):
            reason = group.get("reason") or "后验查询失败"
            for row in group.get("target_rows") or []:
                issues.setdefault(str(row), reason)
    if not post_query.get("ok") and not issues:
        return {"*": post_query.get("reason") or "后验查询失败"}
    return issues

def post_query_skip_reason(rows, exit_code):
    if not rows:
        return "没有完成任何收款单录入行，未执行后验查询"
    if exit_code != 0:
        return "录入/保存阶段未全部成功，未执行后验查询"
    return "后验查询条件未满足"

def format_row_failure_reason(report):
    failed_step = str(report.get("failed_step") or "").strip()
    reason = str(report.get("reason") or "").strip()
    if failed_step.startswith("save"):
        return f"保存失败-{reason or failed_step}"
    if failed_step:
        return (
            f"录入失败-{failed_step}:{reason}" if reason else f"录入失败-{failed_step}"
        )
    return reason or "录入失败"

def fail(row_report, failed_step, timings, reason):
    row_report.update(
        {
            "ok": False,
            "failed_step": failed_step,
            "reason": reason,
        }
    )
    attach_slow_step_summary(row_report, timings)
    return row_report

def attach_slow_step_summary(
    row_report,
    timings,
    threshold_seconds=SLOW_STEP_THRESHOLD_SECONDS,
):
    timing_items = list(getattr(timings, "items", []) or [])
    row_report["timings"] = timing_items
    row_report["slow_step_threshold_seconds"] = float(threshold_seconds)
    slow_steps = []

    def add_step(name, seconds, source, details=None):
        try:
            value = float(seconds)
        except (TypeError, ValueError):
            return
        if value < float(threshold_seconds):
            return
        item = {
            "name": name,
            "seconds": round(value, 3),
            "source": source,
        }
        if details:
            item["details"] = details
        slow_steps.append(item)

    for item in timing_items:
        add_step(item.get("name"), item.get("seconds"), "row")

    open_step = find_report_step(row_report, "open-self-made")
    parsed = (open_step or {}).get("parsed") or {}
    entry_context = parsed.get("entry_context_snapshot")
    if entry_context:
        row_report["open_self_made_entry_context"] = entry_context
    for item in parsed.get("timings") or []:
        add_step(
            f"open.self-made/{item.get('name')}",
            item.get("seconds"),
            "open.self-made",
        )

    slow_steps.sort(key=lambda item: item["seconds"], reverse=True)
    row_report["slow_steps"] = slow_steps[:30]

def find_report_step(row_report, name):
    for step in (row_report or {}).get("steps") or []:
        if step.get("name") == name:
            return step
    return None

def summarize_header_failure(header_steps):
    for step in header_steps or []:
        if step.get("ok"):
            continue
        label = step.get("label")
        reason = (
            step.get("reason")
            or step.get("stage")
            or ((step.get("scope") or {}).get("reason"))
            or ((step.get("path_attempt") or {}).get("reason"))
            or "表头字段写入失败"
        )
        return f"表头字段写入失败: {label or step.get('step') or '未知字段'} - {reason}"
    return "表头字段写入失败"

def serializable(value):
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, dict):
        return {key: serializable(item) for key, item in value.items()}
    if isinstance(value, list | tuple):
        return [serializable(item) for item in value]
    return value

def write_last_report(report):
    path = ROOT / "logs" / "last_receipt_full_flow_report.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(".json.tmp")
    with tmp_path.open("w", encoding="utf-8") as file:
        json.dump(report, file, ensure_ascii=False, indent=2, default=str)
        file.write("\n")
    tmp_path.replace(path)
    summary_path = ROOT / "logs" / "last_receipt_failure_summary.txt"
    summary_path.write_text(
        "\n".join(build_console_report_lines(report, path, summary_path)) + "\n",
        encoding="utf-8",
    )
    return path

def print_report(report, args):
    if args.json:
        text = json.dumps(report, ensure_ascii=False, indent=2, default=str)
        print(text)
        return
    report_path = ROOT / "logs" / "last_receipt_full_flow_report.json"
    summary_path = ROOT / "logs" / "last_receipt_failure_summary.txt"
    for line in build_console_report_lines(report, report_path, summary_path):
        print(line)

def build_console_report_lines(report, report_path=None, summary_path=None):
    lines = ["收款单完整流程结果摘要"]
    lines.append(f"结果：{'成功' if report.get('ok') else '失败'}")
    lines.append(f"总用时：{report.get('total_seconds')}s")
    rows = report.get("rows") or []
    failed_row = next((row for row in rows if not row.get("ok")), None)
    if failed_row:
        lines.append(f"失败行：Sheet 行 {failed_row.get('excel_row')}")
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        if ok_rows:
            lines.append(f"已保存行：{ok_rows}（这些已在 NC 落库，需人工核对/处理）")
        lines.append(f"失败阶段：{failed_row.get('failed_step') or '未知'}")
        if failed_row.get("reason"):
            lines.append(f"失败原因：{failed_row.get('reason')}")
        entry_scope_hwnd = failed_row.get("entry_scope_hwnd")
        entry_dynamic_index = failed_row.get("entry_dynamic_index")
        if entry_scope_hwnd is not None or entry_dynamic_index is not None:
            lines.append(
                "入口上下文："
                f"scope_hwnd={entry_scope_hwnd}, "
                f"entry_dynamic_index={entry_dynamic_index}"
            )
        header_step = first_failed_step(failed_row.get("header_steps"))
        if header_step:
            lines.extend(format_header_failure_lines(header_step))
        modal_events = ((failed_row.get("modal_recovery") or {}).get("events")) or []
        if modal_events:
            last_modal = modal_events[-1]
            lines.append(
                "弹窗恢复："
                f"attempted={last_modal.get('attempted')}, "
                f"ok={last_modal.get('ok')}, "
                f"reason={last_modal.get('reason') or ''}"
            )
        else:
            lines.append("弹窗恢复：本次失败点没有检测到可取消弹窗")
        timings = failed_row.get("timings") or []
        if timings:
            lines.append("关键耗时：" + format_timings(timings))
    elif report.get("post_query_failed_rows"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"录入保存通过行：{ok_rows}")
        lines.append("失败阶段：post-query")
        for row, reason in (report.get("post_query_failed_rows") or {}).items():
            lines.append(f"后验未匹配行 {row}：{reason}")
    elif report.get("ok"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"通过行：{ok_rows}")
        post_query = report.get("post_query") or {}
        if post_query:
            matched = 0
            issues = 0
            for group in post_query.get("groups") or []:
                match = group.get("match") or {}
                matched += len(match.get("matched") or {})
                issues += len(match.get("issues") or {})
            lines.append(f"后验查询：已执行，匹配 {matched} 行，未匹配 {issues} 行")
    elif report.get("reason"):
        lines.append(f"失败原因：{report.get('reason')}")
    if report_path:
        lines.append(f"完整报告：{report_path}")
    if summary_path:
        lines.append(f"摘要文件：{summary_path}")
    return lines

def first_failed_step(steps):
    for step in steps or []:
        if step.get("ok"):
            continue
        return step
    return None

def format_header_failure_lines(step):
    lines = []
    label = step.get("label") or step.get("step") or "未知字段"
    lines.append(f"表头失败字段：{label}")
    if step.get("stage"):
        lines.append(f"表头失败阶段：{step.get('stage')}")
    scope = step.get("scope") or {}
    if scope:
        lines.append(
            "表头 scope："
            f"mode={scope.get('mode')}, "
            f"dynamic_index={scope.get('dynamic_index')}, "
            f"dynamic_prefix={scope.get('dynamic_prefix')}"
        )
    path_attempt = step.get("path_attempt") or {}
    if path_attempt:
        lines.append(
            "path 尝试："
            f"{path_attempt.get('path') or ''} "
            f"({path_attempt.get('reason') or '无原因'})"
        )
    modal_recovery = step.get("modal_recovery") or {}
    if modal_recovery:
        lines.append(
            "字段级弹窗恢复："
            f"attempted={modal_recovery.get('attempted')}, "
            f"ok={modal_recovery.get('ok')}, "
            f"reason={modal_recovery.get('reason') or ''}"
        )
    return lines

def format_timings(timings):
    chunks = []
    for item in timings:
        name = item.get("name")
        seconds = item.get("seconds")
        if name is None or seconds is None:
            continue
        chunks.append(f"{name}={seconds}s")
    return ", ".join(chunks)

def user_excel_locked_message(exc):
    return (
        "Excel 文件无法写入。请先关闭正在打开的 Excel/WPS 文件、关闭资源管理器预览窗格，"
        "或取消“写入选中计划 Sheet2”后重试；原始错误："
        f"{exc}"
    )
