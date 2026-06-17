# 职责：收款单完整流程的报告序列化、落盘与控制台摘要渲染
# 不做什么：不驱动 NC/JAB，不做行编排，不读取 Excel，不触发保存
# 允许依赖层：标准库 json/decimal/pathlib，仓库内 ROOT 路径
# 谁不应该 import：core 层模块不应 import；本模块不应反向 import row_runner/entry

import json
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


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
    elif report.get("ok"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"通过行：{ok_rows}")
    elif report.get("post_query_failed_rows"):
        ok_rows = [row.get("excel_row") for row in rows if row.get("ok")]
        lines.append(f"录入保存通过行：{ok_rows}")
        lines.append("失败阶段：post-query")
        for row, reason in (report.get("post_query_failed_rows") or {}).items():
            lines.append(f"后验未匹配行 {row}：{reason}")
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
