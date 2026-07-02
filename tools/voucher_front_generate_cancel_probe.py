# lifecycle: tool（现场诊断脚本）
# 作用：测试“待生成表批量选中 -> 前台生成 -> 制单窗口打开 -> 取消后重试”链路。
# 风险：会真实点击 NC 的“生成 -> 前台生成”，但不会保存凭证；默认要求人工关闭/取消制单窗口。
# 运行：Windows Python 下在 nc_auto_v2 根目录执行：
#   python tools/voucher_front_generate_cancel_probe.py --config config.json --counts 200,250

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import jab_table_ops  # noqa: E402
from core.jab_operator import JABOperator  # noqa: E402
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="NC 凭证前台生成取消重试探测：只打开制单窗口，不保存。"
    )
    parser.add_argument("--config", default="config.json", help="配置文件路径")
    parser.add_argument("--start-row", type=int, default=0, help="从 NC 待生成表第几行开始选，0 基")
    parser.add_argument("--counts", default="200,250", help="先测数量，逗号分隔")
    parser.add_argument("--step-start", type=int, default=260, help="阶梯测试起始数量")
    parser.add_argument("--step", type=int, default=10, help="阶梯递增数量")
    parser.add_argument("--max-count", type=int, default=300, help="最多测到多少行")
    parser.add_argument("--selection-col", type=int, default=None, help="覆盖选中列，默认用 config 的 jab.selection_col")
    parser.add_argument("--main-table-timeout", type=float, default=5.0, help="等待待生成主表秒数")
    parser.add_argument("--voucher-timeout", type=float, default=15.0, help="等待制单窗口表格秒数")
    parser.add_argument("--cancel-timeout", type=float, default=60.0, help="等待人工取消/关闭制单窗口秒数")
    parser.add_argument("--continue-after-failure", action="store_true", help="失败后继续测后续数量")
    parser.add_argument("--auto-close-voucher", action="store_true", help="自动关闭制单窗口；默认人工取消更安全")
    parser.add_argument("--yes-front-generate", action="store_true", help="确认允许脚本点击前台生成")
    parser.add_argument("--json", action="store_true", help="stdout 输出完整 JSON")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if not args.yes_front_generate:
        print("本探针会真实点击 NC 的“生成 -> 前台生成”，但不会保存。")
        print("请确认当前停在 NC 待生成页，且可以安全打开后取消制单窗口。")
        print("如确认执行，请加参数：--yes-front-generate")
        return 2

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    started = time.perf_counter()
    try:
        jab.ensure_started()
        report = run_probe(jab, cfg, args, started)
    finally:
        jab.close()

    report_path = write_report(report, "voucher_front_generate_cancel_probe")
    report["report_path"] = str(report_path)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_summary(report)
    return 0 if report.get("ok") else 1


def run_probe(jab: JABOperator, cfg: dict, args, started: float) -> dict:
    batch_cfg = cfg.get("jab_batch", {})
    voucher_title = batch_cfg.get("voucher_window_title", "制单")
    voucher_class = batch_cfg.get("voucher_window_class", "SunAwtDialog")
    counts = build_counts(args)
    results = []
    last_ok = None
    first_failed = None

    for count in counts:
        result = run_one_count(
            jab,
            args,
            count=count,
            voucher_title=voucher_title,
            voucher_class=voucher_class,
        )
        results.append(result)
        if result.get("ok"):
            last_ok = count
        elif first_failed is None:
            first_failed = count
            if not args.continue_after_failure:
                break

    return {
        "generated_at": now_text(),
        "ok": True,
        "reason": "probe_finished",
        "elapsed_s": round(time.perf_counter() - started, 3),
        "plan": {
            "counts": counts,
            "start_row": args.start_row,
            "step": args.step,
            "max_count": args.max_count,
            "manual_cancel": not bool(args.auto_close_voucher),
        },
        "last_ok_count": last_ok,
        "first_failed_count": first_failed,
        "suspected_front_generate_limit": last_ok,
        "results": results,
    }


def run_one_count(
    jab: JABOperator,
    args,
    *,
    count: int,
    voucher_title: str,
    voucher_class: str,
) -> dict:
    item = {
        "count": count,
        "started_at": now_text(),
        "ok": False,
        "reason": "",
    }
    started = time.perf_counter()

    table = get_main_table(jab, args.main_table_timeout)
    if not table.get("ok"):
        return {**item, **table, "elapsed_ms": elapsed_ms(started)}

    vm_id = table["vm_id"]
    table_context = table["context"]
    owned_contexts = table["owned_contexts"]
    table_info = table["table_info"]
    try:
        row_count = int(table_info.rowCount)
        col_count = int(table_info.columnCount)
        selection_col = (
            int(args.selection_col)
            if args.selection_col is not None
            else int(jab.resolve_selection_col(None))
        )
        item["pending_table"] = {
            "row_count": row_count,
            "col_count": col_count,
            "selection_col": selection_col,
            "start_row": args.start_row,
        }
        if args.start_row + count > row_count:
            item.update(
                {
                    "reason": "not_enough_pending_rows",
                    "row_count": row_count,
                    "elapsed_ms": elapsed_ms(started),
                }
            )
            return item

        selected = select_rows_from_context(
            jab,
            vm_id,
            table_context,
            row_count=row_count,
            col_count=col_count,
            selection_col=selection_col,
            start_row=args.start_row,
            count=count,
        )
        item["selection"] = selected
        if not selected.get("ok"):
            item.update({"reason": "selection_failed", "elapsed_ms": elapsed_ms(started)})
            return item
    finally:
        jab.release_contexts(vm_id, owned_contexts)

    clicked_at = time.perf_counter()
    if not jab.do_generate_front():
        item.update({"reason": "front_generate_click_failed", "elapsed_ms": elapsed_ms(started)})
        return item
    item["front_generate_click_ms"] = elapsed_ms(clicked_at)

    voucher = wait_voucher_tables(jab, voucher_title, count, timeout=args.voucher_timeout)
    item["voucher_window"] = voucher
    if not voucher.get("ok"):
        item.update({"reason": "voucher_window_not_ready", "elapsed_ms": elapsed_ms(started)})
        cleanup_after_generate(jab, voucher_title, voucher_class, args)
        return item

    cancel = cleanup_after_generate(jab, voucher_title, voucher_class, args)
    item["cancel"] = cancel
    if not cancel.get("ok"):
        item.update({"reason": "voucher_window_not_closed", "elapsed_ms": elapsed_ms(started)})
        return item

    ready = wait_main_table_ready(jab, timeout=args.main_table_timeout)
    item["pending_ready_after_cancel"] = ready
    item.update(
        {
            "ok": bool(ready.get("ok")),
            "reason": "ok" if ready.get("ok") else "pending_table_not_ready_after_cancel",
            "elapsed_ms": elapsed_ms(started),
        }
    )
    return item


def get_main_table(jab: JABOperator, timeout: float) -> dict:
    table_context, vm_id, owned_contexts, table_info = jab_table_ops.find_main_table(
        jab, timeout=timeout
    )
    if not table_context or not table_info:
        return {"ok": False, "reason": "main_table_not_found"}
    return {
        "ok": True,
        "context": table_context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "table_info": table_info,
    }


def select_rows_from_context(
    jab: JABOperator,
    vm_id: int,
    table_context,
    *,
    row_count: int,
    col_count: int,
    selection_col: int,
    start_row: int,
    count: int,
) -> dict:
    if not jab.has_selection_api():
        return {"ok": False, "reason": "selection_api_unavailable"}
    if selection_col < 0 or selection_col >= col_count:
        return {"ok": False, "reason": "selection_col_out_of_range"}
    jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
    expected = []
    for row in range(start_row, start_row + count):
        child_index = row * col_count + selection_col
        expected.append(child_index)
        jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)
    selected = jab.get_selected_child_indexes(vm_id, table_context, row_count * col_count)
    selected_set = set(selected)
    missing = [index for index in expected if index not in selected_set]
    selected_rows = sorted(
        int(index // col_count)
        for index in selected
        if index % col_count == selection_col
    )
    return {
        "ok": not missing,
        "reason": "ok" if not missing else "missing_selected_rows",
        "expected_count": len(expected),
        "selected_index_count": len(selected),
        "selected_row_count_on_selection_col": len(selected_rows),
        "missing_count": len(missing),
        "missing_rows_head": [int(index // col_count) for index in missing[:20]],
        "selected_rows_head": selected_rows[:10],
        "selected_rows_tail": selected_rows[-10:],
    }


def wait_voucher_tables(
    jab: JABOperator,
    voucher_title: str,
    expected_count: int,
    *,
    timeout: float,
) -> dict:
    deadline = time.time() + max(float(timeout), 0.0)
    last = {"ok": False, "reason": "not_started"}
    while True:
        tables = jab.read_window_table_cells(
            voucher_title,
            max_rows=expected_count + 3,
            max_cols=13,
        )
        voucher_tables = [
            table
            for table in tables
            if int(table.get("row_count") or 0) > 0 and int(table.get("col_count") or 0) == 13
        ]
        total_rows = sum(int(table.get("row_count") or 0) for table in voucher_tables)
        last = {
            "ok": bool(voucher_tables and total_rows >= expected_count),
            "reason": "ok" if voucher_tables and total_rows >= expected_count else "voucher_table_rows_short",
            "table_count": len(voucher_tables),
            "total_rows": total_rows,
            "expected_rows": expected_count,
            "tables": [
                {
                    "table_index": table.get("table_index"),
                    "row_count": table.get("row_count"),
                    "col_count": table.get("col_count"),
                    "window_title": table.get("window_title"),
                    "window_class": table.get("window_class"),
                }
                for table in voucher_tables
            ],
        }
        if last["ok"] or time.time() >= deadline:
            return last
        time.sleep(0.2)


def cleanup_after_generate(
    jab: JABOperator,
    voucher_title: str,
    voucher_class: str,
    args,
) -> dict:
    if args.auto_close_voucher:
        jab.close_window_by_title(voucher_title, class_name=voucher_class, wait=0.5)
        return wait_voucher_closed(
            jab,
            voucher_title,
            voucher_class,
            timeout=args.cancel_timeout,
            mode="auto_close",
        )

    print()
    print("已打开制单窗口且未保存。请在 NC 中手工取消/关闭制单窗口，确认回到待生成页后按回车继续。")
    print("如果出现是否保存提示，请选择不保存/否。不要点保存。")
    input("手工取消完成后按回车继续...")
    return wait_voucher_closed(
        jab,
        voucher_title,
        voucher_class,
        timeout=args.cancel_timeout,
        mode="manual_cancel",
    )


def wait_voucher_closed(
    jab: JABOperator,
    voucher_title: str,
    voucher_class: str,
    *,
    timeout: float,
    mode: str,
) -> dict:
    deadline = time.time() + max(float(timeout), 0.0)
    while time.time() <= deadline:
        exists = jab.window_exists(voucher_title, class_name=voucher_class)
        if not exists:
            return {"ok": True, "mode": mode, "reason": "closed"}
        time.sleep(0.2)
    return {"ok": False, "mode": mode, "reason": "still_open"}


def wait_main_table_ready(jab: JABOperator, timeout: float) -> dict:
    table = get_main_table(jab, timeout)
    if not table.get("ok"):
        return table
    try:
        info = table["table_info"]
        return {
            "ok": True,
            "reason": "main_table_ready",
            "row_count": int(info.rowCount),
            "col_count": int(info.columnCount),
        }
    finally:
        jab.release_contexts(table["vm_id"], table["owned_contexts"])


def build_counts(args) -> list[int]:
    counts = []
    for item in str(args.counts or "").split(","):
        item = item.strip()
        if not item:
            continue
        value = int(item)
        if value > 0 and value not in counts:
            counts.append(value)
    current = int(args.step_start)
    while current <= int(args.max_count):
        if current > 0 and current not in counts:
            counts.append(current)
        current += int(args.step)
    return counts


def write_report(report: dict, stem: str) -> Path:
    path = logs_dir() / f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def print_summary(report: dict) -> None:
    print("NC 前台生成取消重试探测")
    print(f"ok={report.get('ok')} reason={report.get('reason')}")
    print(f"last_ok_count={report.get('last_ok_count')}")
    print(f"first_failed_count={report.get('first_failed_count')}")
    print(f"suspected_front_generate_limit={report.get('suspected_front_generate_limit')}")
    for item in report.get("results") or []:
        voucher = item.get("voucher_window") or {}
        cancel = item.get("cancel") or {}
        print(
            f"- count={item.get('count')} ok={item.get('ok')} reason={item.get('reason')} "
            f"selected={((item.get('selection') or {}).get('selected_row_count_on_selection_col'))} "
            f"voucher_rows={voucher.get('total_rows')} voucher_ok={voucher.get('ok')} "
            f"cancel_ok={cancel.get('ok')} elapsed_ms={item.get('elapsed_ms')}"
        )
    print(f"报告已保存: {report.get('report_path')}")


def elapsed_ms(started: float) -> float:
    return round((time.perf_counter() - started) * 1000, 1)


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


if __name__ == "__main__":
    raise SystemExit(main())
