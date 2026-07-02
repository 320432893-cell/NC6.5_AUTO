# lifecycle: tool（现场诊断脚本）
# 作用：只测试 NC 待生成表 JAB 批量选中上限，不点击生成、不保存。
# 风险：会改变当前 NC 表格选中状态；默认测试结束后清空选中。
# 运行：Windows Python 下在 nc_auto_v2 根目录执行：
#   python tools/voucher_selection_limit_probe.py --config config.json

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
        description="NC 待生成表 JAB 批量选中上限探测，不点击生成。"
    )
    parser.add_argument("--config", default="config.json", help="配置文件路径")
    parser.add_argument("--start-row", type=int, default=0, help="从 NC 表第几行开始选，0 基")
    parser.add_argument("--counts", default="200,250", help="先测的数量，逗号分隔")
    parser.add_argument("--step-start", type=int, default=260, help="阶梯测试起始数量")
    parser.add_argument("--step", type=int, default=10, help="阶梯递增数量")
    parser.add_argument("--max-count", type=int, default=500, help="最多测到多少行")
    parser.add_argument("--selection-col", type=int, default=None, help="覆盖选中列，默认用 config 的 jab.selection_col")
    parser.add_argument("--timeout", type=float, default=5.0, help="等待主表秒数")
    parser.add_argument("--sleep", type=float, default=0.15, help="每档测试后的暂停秒数")
    parser.add_argument("--continue-after-failure", action="store_true", help="失败后继续测后续数量")
    parser.add_argument("--leave-selected", action="store_true", help="结束后保留最后一次选中状态")
    parser.add_argument("--json", action="store_true", help="stdout 输出完整 JSON")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    started = time.perf_counter()
    try:
        jab.ensure_started()
        table_context, vm_id, owned_contexts, table_info = jab_table_ops.find_main_table(
            jab, timeout=args.timeout
        )
        if not table_context or not table_info:
            report = {
                "generated_at": now_text(),
                "ok": False,
                "reason": "main_table_not_found",
                "elapsed_s": round(time.perf_counter() - started, 3),
            }
        else:
            try:
                report = run_probe(jab, vm_id, table_context, table_info, args, started)
            finally:
                if not args.leave_selected:
                    clear_selection(jab, vm_id, table_context)
                jab.release_contexts(vm_id, owned_contexts)
    finally:
        jab.close()

    report_path = write_report(report, "voucher_selection_limit_probe")
    report["report_path"] = str(report_path)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_summary(report)
    return 0 if report.get("ok") else 1


def run_probe(jab, vm_id, table_context, table_info, args, started: float) -> dict:
    row_count = int(table_info.rowCount)
    col_count = int(table_info.columnCount)
    selection_col = (
        int(args.selection_col)
        if args.selection_col is not None
        else int(jab.resolve_selection_col(None))
    )
    if selection_col < 0 or selection_col >= col_count:
        return {
            "generated_at": now_text(),
            "ok": False,
            "reason": "selection_col_out_of_range",
            "selection_col": selection_col,
            "row_count": row_count,
            "col_count": col_count,
        }
    if not jab.has_selection_api():
        return {
            "generated_at": now_text(),
            "ok": False,
            "reason": "selection_api_unavailable",
            "row_count": row_count,
            "col_count": col_count,
        }

    counts = build_counts(args)
    results = []
    last_ok = None
    first_failed = None
    for count in counts:
        if args.start_row + count > row_count:
            results.append(
                {
                    "count": count,
                    "ok": False,
                    "reason": "not_enough_rows",
                    "start_row": args.start_row,
                    "row_count": row_count,
                }
            )
            if first_failed is None:
                first_failed = count
            break

        result = try_select_count(
            jab,
            vm_id,
            table_context,
            row_count=row_count,
            col_count=col_count,
            selection_col=selection_col,
            start_row=args.start_row,
            count=count,
        )
        results.append(result)
        if result.get("ok"):
            last_ok = count
        elif first_failed is None:
            first_failed = count
            if not args.continue_after_failure:
                break
        if args.sleep > 0:
            time.sleep(float(args.sleep))

    return {
        "generated_at": now_text(),
        "ok": True,
        "reason": "probe_finished",
        "elapsed_s": round(time.perf_counter() - started, 3),
        "table": {
            "row_count": row_count,
            "col_count": col_count,
            "selection_col": selection_col,
            "start_row": args.start_row,
        },
        "plan": {
            "counts": counts,
            "step": args.step,
            "max_count": args.max_count,
            "continue_after_failure": bool(args.continue_after_failure),
            "clear_after": not bool(args.leave_selected),
        },
        "last_ok_count": last_ok,
        "first_failed_count": first_failed,
        "suspected_limit": last_ok,
        "results": results,
    }


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


def try_select_count(
    jab,
    vm_id: int,
    table_context,
    *,
    row_count: int,
    col_count: int,
    selection_col: int,
    start_row: int,
    count: int,
) -> dict:
    started = time.perf_counter()
    rows = list(range(start_row, start_row + count))
    clear_selection(jab, vm_id, table_context)

    expected_indexes = []
    error = ""
    for row in rows:
        child_index = row * col_count + selection_col
        expected_indexes.append(child_index)
        try:
            jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)
        except Exception as exc:
            error = repr(exc)
            break

    selected_indexes = jab.get_selected_child_indexes(
        vm_id, table_context, row_count * col_count
    )
    selected_set = set(selected_indexes)
    missing = [index for index in expected_indexes if index not in selected_set]
    unexpected = [
        index
        for index in selected_indexes
        if index % col_count == selection_col and index not in expected_indexes
    ]
    selected_rows = sorted(
        {
            int(index // col_count)
            for index in selected_indexes
            if index % col_count == selection_col
        }
    )
    ok = not error and not missing and len(selected_rows) >= count
    return {
        "count": count,
        "ok": ok,
        "reason": "ok" if ok else ("add_selection_error" if error else "missing_selected_rows"),
        "error": error,
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
        "expected_count": len(expected_indexes),
        "selected_index_count": len(selected_indexes),
        "selected_row_count_on_selection_col": len(selected_rows),
        "selected_rows_head": selected_rows[:10],
        "selected_rows_tail": selected_rows[-10:],
        "missing_count": len(missing),
        "missing_rows_head": [int(index // col_count) for index in missing[:20]],
        "unexpected_count": len(unexpected),
    }


def clear_selection(jab, vm_id: int, table_context) -> None:
    try:
        jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
    except Exception:
        pass


def write_report(report: dict, stem: str) -> Path:
    path = logs_dir() / f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def print_summary(report: dict) -> None:
    print("NC 待生成表批量选中上限探测")
    print(f"ok={report.get('ok')} reason={report.get('reason')}")
    if report.get("table"):
        table = report["table"]
        print(
            f"表格: rows={table.get('row_count')} cols={table.get('col_count')} "
            f"selection_col={table.get('selection_col')} start_row={table.get('start_row')}"
        )
    print(f"last_ok_count={report.get('last_ok_count')}")
    print(f"first_failed_count={report.get('first_failed_count')}")
    print(f"suspected_limit={report.get('suspected_limit')}")
    for item in report.get("results") or []:
        print(
            f"- count={item.get('count')} ok={item.get('ok')} "
            f"reason={item.get('reason')} selected_rows={item.get('selected_row_count_on_selection_col')} "
            f"missing={item.get('missing_count')} elapsed_ms={item.get('elapsed_ms')}"
        )
    print(f"报告已保存: {report.get('report_path')}")


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


if __name__ == "__main__":
    raise SystemExit(main())
