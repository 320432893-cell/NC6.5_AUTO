# lifecycle: tool（现场诊断脚本）
# 作用：随机少量选择待生成行，点击“生成 -> 前台生成”，人工取消制单窗口后重复验证。
# 风险：会真实点击 NC 的“生成 -> 前台生成”，但不会保存凭证；默认要求人工关闭/取消制单窗口。
# 运行：Windows Python 下在 nc_auto_v2 根目录执行：
#   python tools/voucher_random_generate_cancel_probe.py --config config.json --yes-front-generate

from __future__ import annotations

import argparse
import json
import random
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TOOLS = Path(__file__).resolve().parent
for path in (ROOT, TOOLS):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from core.jab_operator import JABOperator  # noqa: E402
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402
from voucher_front_generate_cancel_probe import (  # noqa: E402
    cleanup_after_generate,
    get_main_table,
    wait_main_table_ready,
    wait_voucher_tables,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="NC 凭证随机小批量前台生成取消探测：只打开制单窗口，不保存。"
    )
    parser.add_argument("--config", default="config.json", help="配置文件路径")
    parser.add_argument("--rounds", type=int, default=3, help="循环轮数，默认 3")
    parser.add_argument("--count", type=int, default=5, help="每轮随机选几行，默认 5")
    parser.add_argument("--seed", type=int, default=None, help="随机种子；不传则每次不同")
    parser.add_argument("--selection-col", type=int, default=None, help="覆盖选中列，默认用 config 的 jab.selection_col")
    parser.add_argument("--main-table-timeout", type=float, default=5.0, help="等待待生成主表秒数")
    parser.add_argument("--voucher-timeout", type=float, default=15.0, help="等待制单窗口表格秒数")
    parser.add_argument("--cancel-timeout", type=float, default=60.0, help="等待人工取消/关闭制单窗口秒数")
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
    if args.rounds <= 0:
        raise SystemExit("--rounds 必须是正整数")
    if args.count <= 0:
        raise SystemExit("--count 必须是正整数")

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    started = time.perf_counter()
    try:
        jab.ensure_started()
        report = run_probe(jab, cfg, args, started)
    finally:
        jab.close()

    report_path = write_report(report)
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
    rng = random.Random(args.seed)
    results = []

    for round_index in range(1, args.rounds + 1):
        result = run_one_round(
            jab,
            args,
            rng=rng,
            round_index=round_index,
            voucher_title=voucher_title,
            voucher_class=voucher_class,
        )
        results.append(result)
        if not result.get("ok"):
            break

    return {
        "generated_at": now_text(),
        "ok": all(item.get("ok") for item in results) and len(results) == args.rounds,
        "reason": "probe_finished" if len(results) == args.rounds else "probe_stopped_on_failure",
        "elapsed_s": round(time.perf_counter() - started, 3),
        "plan": {
            "rounds": args.rounds,
            "count": args.count,
            "seed": args.seed,
            "manual_cancel": not bool(args.auto_close_voucher),
        },
        "results": results,
    }


def run_one_round(
    jab: JABOperator,
    args,
    *,
    rng: random.Random,
    round_index: int,
    voucher_title: str,
    voucher_class: str,
) -> dict:
    item = {
        "round": round_index,
        "count": args.count,
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
        }
        if args.count > row_count:
            item.update(
                {
                    "reason": "not_enough_pending_rows",
                    "row_count": row_count,
                    "elapsed_ms": elapsed_ms(started),
                }
            )
            return item
        if selection_col < 0 or selection_col >= col_count:
            item.update(
                {
                    "reason": "selection_col_out_of_range",
                    "elapsed_ms": elapsed_ms(started),
                }
            )
            return item

        rows = sorted(rng.sample(range(row_count), args.count))
        item["random_rows"] = rows
        selected = select_specific_rows(
            jab,
            vm_id,
            table_context,
            row_count=row_count,
            col_count=col_count,
            selection_col=selection_col,
            rows=rows,
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

    voucher = wait_voucher_tables(
        jab,
        voucher_title,
        args.count,
        timeout=args.voucher_timeout,
    )
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


def select_specific_rows(
    jab: JABOperator,
    vm_id: int,
    table_context,
    *,
    row_count: int,
    col_count: int,
    selection_col: int,
    rows: list[int],
) -> dict:
    if not jab.has_selection_api():
        return {"ok": False, "reason": "selection_api_unavailable"}
    jab.dll.clearAccessibleSelectionFromContext(vm_id, table_context)
    expected = []
    error = ""
    for row in rows:
        child_index = row * col_count + selection_col
        expected.append(child_index)
        try:
            jab.dll.addAccessibleSelectionFromContext(vm_id, table_context, child_index)
        except Exception as exc:
            error = repr(exc)
            break
    selected = jab.get_selected_child_indexes(vm_id, table_context, row_count * col_count)
    selected_set = set(selected)
    missing = [index for index in expected if index not in selected_set]
    selected_rows = sorted(
        int(index // col_count)
        for index in selected
        if index % col_count == selection_col
    )
    return {
        "ok": not error and not missing,
        "reason": "ok" if not error and not missing else ("add_selection_error" if error else "missing_selected_rows"),
        "error": error,
        "expected_rows": rows,
        "selected_rows": selected_rows,
        "missing_rows": [int(index // col_count) for index in missing],
        "selected_index_count": len(selected),
    }


def write_report(report: dict) -> Path:
    path = logs_dir() / f"voucher_random_generate_cancel_probe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def print_summary(report: dict) -> None:
    print("NC 随机小批量前台生成取消探测")
    print(f"ok={report.get('ok')} reason={report.get('reason')}")
    plan = report.get("plan") or {}
    print(
        f"计划: rounds={plan.get('rounds')} count={plan.get('count')} "
        f"seed={plan.get('seed')} manual_cancel={plan.get('manual_cancel')}"
    )
    for item in report.get("results") or []:
        voucher = item.get("voucher_window") or {}
        cancel = item.get("cancel") or {}
        print(
            f"- round={item.get('round')} ok={item.get('ok')} reason={item.get('reason')} "
            f"rows={item.get('random_rows')} selected={((item.get('selection') or {}).get('selected_rows'))} "
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
