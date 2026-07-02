# lifecycle: tool（现场诊断脚本）
# 作用：只读扫描当前 NC 已生成查询结果页，定位 JAB 表路径与样例列。
# 不做什么：不点击、不生成、不保存、不回填、不修改 Excel。
# 运行：Windows Python 下在 nc_auto_v2 根目录执行：
#   python tools/voucher_generated_table_probe.py --config config.json --wait-timeout 10

from __future__ import annotations

import argparse
import ctypes
import json
import re
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.jab_probe import JOBJECT, enum_windows  # noqa: E402
from core.nc_generated_result_locator import (  # noqa: E402
    locate_generated_result_table,
    wait_generated_result_table,
)
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402

WATCH_CLASSES = {"YonyouUWnd", "SunAwtFrame", "SunAwtCanvas", "SunAwtDialog"}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="NC 凭证已生成结果表 JAB 只读探测。"
    )
    parser.add_argument("--config", default="config.json", help="配置文件路径")
    parser.add_argument("--wait-timeout", type=float, default=10.0, help="正式定位器等待秒数")
    parser.add_argument("--interval", type=float, default=0.5, help="正式定位器轮询间隔")
    parser.add_argument("--sample-rows", type=int, default=5, help="每张表读取前 N 行样例")
    parser.add_argument("--sample-cols", type=int, default=8, help="每张表读取前 N 列样例")
    parser.add_argument("--min-rows", type=int, default=1, help="只输出至少 N 行的表")
    parser.add_argument("--max-depth", type=int, default=None, help="覆盖 JAB 树扫描深度")
    parser.add_argument("--max-children", type=int, default=None, help="覆盖每层最大子节点数")
    parser.add_argument("--json", action="store_true", help="stdout 输出完整 JSON")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    cfg = load_config(args.config)
    if args.max_depth is not None:
        cfg.setdefault("jab", {})["max_depth"] = args.max_depth
    if args.max_children is not None:
        cfg.setdefault("jab", {})["max_children"] = args.max_children

    batch_cfg = cfg.get("jab_batch", {})
    voucher_col = int(batch_cfg.get("generated_voucher_col", 22))
    date_col = int(batch_cfg.get("generated_date_col", 18))
    generated_voucher_max = int(batch_cfg.get("generated_voucher_max", 9999))

    jab = JABOperator(cfg)
    try:
        jab.ensure_started()
        report = {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "config": str(Path(args.config).resolve()),
            "loaded_jab": str(jab.loaded_path),
            "probe": {
                "wait_timeout": args.wait_timeout,
                "interval": args.interval,
                "sample_rows": args.sample_rows,
                "sample_cols": args.sample_cols,
                "voucher_col": voucher_col,
                "generated_date_col": date_col,
            },
            "windows": list_nc_windows(jab),
            "official_locator_once": locate_generated_result_table(
                jab,
                batch_cfg,
                voucher_col,
                generated_voucher_max,
            ),
            "official_locator_wait": wait_generated_result_table(
                jab,
                batch_cfg,
                voucher_col,
                generated_voucher_max,
                timeout=args.wait_timeout,
                interval=args.interval,
            ),
            "tables": collect_tables(
                jab,
                sample_rows=args.sample_rows,
                sample_cols=args.sample_cols,
                min_rows=args.min_rows,
                voucher_col=voucher_col,
                date_col=date_col,
                generated_voucher_max=generated_voucher_max,
            ),
        }
    finally:
        jab.close()

    report_path = write_report(report, "voucher_generated_table_probe")
    report["report_path"] = str(report_path)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_summary(report)
    return 0


def list_nc_windows(jab: JABOperator) -> list[dict]:
    result = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if class_name not in WATCH_CLASSES:
            continue
        is_java = False
        try:
            is_java = bool(jab.dll.isJavaWindow(hwnd))
        except Exception:
            pass
        result.append(
            {
                "hwnd": int(hwnd),
                "pid": int(pid),
                "class_name": class_name,
                "title": title,
                "visible": bool(visible),
                "is_java": is_java,
            }
        )
    return result


def collect_tables(
    jab: JABOperator,
    *,
    sample_rows: int,
    sample_cols: int,
    min_rows: int,
    voucher_col: int,
    date_col: int,
    generated_voucher_max: int,
) -> list[dict]:
    tables: list[dict] = []
    for window in list_nc_windows(jab):
        if not window["is_java"]:
            continue
        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            window["hwnd"], ctypes.byref(vm_id), ctypes.byref(root_context)
        ):
            continue
        try:
            walk_tree(
                jab,
                vm_id.value,
                root_context.value,
                window=window,
                path=[],
                depth=0,
                owned=[],
                tables=tables,
                sample_rows=sample_rows,
                sample_cols=sample_cols,
                min_rows=min_rows,
                voucher_col=voucher_col,
                date_col=date_col,
                generated_voucher_max=generated_voucher_max,
            )
        finally:
            jab.release_contexts(vm_id.value, [root_context.value])
    tables.sort(
        key=lambda item: (item.get("row_count", 0), item.get("col_count", 0)),
        reverse=True,
    )
    return tables


def walk_tree(
    jab: JABOperator,
    vm_id: int,
    context,
    *,
    window: dict,
    path: list[int],
    depth: int,
    owned: list,
    tables: list[dict],
    sample_rows: int,
    sample_cols: int,
    min_rows: int,
    voucher_col: int,
    date_col: int,
    generated_voucher_max: int,
) -> None:
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    role = (info.role_en_US.strip() or info.role.strip()).lower()
    if role == "table":
        table_info = jab.get_table_info(vm_id, context)
        if table_info and table_info.rowCount >= min_rows and table_info.columnCount > 0:
            tables.append(
                describe_table(
                    jab,
                    vm_id,
                    context,
                    table_info,
                    window=window,
                    path=path,
                    sample_rows=sample_rows,
                    sample_cols=sample_cols,
                    voucher_col=voucher_col,
                    date_col=date_col,
                )
            )
        return

    if depth >= jab.max_depth:
        return

    child_count = min(int(info.childrenCount), int(jab.max_children))
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        try:
            walk_tree(
                jab,
                vm_id,
                child,
                window=window,
                path=[*path, index],
                depth=depth + 1,
                owned=[*owned, child],
                tables=tables,
                sample_rows=sample_rows,
                sample_cols=sample_cols,
                min_rows=min_rows,
                voucher_col=voucher_col,
                date_col=date_col,
                generated_voucher_max=generated_voucher_max,
            )
        finally:
            jab.release_contexts(vm_id, [child])


def describe_table(
    jab: JABOperator,
    vm_id: int,
    context,
    table_info,
    *,
    window: dict,
    path: list[int],
    sample_rows: int,
    sample_cols: int,
    voucher_col: int,
    date_col: int,
    generated_voucher_max: int,
) -> dict:
    rows = int(table_info.rowCount)
    cols = int(table_info.columnCount)
    sample = []
    for row in range(min(rows, sample_rows)):
        cells = []
        for col in range(min(cols, sample_cols)):
            cells.append(jab.get_table_cell_text(vm_id, context, row, col).strip())
        sample.append(cells)

    voucher_values = read_column_sample(jab, vm_id, context, rows, cols, voucher_col, sample_rows)
    date_values = read_column_sample(jab, vm_id, context, rows, cols, date_col, sample_rows)
    column_profiles = profile_columns(
        jab,
        vm_id,
        context,
        rows,
        cols,
        sample_rows=max(sample_rows, 10),
        generated_voucher_max=generated_voucher_max,
    )
    return {
        "path": ".".join(str(item) for item in path),
        "window": window,
        "row_count": rows,
        "col_count": cols,
        "sample_cells": sample,
        "voucher_col": voucher_col,
        "voucher_values": voucher_values,
        "generated_date_col": date_col,
        "generated_date_values": date_values,
        "voucher_col_candidates": [
            item for item in column_profiles
            if item["voucher_like_count"] > 0
        ][:8],
        "date_col_candidates": [
            item for item in column_profiles
            if item["date_like_count"] > 0
        ][:8],
        "column_profiles": column_profiles,
    }


def read_column_sample(
    jab: JABOperator,
    vm_id: int,
    context,
    rows: int,
    cols: int,
    col: int,
    limit: int,
) -> list[str]:
    if col < 0 or col >= cols:
        return []
    return [
        jab.get_table_cell_text(vm_id, context, row, col).strip()
        for row in range(min(rows, limit))
    ]


def profile_columns(
    jab: JABOperator,
    vm_id: int,
    context,
    rows: int,
    cols: int,
    *,
    sample_rows: int,
    generated_voucher_max: int,
) -> list[dict]:
    profiles = []
    max_rows = min(rows, sample_rows)
    for col in range(cols):
        values = [
            jab.get_table_cell_text(vm_id, context, row, col).strip()
            for row in range(max_rows)
        ]
        non_empty = [value for value in values if value]
        voucher_like = [
            value for value in non_empty
            if is_strict_voucher_text(value, generated_voucher_max)
        ]
        date_like = [value for value in non_empty if is_date_text(value)]
        profiles.append(
            {
                "col": col,
                "non_empty_count": len(non_empty),
                "voucher_like_count": len(voucher_like),
                "date_like_count": len(date_like),
                "sample": non_empty[:5],
            }
        )
    profiles.sort(
        key=lambda item: (
            item["voucher_like_count"],
            item["date_like_count"],
            item["non_empty_count"],
        ),
        reverse=True,
    )
    return profiles


def is_strict_voucher_text(text: str, generated_voucher_max: int) -> bool:
    value = str(text or "").strip()
    if not re.fullmatch(r"\d+", value):
        return False
    number = int(value)
    return 1 <= number <= int(generated_voucher_max)


def is_date_text(text: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(text or "").strip()))


def write_report(report: dict, stem: str) -> Path:
    path = logs_dir() / f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def print_summary(report: dict) -> None:
    print("NC 凭证已生成结果表探测")
    print(f"JAB: {report.get('loaded_jab')}")
    wait = report.get("official_locator_wait") or {}
    print(f"正式定位器: ok={wait.get('ok')} reason={wait.get('reason')}")
    if wait.get("ok"):
        print(
            "  path="
            f"{wait.get('table_path')} rows={wait.get('row_count')} cols={wait.get('col_count')}"
        )
    else:
        print(f"  attempts={len(wait.get('attempts') or [])}")
        for attempt in (wait.get("attempts") or [])[-5:]:
            print(f"  - {attempt}")
    print()
    print(f"扫描到 table: {len(report.get('tables') or [])}")
    for index, table in enumerate((report.get("tables") or [])[:20], start=1):
        window = table.get("window") or {}
        print(
            f"{index}. path={table.get('path')} rows={table.get('row_count')} "
            f"cols={table.get('col_count')} class={window.get('class_name')} "
            f"title={window.get('title')!r}"
        )
        print(f"   voucher_values={table.get('voucher_values')}")
        print(f"   date_values={table.get('generated_date_values')}")
        print(f"   voucher_col_candidates={table.get('voucher_col_candidates')}")
        print(f"   date_col_candidates={table.get('date_col_candidates')}")
    print()
    print(f"报告已保存: {report.get('report_path')}")


if __name__ == "__main__":
    raise SystemExit(main())
