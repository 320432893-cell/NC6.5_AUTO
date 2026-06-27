import argparse
import ctypes
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from core.jab_probe import AccessibleTableCellInfo  # noqa: E402
from tools.receipt_body_table_locator import (  # noqa: E402
    KEY_COLUMNS,
    locate_receipt_body_table_cached,
    read_key_rows,
)


COLUMN_LABELS = {
    0: "往来对象",
    1: "收款业务类型",
    2: "收款性质",
    3: "币种",
    4: "收款银行账户",
    5: "科目",
    6: "汇率",
    7: "贷方原币金额",
    8: "本币金额",
    11: "结算方式",
    13: "订单客户",
    19: "客户",
}


def parse_indexes(text):
    indexes = []
    for part in str(text or "").split(","):
        part = part.strip()
        if part:
            indexes.append(int(part))
    return indexes


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Read-only probe for receipt body exchange-rate column. It compares "
            "formal key-row snapshots with direct getAccessibleTableCellInfo reads."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--rows", default="0", help="Comma-separated 0-based rows.")
    parser.add_argument(
        "--cols",
        default="5,6,7",
        help="Comma-separated 0-based columns; default probes subject/rate/amount.",
    )
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--interval", type=float, default=0.12)
    parser.add_argument("--max-rows", type=int, default=3)
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument("--json", action="store_true")
    return parser.parse_args(argv)


def main(argv=None):
    stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
    if callable(stdout_reconfigure):
        stdout_reconfigure(encoding="utf-8", errors="replace")

    args = parse_args(argv)
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report = probe(
            jab,
            rows=parse_indexes(args.rows),
            cols=parse_indexes(args.cols),
            repeat=max(1, args.repeat),
            interval=max(0.0, args.interval),
            max_rows=args.max_rows,
            scope_hwnd=args.scope_hwnd,
        )
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    return 0 if report.get("ok") else 1


def probe(jab, rows, cols, repeat, interval, max_rows, scope_hwnd):
    report = {
        "ok": False,
        "read_only": True,
        "requested": {
            "rows": rows,
            "cols": cols,
            "repeat": repeat,
            "interval": interval,
            "max_rows": max_rows,
            "scope_hwnd": scope_hwnd,
        },
        "formal_key_columns": {
            str(col): label for col, label in sorted(KEY_COLUMNS.items())
        },
        "diagnosis": [],
        "located": None,
        "table": None,
        "iterations": [],
    }

    located = locate_receipt_body_table_cached(
        jab,
        cached=None,
        max_rows=max(max_rows, max(rows, default=0) + 1),
        scope_hwnd=scope_hwnd,
    )
    best = located.get("best")
    report["located"] = summarize_located(located)
    if not best:
        report["diagnosis"].append("收款单明细表未定位到，无法判断第 6 列")
        report["reason"] = "receipt body table not located"
        return report

    context, vm_id, owned, window = jab.find_context_by_path_once(
        best["path"],
        class_name=(best.get("window") or {}).get("class_name"),
        scope_hwnd=scope_hwnd or ((best.get("window") or {}).get("hwnd")),
        role="table",
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        report["diagnosis"].append("明细表 path 二次打开失败")
        report["reason"] = "table context not found by located path"
        return report

    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            report["diagnosis"].append("明细表 context 命中，但 table_info 不可读")
            report["reason"] = "getAccessibleTableInfo failed"
            return report

        report["table"] = {
            "path": best.get("path"),
            "window": window or best.get("window"),
            "row_count": int(table_info.rowCount),
            "col_count": int(table_info.columnCount),
            "context": describe_context(jab, vm_id, context),
        }
        for index in range(repeat):
            started_at = time.perf_counter()
            formal_rows = read_key_rows(
                jab,
                vm_id,
                context,
                table_info,
                max_rows=max(max_rows, max(rows, default=0) + 1),
            )
            direct_rows = []
            for row in rows:
                direct_rows.append(
                    {
                        "row_index": row,
                        "cells": [
                            read_direct_cell(jab, vm_id, context, table_info, row, col)
                            for col in cols
                        ],
                    }
                )
            report["iterations"].append(
                {
                    "index": index,
                    "seconds": round(time.perf_counter() - started_at, 3),
                    "formal_key_snapshot": {
                        "has_col6_in_config": 6 in KEY_COLUMNS,
                        "rows": [
                            {
                                "row_index": row.get("row_index"),
                                "cells": row.get("cells") or {},
                                "has_col6": "6" in (row.get("cells") or {}),
                                "col6": (row.get("cells") or {}).get("6"),
                            }
                            for row in formal_rows
                            if int(row.get("row_index") or 0) in set(rows)
                        ],
                    },
                    "direct_cell_reads": direct_rows,
                }
            )
            if index < repeat - 1 and interval:
                time.sleep(interval)
    finally:
        jab.release_contexts(vm_id, owned)

    report["diagnosis"] = diagnose(report)
    report["ok"] = True
    return report


def read_direct_cell(jab, vm_id, table_context, table_info, row, col):
    item = {
        "row": row,
        "col": col,
        "label": COLUMN_LABELS.get(col),
        "ok": False,
        "text_sources": {},
        "cell_info": None,
        "context": None,
        "reason": None,
    }
    if row < 0 or row >= int(table_info.rowCount):
        item["reason"] = f"row out of range: {row}"
        return item
    if col < 0 or col >= int(table_info.columnCount):
        item["reason"] = f"col out of range: {col}"
        return item
    if not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        item["reason"] = "JAB DLL does not expose getAccessibleTableCellInfo"
        return item

    cell_info = AccessibleTableCellInfo()
    api_ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    item["cell_info"] = {
        "api_ok": bool(api_ok),
        "accessible_context": int(cell_info.accessibleContext or 0),
        "index": int(cell_info.index),
        "row": int(cell_info.row),
        "column": int(cell_info.column),
        "row_extent": int(cell_info.rowExtent),
        "column_extent": int(cell_info.columnExtent),
        "is_selected": bool(cell_info.isSelected),
    }
    if not api_ok:
        item["reason"] = "getAccessibleTableCellInfo returned false"
        return item
    if not cell_info.accessibleContext:
        item["reason"] = "cell has no accessibleContext"
        return item

    context = int(cell_info.accessibleContext)
    info = jab.get_context_info(vm_id, context)
    if not info:
        item["reason"] = "getAccessibleContextInfo failed for cell"
        return item

    text_value = jab.get_text_context_value(vm_id, context)
    name = (info.name or "").strip()
    description = (info.description or "").strip()
    item["context"] = describe_context_from_info(jab, vm_id, context, info, text_value)
    item["text_sources"] = {
        "name": name,
        "description": description,
        "accessible_text": text_value,
        "first_nonempty": first_nonempty(name, description, text_value),
    }
    item["ok"] = True
    return item


def describe_context(jab, vm_id, context):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return None
    text_value = jab.get_text_context_value(vm_id, context)
    return describe_context_from_info(jab, vm_id, context, info, text_value)


def describe_context_from_info(jab, vm_id, context, info, text_value):
    role = (info.role_en_US or info.role or "").strip()
    states = (info.states_en_US or info.states or "").strip()
    actions = []
    if info.accessibleAction and hasattr(jab.dll, "getAccessibleActions"):
        actions = get_action_names(jab, vm_id, context)
    return {
        "context": int(context),
        "role": role,
        "name": (info.name or "").strip(),
        "description": (info.description or "").strip(),
        "text": text_value,
        "states": states,
        "bounds": [int(info.x), int(info.y), int(info.width), int(info.height)],
        "children": int(info.childrenCount),
        "accessible_text": bool(info.accessibleText),
        "accessible_action": bool(info.accessibleAction),
        "actions": actions,
    }


def get_action_names(jab, vm_id, context):
    from core.jab_probe import AccessibleActions

    actions = AccessibleActions()
    if not jab.dll.getAccessibleActions(vm_id, context, ctypes.byref(actions)):
        return []
    return [
        actions.actionInfo[index].name.strip()
        for index in range(int(actions.actionsCount))
    ]


def first_nonempty(*values):
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def summarize_located(located):
    best = located.get("best") or {}
    return {
        "cache_hit": bool(located.get("cache_hit")),
        "fallback_used": bool(located.get("fallback_used")),
        "best": {
            "path": best.get("path"),
            "window": best.get("window"),
            "row_count": best.get("row_count"),
            "col_count": best.get("col_count"),
            "score": best.get("score"),
            "reasons": best.get("reasons"),
            "cache_hit": best.get("cache_hit"),
            "validated_by_path": best.get("validated_by_path"),
        }
        if best
        else None,
        "path_validation": located.get("path_validation"),
    }


def diagnose(report):
    messages = []
    if "6" not in report.get("formal_key_columns", {}):
        messages.append("正式 KEY_COLUMNS 不包含第 6 列，所以 pipeline/key 快照默认不会出现 cells['6']")

    direct_col6 = []
    direct_neighbors = []
    formal_missing_col6 = False
    formal_has_col6 = False
    for iteration in report.get("iterations") or []:
        for row in (
            (iteration.get("formal_key_snapshot") or {}).get("rows") or []
        ):
            if not row.get("has_col6"):
                formal_missing_col6 = True
                messages.append(
                    f"第 {iteration.get('index')} 次正式快照 row{row.get('row_index')} 未采集第 6 列"
                )
                break
            formal_has_col6 = True
        for direct_row in iteration.get("direct_cell_reads") or []:
            for cell in direct_row.get("cells") or []:
                if int(cell.get("col")) == 6:
                    direct_col6.append(cell)
                elif int(cell.get("col")) in (5, 7):
                    direct_neighbors.append(cell)

    col6_ok = [cell for cell in direct_col6 if cell.get("ok")]
    col6_text = [
        ((cell.get("text_sources") or {}).get("first_nonempty") or "")
        for cell in col6_ok
    ]
    if direct_col6 and not col6_ok:
        messages.append("直接 getAccessibleTableCellInfo 读取第 6 列全部失败")
    elif col6_ok and not any(text.strip() for text in col6_text):
        messages.append("直接读取第 6 列可拿到 context，但 name/description/text 都为空")
    elif any(text.strip() for text in col6_text) and formal_missing_col6:
        messages.append("直接读取第 6 列能拿到非空文本，正式快照缺失主要是采集列集合问题")
    elif any(text.strip() for text in col6_text) and formal_has_col6:
        messages.append("正式快照和直接读取都能拿到第 6 列文本")

    neighbor_ok = [cell for cell in direct_neighbors if cell.get("ok")]
    if direct_neighbors and not neighbor_ok:
        messages.append("邻列 5/7 也直接读取失败，优先怀疑表格定位或 JAB 状态")
    elif neighbor_ok:
        messages.append("邻列 5/7 可直接读取，用来确认表格/行定位是否正确")
    return list(dict.fromkeys(messages))


def print_text(report):
    print(f"ok={report.get('ok')} read_only={report.get('read_only')}")
    for message in report.get("diagnosis") or []:
        print(f"- {message}")
    table = report.get("table") or {}
    if table:
        print(
            "table: "
            f"rows={table.get('row_count')} cols={table.get('col_count')} "
            f"path={table.get('path')}"
        )
    for iteration in report.get("iterations") or []:
        print(f"iteration {iteration.get('index')} {iteration.get('seconds')}s")
        for direct_row in iteration.get("direct_cell_reads") or []:
            parts = []
            for cell in direct_row.get("cells") or []:
                source = cell.get("text_sources") or {}
                parts.append(
                    f"col{cell.get('col')}({cell.get('label')}):"
                    f"ok={cell.get('ok')},text={source.get('first_nonempty')!r},"
                    f"reason={cell.get('reason')}"
                )
            print(f"  row{direct_row.get('row_index')}: " + " | ".join(parts))


if __name__ == "__main__":
    raise SystemExit(main())
