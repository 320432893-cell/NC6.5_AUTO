import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import JOBJECT, enum_windows  # noqa: E402


WATCH_NAMES = (
    "单据生成",
    "查询",
    "生成",
    "前台生成",
    "正式单据",
    "确定",
    "取消",
    "制单",
)


def main():
    parser = argparse.ArgumentParser(description="只读枚举 NC/JAB 页面状态特征")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--rows", type=int, default=5)
    parser.add_argument("--cols", type=int, default=25)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        report = build_report(jab, cfg, args.rows, args.cols)
    finally:
        jab.close()

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)


def build_report(jab, cfg, max_rows, max_cols):
    jab.ensure_started()
    windows = collect_java_windows(jab)
    controls = collect_controls(jab)
    tables = jab.read_all_table_cells(max_rows=max_rows, max_cols=max_cols)
    batch_cfg = cfg.get("jab_batch", {})

    voucher_window = batch_cfg.get("voucher_window_title", "制单")
    query_title = batch_cfg.get("open_query", {}).get("dialog_title", "查询")
    blockers = [
        window
        for window in windows
        if window["title"] in (voucher_window, query_title)
        and window["class"] == "SunAwtDialog"
        and window["visible"]
    ]

    return {
        "blocking_child_windows": blockers,
        "parent_markers": [
            item for item in controls if item["name"] == "单据生成"
        ],
        "watched_controls": controls,
        "table_signatures": [
            describe_table(table, batch_cfg) for table in tables
        ],
    }


def collect_java_windows(jab):
    windows = []
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        try:
            is_java = bool(jab.dll.isJavaWindow(hwnd))
        except Exception:
            is_java = False
        if is_java:
            windows.append(
                {
                    "hwnd": int(hwnd),
                    "title": title,
                    "class": class_name,
                    "pid": pid,
                    "visible": bool(visible),
                }
            )
    return windows


def collect_controls(jab):
    found = []
    seen = set()
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if not visible or not jab.dll.isJavaWindow(hwnd):
            continue
        vm_id, root = get_root(jab, hwnd)
        if root is None:
            continue
        collect_controls_in_tree(
            jab,
            vm_id,
            root,
            [],
            {
                "hwnd": int(hwnd),
                "title": title,
                "class": class_name,
                "pid": pid,
                "visible": bool(visible),
            },
            found,
            seen,
            0,
        )
    return found


def get_root(jab, hwnd):
    import ctypes

    vm_id = ctypes.c_long()
    root_context = JOBJECT()
    if not jab.dll.getAccessibleContextFromHWND(
        hwnd,
        ctypes.byref(vm_id),
        ctypes.byref(root_context),
    ):
        return None, None
    return vm_id.value, root_context.value


def collect_controls_in_tree(jab, vm_id, context, path, window, found, seen, depth):
    info = jab.get_context_info(vm_id, context)
    if not info:
        return

    name = info.name.strip()
    desc = info.description.strip()
    role = (info.role_en_US.strip() or info.role.strip())
    states = (info.states_en_US.strip() or info.states.strip())
    role_l = role.lower()
    states_l = states.lower()
    text = name or desc

    if text in WATCH_NAMES or any(marker in text for marker in WATCH_NAMES):
        key = (window["hwnd"], ".".join(map(str, path)), name, desc, role)
        if key not in seen:
            seen.add(key)
            found.append(
                {
                    "window_title": window["title"],
                    "window_class": window["class"],
                    "path": ".".join(map(str, path)),
                    "name": name,
                    "description": desc,
                    "role": role,
                    "states": states,
                    "showing": "visible" in states_l and "showing" in states_l,
                    "bounds": [info.x, info.y, info.width, info.height],
                }
            )

    if depth >= jab.max_depth or role_l == "table":
        return

    child_count = min(info.childrenCount, jab.max_children)
    for index in range(child_count):
        child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
        if not child:
            continue
        collect_controls_in_tree(
            jab, vm_id, child, path + [index], window, found, seen, depth + 1
        )
        jab.release_contexts(vm_id, [child])


def describe_table(table, batch_cfg):
    generated_date_col = batch_cfg.get("generated_date_col", 18)
    voucher_col = batch_cfg.get("generated_voucher_col", 22)
    date_values = sample_col(table, generated_date_col)
    voucher_values = sample_col(table, voucher_col)
    return {
        "table_index": table["table_index"],
        "window_title": table.get("window_title"),
        "window_class": table.get("window_class"),
        "row_count": table["row_count"],
        "col_count": table["col_count"],
        "date_col": generated_date_col,
        "date_values": date_values,
        "voucher_col": voucher_col,
        "voucher_values": voucher_values,
        "sample_rows": [
            {
                "row_index": row["row_index"],
                "cells": row["cells"],
                "selected": row["selected"],
            }
            for row in table["rows"]
        ],
    }


def sample_col(table, col):
    values = []
    if col is None:
        return values
    for row in table["rows"]:
        cells = row["cells"]
        if 0 <= col < len(cells):
            text = str(cells[col]).strip()
            if text:
                values.append(text)
    return values[:8]


def print_text(report):
    blockers = report["blocking_child_windows"]
    print("blocking_child_windows:", len(blockers))
    for item in blockers:
        print(
            f"  {item['title']!r} {item['class']!r} "
            f"hwnd={item['hwnd']} visible={item['visible']}"
        )

    print("parent_markers:", len(report["parent_markers"]))
    for item in report["parent_markers"][:20]:
        print(
            f"  path={item['path']} name={item['name']!r} role={item['role']!r} "
            f"showing={item['showing']} win={item['window_title']!r}/{item['window_class']!r}"
        )

    print("watched_controls:", len(report["watched_controls"]))
    for item in report["watched_controls"][:80]:
        print(
            f"  path={item['path']} name={item['name']!r} desc={item['description']!r} "
            f"role={item['role']!r} showing={item['showing']} "
            f"win={item['window_title']!r}/{item['window_class']!r}"
        )

    print("table_signatures:", len(report["table_signatures"]))
    for table in report["table_signatures"]:
        print(
            f"  table={table['table_index']} win={table['window_title']!r}/"
            f"{table['window_class']!r} rows={table['row_count']} cols={table['col_count']}"
        )
        print(f"    date_col[{table['date_col']}]: {table['date_values']}")
        print(f"    voucher_col[{table['voucher_col']}]: {table['voucher_values']}")
        for row in table["sample_rows"][:3]:
            print(f"    row {row['row_index']}: {row['cells']}")


if __name__ == "__main__":
    main()
