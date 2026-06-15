# 职责: 只读探测收款单查询结果页表格和分页控件 path 规律
# 不做什么: 不填写查询条件，不改每页条数，不点击下一页，不写 Excel
# 允许依赖层: core JAB/config、tools.receipt_query_pagination_paths
# 谁不应该 import: 正式流程、core 模块和测试不应 import 本临时探针
# 生命周期: T0 临时探针（删除条件：查询结果分页 path 规律完成现场复核并沉淀到正式模块/文档）

import argparse
import ctypes
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready  # noqa: E402
from tools.jab_probe import JOBJECT  # noqa: E402
from tools.receipt_query_pagination import (  # noqa: E402
    read_page_label,
    read_page_size_text,
)
from tools.receipt_query_pagination_paths import (  # noqa: E402
    RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX,
    RECEIPT_PAGE_LABEL_PATH_SUFFIX,
    RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX,
    RECEIPT_RESULT_TABLE_PATH_SUFFIX,
    infer_result_area_prefix_from_table_path,
    join_context_path,
    resolve_receipt_pagination_paths_dynamic,
    validate_context_path,
)
from tools.receipt_query_result_tables import (  # noqa: E402
    enumerate_receipt_result_table_paths,
    find_table_paths_in_context,
)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="只读探测当前收款单查询结果页的表格/分页 path 规律"
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--start-delay", type=float, default=1.0)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    if not args.json:
        print("收款单查询结果页分页 path 只读探针")
        print("请先手工停在【收款单查询结果页】，不要停在查询条件窗口。")
        print("本探针不会写入、不会改每页 500、不会点下一页。")
        print(f"启动后等待 {args.start_delay:g} 秒。")
    time.sleep(max(args.start_delay, 0))

    config = load_config(args.config)
    report = probe_query_pagination_paths(config)
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    return 0 if report.get("ok") else 1


def probe_query_pagination_paths(config):
    query_cfg = config["receipt_entry"]["query"]
    pagination = query_cfg.get("pagination") or {}
    window_class = pagination.get("window_class", "SunAwtCanvas")
    jab = JABOperator(config)
    try:
        started = time.perf_counter()
        jab.ensure_started()
        health = check_jab_ready(jab)
        official_candidates = enumerate_receipt_result_table_paths(
            jab, query_cfg, window_class
        )
        candidates = enumerate_all_table_paths(jab, window_class)
        dynamic = resolve_receipt_pagination_paths_dynamic(jab, query_cfg)
        candidate_reports = [
            inspect_candidate(jab, query_cfg, pagination, window_class, candidate)
            for candidate in candidates[:10]
        ]
        return {
            "ok": bool(dynamic.get("ok")),
            "seconds": round(time.perf_counter() - started, 3),
            "jab_health": health,
            "suffix_contract": {
                "result_table": RECEIPT_RESULT_TABLE_PATH_SUFFIX,
                "page_label": RECEIPT_PAGE_LABEL_PATH_SUFFIX,
                "page_size_text": RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX,
                "next_page_button": RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX,
            },
            "dynamic_resolution": dynamic,
            "official_candidate_count": len(official_candidates),
            "official_candidates": official_candidates[:10],
            "candidate_count": len(candidates),
            "candidates": candidate_reports,
        }
    finally:
        jab.close()


def enumerate_all_table_paths(jab, window_class):
    if not hasattr(jab, "dll"):
        return []
    jab.ensure_started()
    if jab.dll is None:
        return []
    result = []
    table_index = 0
    for hwnd, title, class_name, pid, visible in jab.get_scoped_windows(
        include_children=True
    ):
        if class_name != window_class or not visible or not jab.dll.isJavaWindow(hwnd):
            continue
        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(root_context)
        ):
            continue
        contexts_to_release = [root_context.value]
        try:
            tables = find_table_paths_in_context(
                jab,
                vm_id.value,
                root_context.value,
                depth=0,
                index_path=[],
                owned_contexts=[],
            )
            for table in tables:
                table_info = table["table_info"]
                result.append(
                    {
                        "table_index": table_index,
                        "path": "0"
                        + "".join(f".{index}" for index in table["index_path"]),
                        "hwnd": int(hwnd),
                        "window_title": title,
                        "window_class": class_name,
                        "window_visible": visible,
                        "pid": pid,
                        "row_count": int(table_info.rowCount),
                        "col_count": int(table_info.columnCount),
                    }
                )
                table_index += 1
                contexts_to_release.extend(table["owned_contexts"])
        finally:
            jab.release_contexts(vm_id.value, list(dict.fromkeys(contexts_to_release)))
    result.sort(
        key=lambda item: (
            -int(item.get("row_count") or 0),
            -int(item.get("col_count") or 0),
            item.get("path") or "",
        )
    )
    return result


def inspect_candidate(jab, query_cfg, pagination, window_class, candidate):
    prefix = infer_result_area_prefix_from_table_path(candidate["path"])
    paths = {
        "result_table_path": candidate["path"],
        "result_area_prefix": prefix,
        "page_label_path": join_context_path(prefix, RECEIPT_PAGE_LABEL_PATH_SUFFIX)
        if prefix
        else None,
        "page_size_text_path": join_context_path(
            prefix, RECEIPT_PAGE_SIZE_TEXT_PATH_SUFFIX
        )
        if prefix
        else None,
        "next_page_button_path": join_context_path(
            prefix, RECEIPT_NEXT_PAGE_BUTTON_PATH_SUFFIX
        )
        if prefix
        else None,
    }
    timeout = float(pagination.get("dynamic_path_timeout", 0.2))
    validations = {}
    if paths["page_label_path"]:
        validations["page_label"] = validate_context_path(
            jab,
            paths["page_label_path"],
            window_class,
            role="label",
            scope_hwnd=candidate["hwnd"],
            timeout=timeout,
        )
        validations["page_size"] = validate_context_path(
            jab,
            paths["page_size_text_path"],
            window_class,
            role="text",
            scope_hwnd=candidate["hwnd"],
            timeout=timeout,
        )
        validations["next_page"] = validate_context_path(
            jab,
            paths["next_page_button_path"],
            window_class,
            role="push button",
            scope_hwnd=candidate["hwnd"],
            timeout=timeout,
        )
    page_label_text = (
        read_page_label(
            jab,
            paths["page_label_path"],
            window_class,
            candidate["hwnd"],
        )
        if paths["page_label_path"]
        else None
    )
    page_size_text = (
        read_page_size_text(
            jab,
            paths["page_size_text_path"],
            window_class,
            candidate["hwnd"],
        )
        if paths["page_size_text_path"]
        else None
    )
    return {
        **candidate,
        **paths,
        "validations": validations,
        "page_label_text": page_label_text,
        "page_size_text": page_size_text,
    }


if __name__ == "__main__":
    raise SystemExit(main())
