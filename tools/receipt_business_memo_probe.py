import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    extract_receipt_header_dynamic_index,
    find_receipt_header_field_by_dynamic_path,
    find_receipt_header_field_by_semantic_label,
    find_header_label_context_with_window,
    receipt_header_dynamic_prefix,
    split_header_path,
)


def _suffix_from_path(path, dynamic_index):
    # 原 infer_receipt_extra_text_field_suffix(已删)的等价内联:从动态 prefix 后截 suffix
    prefix = f"{receipt_header_dynamic_prefix(dynamic_index)}."
    if not path or dynamic_index is None or not str(path).startswith(prefix):
        return None
    return str(path)[len(prefix):]


def main():
    parser = argparse.ArgumentParser(description="只读探测收款单商务领款备忘附近控件")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--label", default="商务领款备忘")
    parser.add_argument("--title", default=None)
    parser.add_argument("--class-name", default="SunAwtCanvas")
    parser.add_argument("--hwnd", type=int, default=None)
    parser.add_argument("--max-vertical-distance", type=int, default=32)
    parser.add_argument("--max-right-distance", type=int, default=520)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        report = {
            "label": args.label,
            "windows": jab.describe_controls_near_label(
                args.label,
                title=args.title,
                class_name=args.class_name,
                hwnd=args.hwnd,
                require_showing=True,
                max_vertical_distance=args.max_vertical_distance,
                max_right_distance=args.max_right_distance,
            ),
        }
        report["path_probe"] = probe_path(jab, args)
        report["sibling_probe"] = probe_label_sibling_controls(jab, args)
    finally:
        jab.close()

    output_path = logs_dir() / f"receipt_business_memo_probe_{time.strftime('%Y%m%d_%H%M%S')}.json"
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"探测详情: {output_path}")
        print_text_summary(report)


def print_text_summary(report):
    path_probe = report.get("path_probe") or {}
    if path_probe:
        print(
            "path 探测: "
            f"ok={path_probe.get('ok')} source={path_probe.get('source')} "
            f"path={path_probe.get('path')} suffix={path_probe.get('suffix')}"
        )
        if path_probe.get("dynamic_prefix"):
            print(f"动态前缀: {path_probe.get('dynamic_prefix')}")
    label_count = 0
    for window in report.get("windows") or []:
        labels = window.get("labels") or []
        if not labels:
            continue
        print(
            f"窗口 hwnd={window.get('hwnd')} title={window.get('title')!r} "
            f"class={window.get('class')!r} labels={len(labels)}"
        )
        label_count += len(labels)
        for label in labels:
            label_info = label.get("label") or {}
            print(
                "  label "
                f"bounds={label_info.get('x')},{label_info.get('y')},"
                f"{label_info.get('width')},{label_info.get('height')}"
            )
            for nearby in label.get("nearby") or []:
                control = nearby.get("control") or {}
                print(
                    "  nearby "
                    f"role={control.get('role')!r} name={control.get('name')!r} "
                    f"states={control.get('states')!r} "
                    f"bounds={control.get('x')},{control.get('y')},"
                    f"{control.get('width')},{control.get('height')} "
                    f"dy={nearby.get('dy')} right={nearby.get('right_distance')} "
                    f"actions={nearby.get('actions')}"
                )
    if label_count == 0:
        print(f"未找到 label: {report.get('label')!r}")


def probe_path(jab, args):
    semantic = find_receipt_header_field_by_semantic_label(
        jab,
        args.label,
        scope_hwnd=args.hwnd,
        timeout=1.2,
    )
    if not semantic.get("ok"):
        return {
            "ok": False,
            "source": "semantic",
            "reason": semantic.get("reason"),
            "semantic_attempt": redact_contexts(semantic),
        }
    path = semantic.get("path")
    dynamic_index = extract_receipt_header_dynamic_index(path)
    suffix = _suffix_from_path(path, dynamic_index)
    dynamic = (
        find_receipt_header_field_by_dynamic_path(
            jab,
            args.label,
            dynamic_index,
            scope_hwnd=(semantic.get("window") or {}).get("hwnd") or args.hwnd,
            require_showing=True,
            require_valid_bounds=False,
        )
        if dynamic_index is not None
        else {"ok": False, "reason": "dynamic index missing"}
    )
    release_found(jab, dynamic)
    release_found(jab, semantic)
    return {
        "ok": bool(dynamic.get("ok")),
        "source": "dynamic-path" if dynamic.get("ok") else "semantic-only",
        "label": args.label,
        "path": path,
        "suffix": suffix,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": (
            receipt_header_dynamic_prefix(dynamic_index)
            if dynamic_index is not None
            else None
        ),
        "dynamic_path_attempt": redact_contexts(dynamic),
    }


def probe_label_sibling_controls(jab, args):
    label_found = find_header_label_context_with_window(
        jab,
        args.label,
        timeout=1.2,
        require_showing=True,
        scope_hwnd=args.hwnd,
    )
    label_context, vm_id, owned_contexts, owned_indexes, window = label_found
    if not label_context:
        return {
            "ok": False,
            "label": args.label,
            "reason": "semantic label not found",
        }
    label_path = "0" + "".join(f".{index}" for index in owned_indexes)
    label_parts = split_header_path(label_path)
    jab.release_contexts(vm_id, owned_contexts)
    if len(label_parts) < 2:
        return {
            "ok": False,
            "label": args.label,
            "label_path": label_path,
            "reason": "label path missing parent/index",
        }

    parent_path = ".".join(str(part) for part in label_parts[:-1])
    label_index = label_parts[-1]
    dynamic_index = extract_receipt_header_dynamic_index(label_path)
    scope_hwnd = (window or {}).get("hwnd") or args.hwnd
    candidates = []
    for index in range(label_index + 1, label_index + 9):
        path = f"{parent_path}.{index}"
        item = inspect_path(jab, path, scope_hwnd=scope_hwnd, max_child_depth=1)
        if not item.get("ok"):
            candidates.append(
                {
                    "path": path,
                    "ok": False,
                    "reason": item.get("reason"),
                    "index": index,
                }
            )
            continue
        item["index"] = index
        item["suffix"] = _suffix_from_path(path, dynamic_index)
        candidates.append(item)

    visible_right_controls = [
        item
        for item in candidates
        if item.get("ok")
        and control_is_showing(item.get("control"))
        and (item.get("control") or {}).get("role") in {"combo box", "panel", "text"}
    ]
    likely = None
    for item in visible_right_controls:
        if (item.get("control") or {}).get("role") == "combo box":
            likely = item
            break
    if likely is None and visible_right_controls:
        likely = visible_right_controls[0]

    return {
        "ok": bool(likely),
        "label": args.label,
        "label_path": label_path,
        "parent_path": parent_path,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": (
            receipt_header_dynamic_prefix(dynamic_index)
            if dynamic_index is not None
            else None
        ),
        "window": window,
        "likely_control": likely,
        "candidates": candidates,
    }


def inspect_path(jab, path, scope_hwnd=None, max_child_depth=0):
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {"ok": False, "path": path, "reason": "path not found"}
    try:
        info = jab.get_context_info(vm_id, context)
        if not info:
            return {"ok": False, "path": path, "reason": "context info not readable"}
        return {
            "ok": True,
            "path": path,
            "control": jab.info_to_dict(info),
            "actions": jab.get_action_names(vm_id, context),
            "children": inspect_children(
                jab,
                vm_id,
                context,
                path,
                depth=0,
                max_child_depth=max_child_depth,
            ),
            "window": window_info,
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def inspect_children(jab, vm_id, context, path, depth, max_child_depth):
    if depth >= max_child_depth:
        return []
    info = jab.get_context_info(vm_id, context)
    if not info:
        return []
    children = []
    owned = []
    try:
        for index in range(min(info.childrenCount, jab.max_children)):
            child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
            if not child:
                continue
            owned.append(child)
            child_info = jab.get_context_info(vm_id, child)
            if not child_info:
                continue
            child_path = f"{path}.{index}"
            children.append(
                {
                    "path": child_path,
                    "control": jab.info_to_dict(child_info),
                    "actions": jab.get_action_names(vm_id, child),
                    "children": inspect_children(
                        jab,
                        vm_id,
                        child,
                        child_path,
                        depth + 1,
                        max_child_depth,
                    ),
                }
            )
    finally:
        jab.release_contexts(vm_id, owned)
    return children


def control_is_showing(control):
    states = str((control or {}).get("states") or "").lower()
    return "visible" in states and "showing" in states


def release_found(jab, found):
    if found.get("context") and found.get("vm_id") is not None:
        jab.release_contexts(found["vm_id"], found.get("owned_contexts") or [])


def redact_contexts(value):
    if isinstance(value, dict):
        return {
            key: redact_contexts(item)
            for key, item in value.items()
            if key not in {"context", "owned_contexts", "vm_id"}
        }
    if isinstance(value, list):
        return [redact_contexts(item) for item in value]
    return value


if __name__ == "__main__":
    main()
