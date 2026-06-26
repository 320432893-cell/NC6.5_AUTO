import argparse
import json
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.paths import logs_dir  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.archive.probe_receipt_counterparty_popup_tree import (  # noqa: E402
    resolve_current_scope,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    find_finance_org_header_scope_by_shallow_semantic,
    find_label_following_text_from_container,
    find_receipt_header_field_by_live_semantic,
    header_container_candidates_from_anchor,
    split_header_path,
)


DEFAULT_LABELS = ["客户", "单据日期", "币种", "结算方式", "备注", "商务领款备忘"]


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Read-only probe for receipt header containers derived from 财务组织 label_path."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument("--labels", default=",".join(DEFAULT_LABELS))
    parser.add_argument("--max-depth", type=int, default=8)
    parser.add_argument("--max-nodes", type=int, default=220)
    parser.add_argument("--json", action="store_true")
    return parser.parse_args(argv)


def main(argv=None):
    stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
    if callable(stdout_reconfigure):
        stdout_reconfigure(encoding="utf-8", errors="replace")
    args = parse_args(argv)
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    report = {
        "ok": False,
        "read_only": True,
        "requested": {
            "labels": parse_labels(args.labels),
            "max_depth": int(args.max_depth),
            "max_nodes": int(args.max_nodes),
        },
        "scope": None,
        "finance_scope": None,
        "anchor_paths": [],
        "containers": [],
        "live_semantic": [],
    }
    try:
        jab.ensure_started()
        scope = resolve_current_scope(jab, args.scope_hwnd)
        report["scope"] = scope
        if not scope.get("ok"):
            report["reason"] = scope.get("reason") or "receipt scope not found"
            return finish(report, args)

        scope_hwnd = scope.get("hwnd")
        dynamic_index = scope.get("dynamic_index")
        started = time.perf_counter()
        finance_scope = find_finance_org_header_scope_by_shallow_semantic(
            jab,
            scope_hwnd,
            order="dfs",
            max_depth=30,
            max_nodes=800,
        )
        finance_scope["probe_seconds"] = round(time.perf_counter() - started, 4)
        report["finance_scope"] = strip_handles(finance_scope)
        anchor_paths = unique_paths(
            (finance_scope or {}).get("semantic_label_path"),
            (finance_scope or {}).get("label_path"),
        )
        report["anchor_paths"] = anchor_paths
        if not anchor_paths:
            report["reason"] = "finance org label path missing"
            return finish(report, args)

        for anchor_path in anchor_paths:
            for container_path in header_container_candidates_from_anchor(anchor_path):
                if any(
                    item.get("container_path") == container_path
                    for item in report["containers"]
                ):
                    continue
                report["containers"].append(
                    inspect_container(
                        jab,
                        container_path,
                        dynamic_index,
                        scope_hwnd,
                        parse_labels(args.labels),
                        max_depth=int(args.max_depth),
                        max_nodes=int(args.max_nodes),
                    )
                )

        for label in parse_labels(args.labels):
            report["live_semantic"].append(
                time_call(
                    "live-semantic",
                    lambda label=label: find_receipt_header_field_by_live_semantic(
                        jab,
                        label,
                        scope_hwnd=scope_hwnd,
                    ),
                )
            )
            release_result_contexts(jab, report["live_semantic"][-1].get("result"))

        report["ok"] = True
        report["diagnosis"] = diagnose(report)
    finally:
        jab.close()
    return finish(report, args)


def inspect_container(
    jab,
    container_path,
    dynamic_index,
    scope_hwnd,
    labels,
    max_depth=8,
    max_nodes=220,
):
    item = {
        "container_path": container_path,
        "ok": False,
        "summary": None,
        "label_following": [],
    }
    started = time.perf_counter()
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        container_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    item["resolve_seconds"] = round(time.perf_counter() - started, 4)
    if not context:
        item["reason"] = "container path not found"
        return item
    try:
        item["ok"] = True
        item["window"] = strip_handles(window_info)
        item["summary"] = summarize_container_tree(
            jab,
            vm_id,
            context,
            container_path,
            max_depth=max_depth,
            max_nodes=max_nodes,
        )
        for label in labels:
            item["label_following"].append(
                time_call(
                    f"label-following:{label}",
                    lambda label=label: find_label_following_text_from_container(
                        jab,
                        label,
                        container_path,
                        dynamic_index,
                        scope_hwnd=scope_hwnd,
                    ),
                )
            )
            release_result_contexts(jab, item["label_following"][-1].get("result"))
    finally:
        jab.release_contexts(vm_id, owned_contexts)
    return item


def summarize_container_tree(jab, vm_id, root_context, root_path, max_depth=8, max_nodes=220):
    started = time.perf_counter()
    stack = [(root_context, root_path, 0)]
    scanned = 0
    role_counts = Counter()
    interesting = []
    truncated = False
    while stack:
        context, path, depth = stack.pop()
        if scanned >= max_nodes:
            truncated = True
            break
        info = jab.get_context_info(vm_id, context)
        if not info:
            continue
        scanned += 1
        role = role_of(info)
        states = states_of(info)
        role_counts[role] += 1
        name = text_of(getattr(info, "name", ""))
        desc = text_of(getattr(info, "description", ""))
        if role in {"label", "text", "combo box"} and (name or desc or role != "label"):
            interesting.append(
                {
                    "path": path,
                    "depth": depth,
                    "role": role,
                    "name": name,
                    "description": desc,
                    "states": states,
                    "bounds": bounds_of(info),
                    "children": int(getattr(info, "childrenCount", 0) or 0),
                }
            )
        if depth >= max_depth or role == "table":
            continue
        child_count = min(int(getattr(info, "childrenCount", 0) or 0), jab.max_children)
        children = []
        for index in range(child_count):
            child = jab.dll.getAccessibleChildFromContext(vm_id, context, index)
            if child:
                children.append((child, f"{path}.{index}", depth + 1))
        stack.extend(reversed(children))
    return {
        "seconds": round(time.perf_counter() - started, 4),
        "scanned_nodes": scanned,
        "truncated": truncated,
        "role_counts": dict(sorted(role_counts.items())),
        "interesting_count": len(interesting),
        "interesting": interesting[:80],
    }


def time_call(method, func):
    started = time.perf_counter()
    try:
        result = func()
        return {
            "method": method,
            "ok": is_ok(result),
            "seconds": round(time.perf_counter() - started, 4),
            "result": summarize_result(result),
        }
    except Exception as exc:
        return {
            "method": method,
            "ok": False,
            "seconds": round(time.perf_counter() - started, 4),
            "error": f"{type(exc).__name__}: {exc}",
        }


def diagnose(report):
    containers = report.get("containers") or []
    diagnosis = {
        "container_count": len(containers),
        "containers_with_customer_label": [],
        "containers_with_customer_following_text": [],
    }
    for item in containers:
        summary = item.get("summary") or {}
        customer_hits = [
            node
            for node in summary.get("interesting") or []
            if node.get("role") == "label"
            and (node.get("name") == "客户" or node.get("description") == "客户")
        ]
        if customer_hits:
            diagnosis["containers_with_customer_label"].append(
                {
                    "container_path": item.get("container_path"),
                    "hits": customer_hits[:5],
                }
            )
        following_hits = [
            entry
            for entry in item.get("label_following") or []
            if entry.get("ok")
            and str(entry.get("method") or "").endswith(":客户")
        ]
        if following_hits:
            diagnosis["containers_with_customer_following_text"].append(
                {
                    "container_path": item.get("container_path"),
                    "hits": following_hits,
                }
            )
    if diagnosis["containers_with_customer_following_text"]:
        diagnosis["recommended"] = "finance-org-container route can be strengthened from matching container"
    elif diagnosis["containers_with_customer_label"]:
        diagnosis["recommended"] = "customer label exists but following-text matcher missed; inspect sibling order"
    else:
        diagnosis["recommended"] = "finance-org ancestor containers do not include customer label; need broader same-scope fast scan"
    return diagnosis


def finish(report, args):
    output_path = (
        logs_dir()
        / f"receipt_header_container_tree_{time.strftime('%Y%m%d_%H%M%S')}.json"
    )
    report["output_path"] = str(output_path)
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"探测详情: {output_path}")
        print(f"ok={report.get('ok')} reason={report.get('reason')}")
        diagnosis = report.get("diagnosis") or {}
        if diagnosis:
            print(f"诊断: {diagnosis.get('recommended')}")
    return 0 if report.get("ok") else 1


def parse_labels(text):
    return [part.strip() for part in str(text or "").replace("，", ",").split(",") if part.strip()]


def unique_paths(*values):
    result = []
    for value in values:
        text = str(value or "").strip()
        if text and text not in result:
            result.append(text)
    return result


def is_ok(result):
    return bool(isinstance(result, dict) and result.get("ok"))


def summarize_result(result):
    if not isinstance(result, dict):
        return result
    return strip_handles(
        {
            key: value
            for key, value in result.items()
            if key
            in {
                "ok",
                "reason",
                "source",
                "path",
                "label_path",
                "container_path",
                "dynamic_index",
                "dynamic_prefix",
                "scope_hwnd",
                "window",
                "scanned_nodes",
                "anchor_paths",
                "attempts",
            }
        }
    )


def release_result_contexts(jab, result):
    if isinstance(result, dict) and result.get("vm_id") and result.get("owned_contexts"):
        try:
            jab.release_contexts(result["vm_id"], result["owned_contexts"])
        except Exception:
            pass


def strip_handles(value):
    if isinstance(value, dict):
        return {
            str(key): strip_handles(item)
            for key, item in value.items()
            if key not in {"context", "owned_contexts", "vm_id"}
        }
    if isinstance(value, (list, tuple)):
        return [strip_handles(item) for item in value[:100]]
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def role_of(info):
    return text_of(getattr(info, "role_en_US", "") or getattr(info, "role", "")).lower()


def states_of(info):
    return text_of(getattr(info, "states_en_US", "") or getattr(info, "states", ""))


def text_of(value):
    return str(value or "").strip()


def bounds_of(info):
    return [
        int(getattr(info, "x", -1)),
        int(getattr(info, "y", -1)),
        int(getattr(info, "width", -1)),
        int(getattr(info, "height", -1)),
    ]


if __name__ == "__main__":
    raise SystemExit(main())
