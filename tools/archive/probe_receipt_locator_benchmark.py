import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.jab_near_label import find_text_context_near_label_once  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.archive.probe_receipt_counterparty_popup_tree import (  # noqa: E402
    resolve_current_scope,
)
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_full_flow_entry import (  # noqa: E402
    build_body_table_cached_path,
    find_counterparty_combo,
    find_counterparty_combo_nearby,
    resolve_body_table_by_dynamic_prefix,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    HEADER_FORM_TEXT_INDEXES,
    HEADER_SCOPE_ANCHOR_LABEL,
    find_label_following_text,
    find_finance_org_header_scope_by_shallow_semantic,
    find_receipt_header_field_by_dynamic_path,
    find_receipt_header_field_by_live_semantic,
    find_receipt_header_field_by_scoped_label,
    find_receipt_header_field_by_semantic_label,
    infer_header_path_template_from_field,
)


DEFAULT_LABELS = [
    "财务组织",
    "客户",
    "单据日期",
    "币种",
    "结算方式",
    "商务领款备忘",
    "备注",
    "往来对象",
]
EXTRA_LABELS = {"商务领款备忘", "备注"} & set(HEADER_FORM_TEXT_INDEXES)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Read-only benchmark for NC receipt locator fallback strategies."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument(
        "--labels",
        default=",".join(DEFAULT_LABELS),
        help="Comma-separated header labels to benchmark.",
    )
    parser.add_argument("--repeat", type=int, default=1)
    parser.add_argument("--semantic-timeout", type=float, default=0.35)
    parser.add_argument(
        "--include-body",
        action="store_true",
        help="Also benchmark receipt body table cached path and fallback scan.",
    )
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
            "repeat": max(int(args.repeat or 1), 1),
            "semantic_timeout": float(args.semantic_timeout),
            "include_body": bool(args.include_body),
        },
        "scope": None,
        "benchmarks": [],
        "body_table": [],
    }
    try:
        jab.ensure_started()
        scope = resolve_current_scope(jab, args.scope_hwnd)
        report["scope"] = scope
        if not scope.get("ok"):
            report["reason"] = scope.get("reason") or "receipt scope not found"
            return finish(report, args)

        dynamic_index = scope.get("dynamic_index")
        scope_hwnd = scope.get("hwnd")
        finance_scope = find_finance_org_header_scope_by_shallow_semantic(
            jab,
            scope_hwnd,
            order="dfs",
            max_depth=30,
            max_nodes=800,
        )
        report["finance_scope"] = summarize_result(finance_scope)
        for label in parse_labels(args.labels):
            report["benchmarks"].append(
                benchmark_label(
                    jab,
                    label,
                    dynamic_index,
                    scope_hwnd,
                    finance_scope=finance_scope,
                    repeat=max(int(args.repeat or 1), 1),
                    semantic_timeout=float(args.semantic_timeout),
                )
            )
        if args.include_body:
            report["body_table"] = benchmark_body_table(
                jab,
                dynamic_index,
                scope_hwnd,
                repeat=max(int(args.repeat or 1), 1),
            )
        report["ok"] = True
    finally:
        jab.close()
    return finish(report, args)


def benchmark_label(
    jab,
    label,
    dynamic_index,
    scope_hwnd,
    finance_scope=None,
    repeat=1,
    semantic_timeout=0.35,
):
    item = {"label": label, "methods": []}
    if label == HEADER_SCOPE_ANCHOR_LABEL:
        item["methods"].append(
            run_repeated(
                jab,
                "finance-org-shallow-semantic",
                repeat,
                lambda: find_finance_org_header_scope_by_shallow_semantic(
                    jab,
                    scope_hwnd,
                    order="dfs",
                    max_depth=30,
                    max_nodes=800,
                ),
            )
        )
    elif label == "往来对象":
        item["methods"].append(
            run_repeated(
                jab,
                "counterparty-learned-or-nearby",
                repeat,
                lambda: find_counterparty_combo(
                    jab,
                    dynamic_index,
                    scope_hwnd=scope_hwnd,
                ),
            )
        )
        item["methods"].append(
            run_repeated(
                jab,
                "counterparty-nearby-bounds",
                repeat,
                lambda: find_counterparty_combo_nearby(
                    jab,
                    dynamic_index,
                    scope_hwnd=scope_hwnd,
                ),
            )
        )
    else:
        semantic_func = (
            find_receipt_header_field_by_live_semantic
            if label in EXTRA_LABELS
            else find_receipt_header_field_by_semantic_label
        )
        item["methods"].append(
            run_repeated(
                jab,
                "dynamic-path",
                repeat,
                lambda: find_receipt_header_field_by_dynamic_path(
                    jab,
                    label,
                    dynamic_index,
                    scope_hwnd=scope_hwnd,
                    require_showing=True,
                    require_valid_bounds=False,
                ),
            )
        )
        if label == "客户":
            item["methods"].append(
                run_repeated(
                    jab,
                    "finance-org-container-label-following-text",
                    repeat,
                    lambda: find_header_field_from_finance_org_container(
                        jab,
                        label,
                        dynamic_index,
                        scope_hwnd,
                        finance_scope=finance_scope,
                    ),
                )
            )
        item["methods"].append(
            run_repeated(
                jab,
                "semantic-label-dfs",
                repeat,
                lambda: semantic_func(
                    jab,
                    label,
                    scope_hwnd=scope_hwnd,
                    timeout=semantic_timeout,
                )
            )
        )
        item["methods"].append(
            run_repeated(
                jab,
                "scoped-label-following-text",
                repeat,
                lambda: find_receipt_header_field_by_scoped_label(
                    jab,
                    label,
                    scope_hwnd=scope_hwnd,
                ),
            )
        )
        item["methods"].append(
            run_repeated(
                jab,
                "near-label-bounds",
                repeat,
                lambda: find_text_context_near_label_once(
                    jab,
                    label,
                    class_name="SunAwtCanvas",
                    hwnd=scope_hwnd,
                    require_showing=True,
                ),
            )
        )
    return item


def find_header_field_from_finance_org_container(
    jab,
    label,
    dynamic_index,
    scope_hwnd,
    finance_scope=None,
):
    started = time.perf_counter()
    anchor_paths = []
    for key in ("semantic_label_path", "label_path"):
        value = (finance_scope or {}).get(key)
        if value and value not in anchor_paths:
            anchor_paths.append(value)
    if not anchor_paths:
        return {
            "ok": False,
            "label": label,
            "source": "finance-org-container-label-following-text",
            "reason": "finance org label path missing",
            "seconds": round(time.perf_counter() - started, 4),
        }

    attempts = []
    for anchor_path in anchor_paths:
        for container_path in header_container_candidates(anchor_path):
            if any(item.get("container_path") == container_path for item in attempts):
                continue
            attempt = find_label_following_text_from_container(
                jab,
                label,
                container_path,
                dynamic_index,
                scope_hwnd,
            )
            attempts.append(attempt)
            if attempt.get("ok"):
                template = infer_header_path_template_from_field(
                    attempt.get("path"),
                    dynamic_index,
                    label,
                )
                return {
                    **attempt,
                    "source": "finance-org-container-label-following-text",
                    "anchor_paths": anchor_paths,
                    "attempts": summarize_container_attempts(attempts),
                    "header_path_template": template,
                    "seconds": round(time.perf_counter() - started, 4),
                }
    return {
        "ok": False,
        "label": label,
        "source": "finance-org-container-label-following-text",
        "anchor_paths": anchor_paths,
        "attempts": summarize_container_attempts(attempts),
        "reason": "customer label-following text not found from finance-org containers",
        "seconds": round(time.perf_counter() - started, 4),
    }


def header_container_candidates(anchor_path):
    parts = split_path(anchor_path)
    if len(parts) < 4:
        return []
    candidates = []
    # Current observed header layouts put ordinary fields under ancestors around
    # the finance-org label path's parent/grandparent branch. Trying these few
    # ancestors keeps this read-only probe bounded and avoids whole-canvas DFS.
    for keep in (2, 3, 4, 5, 6):
        if len(parts) > keep:
            candidate = ".".join(str(part) for part in parts[:-keep])
            if candidate and candidate not in candidates:
                candidates.append(candidate)
    return candidates


def find_label_following_text_from_container(
    jab,
    label,
    container_path,
    dynamic_index,
    scope_hwnd,
):
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        container_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "label": label,
            "container_path": container_path,
            "reason": "container path not found",
        }
    before_calls = getattr(jab, "_container_probe_info_calls", 0)
    setattr(jab, "_container_probe_info_calls", 0)
    try:
        result = find_label_following_text_counted(
            jab,
            vm_id,
            context,
            label,
            container_path,
            depth=0,
            owned_contexts=list(owned_contexts or []),
        )
        scanned_nodes = getattr(jab, "_container_probe_info_calls", 0)
    finally:
        setattr(jab, "_container_probe_info_calls", before_calls)
    if not result:
        jab.release_contexts(vm_id, owned_contexts)
        return {
            "ok": False,
            "label": label,
            "container_path": container_path,
            "scanned_nodes": scanned_nodes,
            "reason": "label-following text not found in container",
        }
    text_context, owned, text_path, label_path = result
    return {
        "ok": True,
        "label": label,
        "context": text_context,
        "vm_id": vm_id,
        "owned_contexts": owned,
        "path": text_path,
        "label_path": label_path,
        "container_path": container_path,
        "dynamic_index": dynamic_index,
        "scanned_nodes": scanned_nodes,
        "window": window_info,
    }


def find_label_following_text_counted(
    jab,
    vm_id,
    context,
    label,
    path,
    depth,
    owned_contexts,
):
    original_get_context_info = jab.get_context_info

    def counted_get_context_info(vm_id_arg, context_arg):
        current = getattr(jab, "_container_probe_info_calls", 0)
        setattr(jab, "_container_probe_info_calls", current + 1)
        return original_get_context_info(vm_id_arg, context_arg)

    jab.get_context_info = counted_get_context_info
    try:
        return find_label_following_text(
            jab,
            vm_id,
            context,
            label,
            path,
            depth,
            owned_contexts,
        )
    finally:
        jab.get_context_info = original_get_context_info


def summarize_container_attempts(attempts):
    result = []
    for item in attempts or []:
        result.append(
            strip_value(
                {
                    "ok": item.get("ok"),
                    "container_path": item.get("container_path"),
                    "path": item.get("path"),
                    "label_path": item.get("label_path"),
                    "scanned_nodes": item.get("scanned_nodes"),
                    "reason": item.get("reason"),
                }
            )
        )
    return result


def split_path(path):
    try:
        return [int(part) for part in str(path).split(".") if part != ""]
    except ValueError:
        return []


def benchmark_body_table(jab, dynamic_index, scope_hwnd, repeat):
    cached = build_body_table_cached_path(dynamic_index, scope_hwnd=scope_hwnd)
    return [
        run_repeated(
            jab,
            "body-dynamic-prefix-cached-path",
            repeat,
            lambda: resolve_body_table_by_dynamic_prefix(
                jab,
                dynamic_index,
                scope_hwnd=scope_hwnd,
            ),
        ),
        run_repeated(
            jab,
            "body-semantic-table-scan",
            repeat,
            lambda: locate_receipt_body_table(
                jab,
                max_rows=5,
                scope_hwnd=scope_hwnd,
            ),
        ),
        {
            "method": "body-cached-path",
            "cached_path": strip_value(cached.get("best") if cached else None),
        },
    ]


def run_repeated(jab, method, repeat, func):
    runs = []
    for _index in range(repeat):
        started = time.perf_counter()
        result = None
        try:
            result = func()
            runs.append(
                {
                    "ok": bool(is_ok_result(result)),
                    "seconds": round(time.perf_counter() - started, 4),
                    "result": summarize_result(result),
                }
            )
        except Exception as exc:
            runs.append(
                {
                    "ok": False,
                    "seconds": round(time.perf_counter() - started, 4),
                    "error": repr(exc),
                }
            )
        finally:
            release_result_contexts(jab, result)
    seconds = [run.get("seconds", 0.0) for run in runs]
    return {
        "method": method,
        "ok": any(run.get("ok") for run in runs),
        "runs": runs,
        "min_seconds": min(seconds) if seconds else None,
        "max_seconds": max(seconds) if seconds else None,
        "avg_seconds": round(sum(seconds) / len(seconds), 4) if seconds else None,
    }


def is_ok_result(result):
    if isinstance(result, dict):
        if "best" in result:
            return bool(result.get("best"))
        return bool(result.get("ok"))
    if isinstance(result, tuple):
        return bool(result and result[0])
    return bool(result)


def summarize_result(result):
    if isinstance(result, tuple):
        context, vm_id, owned_contexts, label_info, text_info, window_info = result
        return {
            "ok": bool(context),
            "vm_id": vm_id,
            "owned_count": len(owned_contexts or []),
            "label": info_summary(label_info),
            "text": info_summary(text_info),
            "window": strip_value(window_info),
        }
    if isinstance(result, dict):
        return strip_value(
            {
                key: value
                for key, value in result.items()
                if key
                in {
                    "ok",
                    "reason",
                    "source",
                    "method",
                    "mode",
                    "path",
                    "label_path",
                    "dynamic_index",
                    "dynamic_prefix",
                    "scope_hwnd",
                    "window",
                    "cache_hit",
                    "fallback_used",
                    "best",
                    "candidate_count",
                    "scanned_nodes",
                    "path_validation",
                    "customer_index_correction",
                }
            }
        )
    return str(result)


def release_result_contexts(jab, result):
    if isinstance(result, dict) and result.get("vm_id") and result.get("owned_contexts"):
        try:
            jab.release_contexts(result["vm_id"], result["owned_contexts"])
        except Exception:
            pass
    elif isinstance(result, tuple) and len(result) >= 3:
        _context, vm_id, owned_contexts = result[:3]
        if vm_id and owned_contexts:
            try:
                jab.release_contexts(vm_id, owned_contexts)
            except Exception:
                pass


def info_summary(info):
    if not info:
        return None
    return {
        "name": str(getattr(info, "name", "") or "").strip(),
        "description": str(getattr(info, "description", "") or "").strip(),
        "role": str(
            getattr(info, "role_en_US", "") or getattr(info, "role", "") or ""
        ).strip(),
        "states": str(
            getattr(info, "states_en_US", "") or getattr(info, "states", "") or ""
        ).strip(),
        "bounds": [
            getattr(info, "x", None),
            getattr(info, "y", None),
            getattr(info, "width", None),
            getattr(info, "height", None),
        ],
    }


def strip_value(value):
    if isinstance(value, dict):
        return {
            str(key): strip_value(item)
            for key, item in value.items()
            if key not in {"context", "owned_contexts", "vm_id"}
        }
    if isinstance(value, (list, tuple)):
        return [strip_value(item) for item in value[:20]]
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def parse_labels(text):
    labels = []
    for item in str(text or "").replace("，", ",").split(","):
        label = item.strip()
        if label and label not in labels:
            labels.append(label)
    return labels


def finish(report, args):
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print_text(report)
    output_dir = Path("logs")
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / (
        f"receipt_locator_benchmark_{time.strftime('%Y%m%d_%H%M%S')}.json"
    )
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if report.get("ok") else 1


def print_text(report):
    print(f"ok={report.get('ok')} reason={report.get('reason')}")
    scope = report.get("scope") or {}
    print(
        "scope: "
        f"hwnd={scope.get('hwnd')} dynamic_index={scope.get('dynamic_index')} "
        f"mode={scope.get('mode')}"
    )
    for item in report.get("benchmarks") or []:
        print(f"\n[{item.get('label')}]")
        for method in item.get("methods") or []:
            print(
                f"  {method.get('method')}: ok={method.get('ok')} "
                f"avg={method.get('avg_seconds')} min={method.get('min_seconds')} "
                f"max={method.get('max_seconds')}"
            )
            first = (method.get("runs") or [{}])[0]
            result = first.get("result") or {}
            if result.get("path") or result.get("reason"):
                print(f"    path={result.get('path')} reason={result.get('reason')}")
    if report.get("body_table"):
        print("\n[body_table]")
        for method in report.get("body_table") or []:
            print(
                f"  {method.get('method')}: ok={method.get('ok')} "
                f"avg={method.get('avg_seconds')}"
            )


if __name__ == "__main__":
    raise SystemExit(main())
