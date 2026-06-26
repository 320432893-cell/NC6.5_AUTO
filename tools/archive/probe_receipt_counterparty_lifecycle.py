import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.jab_popup import close_popup_hwnd, collect_visible_popup_windows  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.archive.probe_receipt_counterparty_methods import (  # noqa: E402
    action_embedded_customer_option,
    activate_and_press,
    cleanup_all_visible_popups,
    cleanup_popups,
    commit_method,
    find_counterparty_in_tree,
    find_counterparty_popup,
    probe_embedded_counterparty_tree,
    request_focus,
    root_hwnd,
    run_open_method,
    select_embedded_customer_option,
    snapshot_combo_tree,
    wait_for_any_counterparty_popup,
)
from tools.archive.probe_receipt_counterparty_popup_tree import (  # noqa: E402
    find_counterparty_combo_near_label,
    resolve_current_scope,
    strip_handles,
    summarize_windows,
)
from tools.archive.probe_receipt_counterparty_sync import (  # noqa: E402
    compare_counterparty,
    read_detail_counterparty,
    read_header_counterparty,
)
from tools.receipt_full_flow_entry import (  # noqa: E402
    find_counterparty_combo,
    read_counterparty_combo_state,
)

EXPECTED = "客户"
KNOWN_OPTIONS = {"客户", "部门", "业务员", "供应商"}


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Comprehensive NC receipt counterparty lifecycle probe. Default is "
            "read-only; --commit may change 往来对象 to 客户."
        )
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument("--row", type=int, default=0)
    parser.add_argument("--col", type=int, default=0)
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--interval", type=float, default=0.25)
    parser.add_argument("--wait", type=float, default=0.5)
    parser.add_argument("--poll", type=float, default=0.08)
    parser.add_argument("--max-depth", type=int, default=8)
    parser.add_argument("--max-children", type=int, default=160)
    parser.add_argument(
        "--methods",
        default=(
            "snapshot,toggle,esc-toggle,double-toggle,focus-alt-down,focus-f4,"
            "focus-space,focus-enter,embedded-select-enter,"
            "embedded-action-enter,embedded-select-activate-enter,"
            "activate-home-enter"
        ),
        help=(
            "Comma-separated lifecycle methods. snapshot is read-only; other "
            "methods may focus/open dropdown. With --commit, commit-capable "
            "methods may change 往来对象."
        ),
    )
    parser.add_argument(
        "--commit",
        action="store_true",
        help="Allow commit probes that may set 往来对象 to 客户.",
    )
    parser.add_argument(
        "--cleanup-popups",
        action="store_true",
        help="Close visible SunAwtWindow popups before/after each method.",
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
    jab.hide_blank_awt_windows_enabled = False
    report = {
        "ok": False,
        "read_only": not bool(args.commit),
        "commit": bool(args.commit),
        "requested": {
            "row": args.row,
            "col": args.col,
            "repeat": args.repeat,
            "methods": parse_methods(args.methods),
        },
        "scope": None,
        "target": None,
        "snapshots": [],
        "lifecycle": [],
        "diagnosis": None,
    }
    try:
        jab.ensure_started()
        if args.cleanup_popups:
            report["initial_cleanup"] = cleanup_all_visible_popups(jab)
        scope = resolve_current_scope(jab, args.scope_hwnd)
        report["scope"] = scope
        if not scope.get("ok"):
            report["reason"] = scope.get("reason") or "receipt header scope not found"
            return finish(report, args)

        dynamic_index = scope.get("dynamic_index")
        scope_hwnd = scope.get("hwnd")
        found = locate_counterparty_combo(jab, dynamic_index, scope_hwnd)
        report["target"] = strip_handles(found)
        if not found.get("ok"):
            report["reason"] = found.get("reason") or "counterparty combo not found"
            return finish(report, args)

        try:
            for index in range(max(int(args.repeat or 1), 1)):
                report["snapshots"].append(
                    collect_snapshot(
                        jab,
                        found,
                        dynamic_index,
                        scope_hwnd,
                        args,
                        label=f"snapshot-{index}",
                    )
                )
                time.sleep(max(float(args.interval or 0), 0.0))

            for method in parse_methods(args.methods):
                item = probe_lifecycle_method(
                    jab,
                    found,
                    dynamic_index,
                    scope_hwnd,
                    method,
                    args,
                )
                report["lifecycle"].append(item)
                if args.cleanup_popups:
                    item["cleanup_after_method"] = cleanup_all_visible_popups(jab)
                time.sleep(max(float(args.interval or 0), 0.0))
        finally:
            jab.release_contexts(found["vm_id"], found["owned_contexts"])

        report["diagnosis"] = diagnose_report(report)
        report["ok"] = bool(report["diagnosis"].get("usable_read"))
        if not report["ok"]:
            report["reason"] = report["diagnosis"].get("reason")
    finally:
        jab.close()
    return finish(report, args)


def locate_counterparty_combo(jab, dynamic_index, scope_hwnd):
    path_found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
    if path_found.get("ok"):
        path_found["find_method"] = path_found.get("source") or "formal-finder"
        return path_found
    near = find_counterparty_combo_near_label(jab, scope_hwnd=scope_hwnd)
    if near.get("ok"):
        near["find_method"] = "near-label"
        near["path_found"] = strip_handles(path_found)
        return near
    return {
        "ok": False,
        "reason": near.get("reason") or path_found.get("reason"),
        "path_found": strip_handles(path_found),
        "near_label_found": strip_handles(near),
    }


def collect_snapshot(jab, found, dynamic_index, scope_hwnd, args, label):
    vm_id = found["vm_id"]
    context = found["context"]
    visible_popups = collect_visible_popup_windows(
        jab, max_depth=args.max_depth, max_children=args.max_children
    )
    header = read_header_counterparty(jab, dynamic_index, scope_hwnd)
    detail = read_detail_counterparty(
        jab,
        row=args.row,
        col=args.col,
        max_rows=3,
        scope_hwnd=scope_hwnd,
    )
    state = read_counterparty_combo_state(jab, vm_id, context)
    tree = snapshot_combo_tree(jab, vm_id, context, depth=6)
    embedded = probe_embedded_counterparty_tree(jab, vm_id, context)
    comparison = compare_counterparty(header, detail)
    return {
        "label": label,
        "at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "header": header,
        "detail": detail,
        "comparison": comparison,
        "combo_state": state,
        "combo_tree_summary": summarize_combo_tree(tree),
        "embedded": embedded,
        "visible_popups": summarize_windows(
            visible_popups, args.max_depth, args.max_children
        ),
        "lifecycle": diagnose_snapshot_lifecycle(state, tree, embedded, visible_popups),
    }


def probe_lifecycle_method(jab, found, dynamic_index, scope_hwnd, method, args):
    vm_id = found["vm_id"]
    context = found["context"]
    window_hwnd = (found.get("window") or {}).get("hwnd")
    window_root = root_hwnd(window_hwnd) or window_hwnd
    before = collect_snapshot(
        jab,
        found,
        dynamic_index,
        scope_hwnd,
        args,
        label=f"before-{method}",
    )
    before_popups = collect_visible_popup_windows(
        jab, max_depth=args.max_depth, max_children=args.max_children
    )
    action = run_open_method(
        jab, vm_id, context, method, window_root=window_root
    )
    popup_probe = wait_for_any_counterparty_popup(
        jab,
        before_popups,
        timeout=args.wait,
        interval=args.poll,
    )
    after_open = collect_snapshot(
        jab,
        found,
        dynamic_index,
        scope_hwnd,
        args,
        label=f"after-open-{method}",
    )
    commit = None
    if args.commit:
        commit = run_commit_probe(jab, vm_id, context, method, window_root)
        time.sleep(0.15)
    after_commit = collect_snapshot(
        jab,
        found,
        dynamic_index,
        scope_hwnd,
        args,
        label=f"after-commit-{method}",
    )
    popup_after = collect_visible_popup_windows(
        jab, max_depth=args.max_depth, max_children=args.max_children
    )
    return {
        "method": method,
        "allowed_to_commit": bool(args.commit),
        "before": before,
        "action": action,
        "popup_probe": popup_probe,
        "counterparty_popup": find_counterparty_popup(popup_after),
        "after_open": after_open,
        "commit": commit,
        "after_commit": after_commit,
        "method_diagnosis": diagnose_method(before, action, popup_probe, after_open, commit, after_commit),
        "popup_cleanup": cleanup_popups(popup_after) if args.cleanup_popups else None,
    }


def run_commit_probe(jab, vm_id, context, method, window_root):
    if method == "embedded-select-enter":
        return select_embedded_customer_option(jab, vm_id, context, press_enter=True)
    if method == "embedded-action-enter":
        return action_embedded_customer_option(jab, vm_id, context, press_enter=True)
    if method == "embedded-select-activate-enter":
        selected = select_embedded_customer_option(jab, vm_id, context, press_enter=False)
        pressed = activate_and_press(jab, window_root, ["enter"])
        return {
            "ok": bool(selected.get("ok") and pressed.get("ok")),
            "method": method,
            "select": selected,
            "activate_press": pressed,
        }
    return commit_method(jab, vm_id, context, method, window_root=window_root)


def summarize_combo_tree(tree):
    root = (tree or {}).get("root") or {}
    labels = []
    collect_labels(root, labels)
    return {
        "root": {
            key: root.get(key)
            for key in (
                "role",
                "name",
                "description",
                "states",
                "bounds",
                "children_count",
                "accessible_action",
                "accessible_selection",
                "accessible_text",
                "actions",
                "text",
                "selection",
                "selected_child_indexes",
            )
        },
        "labels": labels,
        "counterparty_labels": [
            item for item in labels if item.get("name") in KNOWN_OPTIONS
        ],
    }


def collect_labels(node, labels):
    if not isinstance(node, dict):
        return
    if str(node.get("role") or "").strip().lower() == "label":
        labels.append(
            {
                "path": node.get("path"),
                "name": node.get("name"),
                "description": node.get("description"),
                "states": node.get("states"),
                "bounds": node.get("bounds"),
            }
        )
    for child in node.get("children") or []:
        collect_labels(child, labels)


def diagnose_snapshot_lifecycle(state, tree, embedded, popups):
    states = str((state or {}).get("states") or "").lower()
    root = ((tree or {}).get("root") or {})
    text = normalized_counterparty_text(
        (embedded or {}).get("selected_labels"),
        (state or {}).get("description"),
        (state or {}).get("text"),
        (state or {}).get("name"),
    )
    visible_counterparty_popups = [
        popup for popup in popups or [] if find_counterparty_popup([popup])
    ]
    issues = []
    if "expanded" in states and not visible_counterparty_popups:
        issues.append("combo-expanded-without-visible-popup")
    bounds = root.get("bounds") or []
    if len(bounds) == 4 and (bounds[2] <= 0 or bounds[3] <= 0):
        issues.append("combo-invalid-bounds")
    if not text:
        issues.append("counterparty-read-empty-or-garbage")
    embedded_labels = (embedded or {}).get("labels") or []
    if EXPECTED in embedded_labels and not (embedded or {}).get("selected_labels"):
        issues.append("embedded-options-present-no-selected-label")
    return {
        "value": text,
        "issues": issues,
        "expanded": "expanded" in states,
        "collapsed": "collapsed" in states,
        "visible_counterparty_popup_count": len(visible_counterparty_popups),
    }


def diagnose_method(before, action, popup_probe, after_open, commit, after_commit):
    before_life = (before or {}).get("lifecycle") or {}
    after_life = (after_open or {}).get("lifecycle") or {}
    after_commit_life = (after_commit or {}).get("lifecycle") or {}
    return {
        "opened_detectable_popup": bool((popup_probe or {}).get("counterparty_popup")),
        "action_ok": bool((action or {}).get("ok")),
        "before_value": before_life.get("value"),
        "after_open_value": after_life.get("value"),
        "after_commit_value": after_commit_life.get("value"),
        "after_open_issues": after_life.get("issues") or [],
        "after_commit_issues": after_commit_life.get("issues") or [],
        "commit_ok": None if commit is None else bool(commit.get("ok")),
        "commit_changed_to_customer": after_commit_life.get("value") == EXPECTED,
    }


def diagnose_report(report):
    snapshots = report.get("snapshots") or []
    values = [
        ((item.get("lifecycle") or {}).get("value") or "")
        for item in snapshots
    ]
    issues = []
    for item in snapshots:
        issues.extend((item.get("lifecycle") or {}).get("issues") or [])
    method_diagnoses = [
        item.get("method_diagnosis") or {} for item in report.get("lifecycle") or []
    ]
    commit_successes = [
        item
        for item in method_diagnoses
        if item.get("commit_ok") and item.get("commit_changed_to_customer")
    ]
    usable_read = any(value == EXPECTED for value in values)
    return {
        "usable_read": usable_read,
        "values": values,
        "issue_counts": count_items(issues),
        "stable_customer_read": bool(values and all(value == EXPECTED for value in values)),
        "empty_or_garbage_reads": sum(1 for value in values if not value),
        "commit_success_methods": commit_successes,
        "reason": None
        if usable_read
        else "counterparty customer value not reliably readable",
    }


def normalized_counterparty_text(*values):
    for value in values:
        if isinstance(value, list):
            for item in value:
                text = normalize_text(item)
                if text:
                    return text
            continue
        text = normalize_text(value)
        if text:
            return text
    return ""


def normalize_text(value):
    text = str(value or "").strip()
    if text in {"翸", "ɲ", "ǁ", "|", "||"}:
        return ""
    return text


def count_items(items):
    result = {}
    for item in items or []:
        result[item] = result.get(item, 0) + 1
    return dict(sorted(result.items()))


def parse_methods(text):
    return [item.strip() for item in str(text or "").split(",") if item.strip()]


def finish(report, args):
    output_dir = ROOT / "logs"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / (
        f"counterparty_lifecycle_probe_{time.strftime('%Y%m%d_%H%M%S')}.json"
    )
    output_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    report["output_path"] = str(output_path)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json:
        print(text)
    else:
        print(f"探测结果: {output_path}")
        print(text)
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
