import argparse
import ctypes
import json
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.jab_popup import (  # noqa: E402
    collect_new_visible_popup_windows,
    collect_visible_popup_windows,
    close_popup_hwnd,
)
from core.jab_near_label import collect_controls_for_bounds_scan, info_to_dict  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_probe import JOBJECT, enum_windows  # noqa: E402
from tools.receipt_full_flow_entry import (  # noqa: E402
    find_counterparty_combo,
    read_counterparty_combo_state,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    resolve_receipt_header_anchor_in_canvas,
)

COUNTERPARTY_LABEL = "往来对象"


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Read-only probe for the receipt header counterparty popup tree."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument(
        "--popup-only",
        action="store_true",
        help="Only dump currently visible popup windows; open the dropdown manually first.",
    )
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument("--wait", type=float, default=0.8)
    parser.add_argument("--poll", type=float, default=0.08)
    parser.add_argument("--max-depth", type=int, default=8)
    parser.add_argument("--max-children", type=int, default=120)
    parser.add_argument("--json", action="store_true")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        if args.popup_only:
            popups = collect_visible_popup_windows(jab)
            report = {
                "ok": bool(popups),
                "mode": "popup-only",
                "popup_count": len(popups),
                "popup_windows": summarize_windows(
                    popups, args.max_depth, args.max_children
                ),
                "reason": None if popups else "no visible popup windows found",
            }
            return finish(report, args)
        scope = resolve_current_scope(jab, args.scope_hwnd)
        if not scope.get("ok"):
            report = {
                "ok": False,
                "reason": scope.get("reason") or "unable to resolve receipt header scope",
                "scope": scope,
            }
            return finish(report, args)

        dynamic_index = scope.get("dynamic_index")
        scope_hwnd = scope.get("hwnd")
        path_found = None
        found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
        near_label_found = None
        if not found.get("ok"):
            near_label_found = find_counterparty_combo_near_label(
                jab,
                scope_hwnd=scope_hwnd,
                label=COUNTERPARTY_LABEL,
            )
            found = near_label_found
        if not found.get("ok"):
            report = {
                "ok": False,
                "reason": found.get("reason") or "counterparty combo not found",
                "scope": scope,
                "path": None,
                "path_found": strip_handles(path_found),
                "near_label_found": strip_handles(near_label_found),
            }
            return finish(report, args)

        try:
            before = read_counterparty_combo_state(jab, found["vm_id"], found["context"])
            before_popups = collect_visible_popup_windows(jab)
            action_ok = bool(
                jab.do_action(
                    found["vm_id"],
                    found["context"],
                    action_name="togglePopup",
                )
            )
            popup_probe = wait_for_popup_windows(
                jab,
                before_popups,
                timeout=args.wait,
                interval=args.poll,
            )
            after_popups = collect_visible_popup_windows(jab)
            new_popups = collect_new_visible_popup_windows(jab, before_popups)
            popup_windows = new_popups or popup_probe.get("windows") or after_popups
            popup_cleanup = None
            popup_hwnd = (popup_windows[0] or {}).get("hwnd") if popup_windows else None
            if popup_hwnd:
                popup_cleanup = close_popup_hwnd(popup_hwnd)
            report = {
                "ok": bool(action_ok and popup_windows),
                "scope": scope,
                "path": found.get("path") or path,
                "find_method": found.get("method") or "path",
                "path_found": strip_handles(path_found),
                "near_label_found": strip_handles(near_label_found),
                "before": before,
                "toggle_ok": action_ok,
                "popup_probe": popup_probe,
                "new_popups": summarize_windows(new_popups, args.max_depth, args.max_children),
                "popup_windows": summarize_windows(
                    popup_windows, args.max_depth, args.max_children
                ),
                "after_popups": summarize_windows(
                    after_popups, args.max_depth, args.max_children
                ),
                "popup_cleanup": popup_cleanup,
                "reason": None
                if action_ok and popup_windows
                else "popup not detected after togglePopup",
            }
        finally:
            jab.release_contexts(found["vm_id"], found["owned_contexts"])
    finally:
        jab.close()

    return finish(report, args)


def find_counterparty_combo_near_label(jab, scope_hwnd, label=COUNTERPARTY_LABEL):
    result = {
        "ok": False,
        "method": "near-label",
        "label": label,
        "scope_hwnd": int(scope_hwnd) if scope_hwnd else None,
        "candidates": [],
    }
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if scope_hwnd is not None and int(hwnd) != int(scope_hwnd):
            continue
        if class_name != "SunAwtCanvas" or not visible:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            continue

        controls = []
        owned = []
        selected_contexts = set()
        try:
            collect_controls_for_bounds_scan(
                jab,
                vm_id.value,
                root_context.value,
                controls,
                owned,
                require_showing=True,
                depth=0,
            )
            labels = [
                (context, info)
                for context, info in controls
                if (info.role_en_US.strip() or info.role.strip()).lower() == "label"
                and info.name.strip() == label
                and jab.context_info_has_valid_bounds(info)
            ]
            labels.sort(key=lambda item: (item[1].y, item[1].x))
            for label_context, label_info in labels:
                label_mid_y = label_info.y + label_info.height / 2
                label_right = label_info.x + label_info.width
                row_candidates = []
                for context, info in controls:
                    role = (info.role_en_US.strip() or info.role.strip()).lower()
                    if role != "combo box":
                        continue
                    if not jab.context_info_has_valid_bounds(info):
                        continue
                    mid_y = info.y + info.height / 2
                    if info.x <= label_right:
                        continue
                    if abs(mid_y - label_mid_y) > max(label_info.height, 24):
                        continue
                    if info.x - label_right > 520:
                        continue
                    actions = jab.get_action_names(vm_id.value, context)
                    row_candidates.append(
                        {
                            "context": context,
                            "info": info,
                            "score": (info.x - label_right, abs(mid_y - label_mid_y)),
                            "actions": actions,
                        }
                    )

                row_candidates.sort(key=lambda item: item["score"])
                for item in row_candidates:
                    result["candidates"].append(
                        {
                            "label": info_to_dict(label_info),
                            "control": info_to_dict(item["info"]),
                            "actions": item["actions"],
                            "score": list(item["score"]),
                        }
                    )
                if row_candidates:
                    target = row_candidates[0]
                    selected_contexts = {target["context"], label_context}
                    release = [
                        context
                        for context in owned
                        if context not in selected_contexts
                    ]
                    jab.release_contexts(vm_id.value, release)
                    return {
                        "ok": True,
                        "method": "near-label",
                        "label": label,
                        "context": target["context"],
                        "vm_id": vm_id.value,
                        "owned_contexts": list(selected_contexts),
                        "window": {
                            "hwnd": int(hwnd),
                            "title": title,
                            "class_name": class_name,
                            "pid": int(pid or 0),
                            "visible": bool(visible),
                        },
                        "target": {
                            "label": info_to_dict(label_info),
                            "control": info_to_dict(target["info"]),
                            "actions": target["actions"],
                            "score": list(target["score"]),
                        },
                        "candidate_count": len(result["candidates"]),
                        "candidates": result["candidates"],
                    }
        finally:
            if not selected_contexts:
                jab.release_contexts(vm_id.value, owned)

    result["reason"] = "未在往来对象标签右侧找到 combo box"
    return result


def resolve_current_scope(jab, scope_hwnd=None):
    candidates = []
    fg_root = foreground_root_hwnd()
    for hwnd, title, class_name, pid, visible in enum_windows(include_children=True):
        if class_name != "SunAwtCanvas" or not visible:
            continue
        if not jab.dll.isJavaWindow(hwnd):
            continue
        if scope_hwnd is not None and int(scope_hwnd) != int(hwnd):
            continue
        scope = resolve_receipt_header_anchor_in_canvas(jab, hwnd, timeout=0.6)
        if not scope.get("ok"):
            continue
        candidates.append(
            {
                "ok": True,
                "hwnd": int(hwnd),
                "title": title,
                "class_name": class_name,
                "pid": int(pid or 0),
                "visible": bool(visible),
                "root_hwnd": root_hwnd(hwnd),
                "is_foreground_root": bool(fg_root and root_hwnd(hwnd) == fg_root),
                **scope,
            }
        )
    if not candidates:
        return {"ok": False, "reason": "no receipt header scope found"}
    candidates.sort(
        key=lambda item: (
            not item.get("is_foreground_root"),
            item.get("root_hwnd") or 0,
            item.get("hwnd") or 0,
        )
    )
    return candidates[0]


def foreground_root_hwnd():
    if sys.platform != "win32":
        return 0
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return 0
    return int(user32.GetAncestor(hwnd, 2) or 0)


def root_hwnd(hwnd):
    if sys.platform != "win32" or not hwnd:
        return 0
    return int(ctypes.windll.user32.GetAncestor(hwnd, 2) or 0)


def wait_for_popup_windows(jab, before_windows, timeout=0.8, interval=0.08):
    deadline = time.monotonic() + max(float(timeout or 0), 0.0)
    last_windows = []
    while time.monotonic() < deadline:
        last_windows = collect_new_visible_popup_windows(jab, before_windows)
        if last_windows:
            return {"ok": True, "windows": last_windows}
        time.sleep(interval)
    return {
        "ok": False,
        "windows": last_windows,
        "reason": "popup not detected",
    }


def summarize_windows(windows, max_depth, max_children):
    result = []
    for window in windows or []:
        role_counts = Counter()
        for control in window.get("all_controls", []) or []:
            role = str(control.get("role") or "").strip().lower()
            if role:
                role_counts[role] += 1
        result.append(
            {
                "hwnd": window.get("hwnd"),
                "title": window.get("title"),
                "class_name": window.get("class_name"),
                "visible": window.get("visible"),
                "root": window.get("root"),
                "control_count": len(window.get("all_controls") or []),
                "role_counts": dict(sorted(role_counts.items())),
                "controls": [
                    {
                        "path": control.get("path"),
                        "role": control.get("role"),
                        "name": control.get("name"),
                        "description": control.get("description"),
                        "states": control.get("states"),
                        "bounds": control.get("bounds"),
                        "accessibleAction": control.get("accessibleAction"),
                        "actions": control.get("actions"),
                    }
                    for control in (window.get("all_controls") or [])[:200]
                ],
                "max_depth": max_depth,
                "max_children": max_children,
            }
        )
    return result


def strip_handles(found):
    result = {}
    for key, value in (found or {}).items():
        if key in {"context", "vm_id", "owned_contexts"}:
            continue
        result[key] = value
    return result


def finish(report, args):
    output_dir = ROOT / "logs"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"counterparty_popup_tree_{time.strftime('%Y%m%d_%H%M%S')}.json"
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
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
