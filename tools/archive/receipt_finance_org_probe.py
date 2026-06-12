import argparse
import json
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_query_fill import (  # noqa: E402
    ReceiptPageGuardError,
    ensure_query_window,
    guard_receipt_parent_page,
)


def read_near_label(jab, label, title, class_name, require_showing=True):
    result = jab.find_text_context_near_label_once(
        label,
        title=title,
        class_name=class_name,
        require_showing=require_showing,
    )
    context, vm_id, owned_contexts, label_info, text_info, window_info = result
    if not context:
        return {
            "found": False,
            "label": label,
            "window": window_info,
        }
    try:
        return {
            "found": True,
            "label": label,
            "window": window_info,
            "value": jab.get_text_context_value(vm_id, context),
            "label_info": jab.info_to_dict(label_info),
            "text_info": jab.info_to_dict(text_info),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def focus_near_label(jab, label, title, class_name, require_showing=True):
    result = jab.find_text_context_near_label_once(
        label,
        title=title,
        class_name=class_name,
        require_showing=require_showing,
    )
    context, vm_id, owned_contexts, label_info, text_info, window_info = result
    if not context:
        return {"found": False, "window": window_info}
    try:
        request_focus_ok = True
        if hasattr(jab.dll, "requestFocus"):
            request_focus_ok = bool(jab.dll.requestFocus(vm_id, context))
        return {
            "found": True,
            "request_focus_ok": request_focus_ok,
            "window": window_info,
            "label_info": jab.info_to_dict(label_info),
            "text_info": jab.info_to_dict(text_info),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def main():
    parser = argparse.ArgumentParser(
        description="Probe the receipt query finance organization field."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--value", default="A003")
    parser.add_argument("--label", default="收款财务组织")
    parser.add_argument("--no-open-query", action="store_true")
    parser.add_argument("--skip-page-guard", action="store_true")
    parser.add_argument("--focus-click", action="store_true")
    parser.add_argument("--describe-nearby", action="store_true")
    parser.add_argument(
        "--click-nearby-role",
        choices=("push button", "combo box"),
        default=None,
    )
    parser.add_argument("--click-nearby-index", type=int, default=0)
    parser.add_argument("--timeout", type=float, default=3.0)
    parser.add_argument("--wait", type=float, default=0.5)
    args = parser.parse_args()

    config = load_config(args.config)
    query_cfg = config["receipt_entry"]["query"]
    jab_cfg = query_cfg["jab"]
    field_cfg = jab_cfg["fields"]["finance_org"]
    label = args.label or field_cfg["label"]
    title = jab_cfg["dialog_title"]
    class_name = jab_cfg["dialog_class"]

    jab = JABOperator(config)
    report = {
        "value": args.value,
        "label": label,
        "query_window": None,
        "page_guard": None,
        "near_label": {},
        "focus": {},
        "nearby": None,
        "click_nearby": None,
        "windows_after_click": None,
    }
    try:
        if not args.skip_page_guard:
            try:
                report["page_guard"] = guard_receipt_parent_page(jab, config, query_cfg)
            except ReceiptPageGuardError as exc:
                report["page_guard"] = {"ok": False, "error": str(exc)}
                print(json.dumps(report, ensure_ascii=True, indent=2))
                return 2

        report["query_window"] = ensure_query_window(
            jab,
            config,
            query_cfg,
            jab_cfg,
            skip_open=args.no_open_query,
        )
        if not report["query_window"] and not args.no_open_query:
            print(json.dumps(report, ensure_ascii=True, indent=2))
            return 3

        report["near_label"]["before"] = read_near_label(
            jab, label, title, class_name, require_showing=True
        )
        if args.focus_click:
            report["focus"]["before_label_write"] = focus_near_label(
                jab, label, title, class_name, require_showing=True
            )
        report["near_label"]["write_ok"] = jab.set_text_near_label(
            label,
            args.value,
            title=title,
            class_name=class_name,
            timeout=args.timeout,
            wait=args.wait,
            require_showing=True,
        )
        report["near_label"]["after"] = read_near_label(
            jab, label, title, class_name, require_showing=True
        )
        if args.describe_nearby:
            report["nearby"] = jab.describe_controls_near_label(
                label,
                title=title,
                class_name=class_name,
                require_showing=True,
            )
        if args.click_nearby_role:
            report["click_nearby"] = jab.click_control_near_label(
                label,
                args.click_nearby_role,
                index=args.click_nearby_index,
                title=title,
                class_name=class_name,
                require_showing=True,
                wait=args.wait,
            )
            time.sleep(args.wait)
            report["windows_after_click"] = [
                {
                    "hwnd": item[0],
                    "title": item[1],
                    "class": item[2],
                    "pid": item[3],
                    "visible": item[4],
                }
                for item in jab.get_scoped_windows(include_children=True)
                if item[4] and (item[1] or item[2].startswith("SunAwt"))
            ]

        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0
    finally:
        jab.close()


if __name__ == "__main__":
    raise SystemExit(main())
