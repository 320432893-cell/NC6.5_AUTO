import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_self_made_fill_trial import post_key_to_hwnd  # noqa: E402

HEADER_FORM_BASE_PATH = "0.0.1.0.0.0.0.2.0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0"


def main():
    parser = argparse.ArgumentParser(
        description="Try only receipt header bank account input."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--account", required=True)
    parser.add_argument("--mode", choices=("settext", "type"), default="settext")
    parser.add_argument("--wait", type=float, default=1.5)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    path = f"{HEADER_FORM_BASE_PATH}.15.0"
    jab = JABOperator(cfg)
    report = {"path": path, "account": args.account}
    try:
        jab.ensure_started()
        context, vm_id, owned, window = jab.find_context_by_path_once(
            path,
            class_name="SunAwtCanvas",
            role="text",
            require_showing=True,
            require_valid_bounds=True,
        )
        if not context:
            report.update({"ok": False, "reason": "account text path not found"})
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 1
        try:
            before = jab.get_context_info(vm_id, context)
            text_before = jab.get_text_context_value(vm_id, context)
            if hasattr(jab.dll, "requestFocus"):
                focus_ok = bool(jab.dll.requestFocus(vm_id, context))
            else:
                focus_ok = None
            if args.mode == "type":
                set_ok = True
                jab.type_text(args.account, interval=0.01)
                time.sleep(0.3)
            else:
                set_ok = jab.set_text_context(vm_id, context, args.account)
            enter_ok = post_key_to_hwnd(window.get("hwnd"), "enter")
            time.sleep(args.wait)
            after = jab.get_context_info(vm_id, context)
            text_after = jab.get_text_context_value(vm_id, context)
            report.update(
                {
                    "ok": bool(set_ok),
                    "window": window,
                    "mode": args.mode,
                    "set_ok": bool(set_ok),
                    "focus_ok": focus_ok,
                    "enter_ok": bool(enter_ok),
                    "description_before": before.description.strip()
                    if before
                    else None,
                    "text_before": text_before,
                    "description_after": after.description.strip() if after else None,
                    "text_after": text_after,
                }
            )
        finally:
            jab.release_contexts(vm_id, owned)
    finally:
        jab.close()

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
