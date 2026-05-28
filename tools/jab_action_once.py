import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402


def build_parser():
    parser = argparse.ArgumentParser(description="Run one JAB path action.")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--path", default=None)
    parser.add_argument("--title", default=None)
    parser.add_argument("--class-name", default=None)
    parser.add_argument("--name", default=None)
    parser.add_argument("--role", default=None)
    parser.add_argument("--action", default=None)
    parser.add_argument("--set-text", default=None)
    parser.add_argument("--set-text-near-label", default=None)
    parser.add_argument(
        "--dry-run-near-label",
        default=None,
        help="Read-only: print the text field that would be selected near this label.",
    )
    parser.add_argument("--guard-path", default=None)
    parser.add_argument("--guard-name", default=None)
    parser.add_argument("--guard-role", default=None)
    parser.add_argument(
        "--click-mode",
        choices=("action", "bounds"),
        default="action",
        help="Use JAB AccessibleAction or click the JAB-reported control bounds.",
    )
    parser.add_argument("--wait", type=float, default=0.0)
    parser.add_argument("--timeout", type=float, default=3.0)
    parser.add_argument("--require-showing", action="store_true")
    return parser


def main():
    args = build_parser().parse_args()
    if (
        args.dry_run_near_label is None
        and args.set_text_near_label is None
        and not args.path
    ):
        raise SystemExit("--path is required unless --dry-run-near-label is used")

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        if args.dry_run_near_label is not None:
            result = jab.describe_text_near_label(
                args.dry_run_near_label,
                title=args.title,
                class_name=args.class_name,
                require_showing=args.require_showing,
            )
            print(json.dumps(result, ensure_ascii=False, indent=2))
            return 0
        if args.set_text_near_label is not None:
            ok = jab.set_text_near_label(
                args.set_text_near_label,
                args.set_text,
                title=args.title,
                class_name=args.class_name,
                wait=args.wait,
                timeout=args.timeout,
                require_showing=args.require_showing,
            )
        elif args.set_text is not None:
            ok = jab.set_text_by_path(
                args.path,
                args.set_text,
                title=args.title,
                class_name=args.class_name,
                name=args.name,
                role=args.role,
                guard_path=args.guard_path,
                guard_name=args.guard_name,
                guard_role=args.guard_role,
                wait=args.wait,
                timeout=args.timeout,
                require_showing=args.require_showing,
            )
        else:
            ok = jab.do_action_by_path(
                args.path,
                title=args.title,
                class_name=args.class_name,
                name=args.name,
                role=args.role,
                action_name=args.action,
                click_mode=None if args.click_mode == "action" else args.click_mode,
                wait=args.wait,
                timeout=args.timeout,
                require_showing=args.require_showing,
            )
        print(f"ok={ok}")
        return 0 if ok else 1
    finally:
        jab.close()


if __name__ == "__main__":
    raise SystemExit(main())
