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
    FINANCE_ORG_ACCEPTED_TEXT,
    find_finance_org_header_scope_by_paths,
    receipt_header_dynamic_prefix,
)


def main():
    parser = argparse.ArgumentParser(
        description="只读探测当前收款单 canvas 内财务组织锚点"
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--hwnd", type=int, default=None)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    jab = JABOperator(cfg)
    try:
        jab.ensure_started()
        report = probe(jab, args)
    finally:
        jab.close()

    output_path = (
        logs_dir()
        / f"receipt_finance_org_path_probe_{time.strftime('%Y%m%d_%H%M%S')}.json"
    )
    output_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"探测详情: {output_path}")
        print_summary(report)


def probe(jab, args):
    scope = find_finance_org_header_scope_by_paths(
        jab,
        args.hwnd,
    )
    return {
        "scope": redact_scope(scope),
        "scope_hwnd": scope.get("scope_hwnd") or args.hwnd,
        "expected_accepted_text": FINANCE_ORG_ACCEPTED_TEXT,
        "best_dynamic_index": scope.get("dynamic_index") if scope.get("ok") else None,
        "best_dynamic_prefix": (
            receipt_header_dynamic_prefix(scope.get("dynamic_index"))
            if scope.get("ok")
            else None
        ),
        "policy": "current-canvas semantic scan only; no dynamic-index path probing",
    }


def is_finance_org_label(result):
    if not result.get("ok"):
        return False
    values = {
        str(result.get("name") or "").strip(),
        str(result.get("description") or "").strip(),
        str(result.get("text") or "").strip(),
    }
    return "财务组织(O)" in values


def is_finance_org_accepted(result):
    if not result.get("ok"):
        return False
    values = {
        str(result.get("name") or "").strip(),
        str(result.get("description") or "").strip(),
        str(result.get("text") or "").strip(),
    }
    return FINANCE_ORG_ACCEPTED_TEXT in values


def redact_scope(scope):
    if not isinstance(scope, dict):
        return scope
    return {
        key: value
        for key, value in scope.items()
        if key not in {"context", "owned_contexts", "vm_id"}
    }


def print_summary(report):
    scope = report.get("scope") or {}
    print(
        "scope: "
        f"ok={scope.get('ok')} mode={scope.get('mode')} "
        f"dynamic_index={scope.get('dynamic_index')} hwnd={report.get('scope_hwnd')}"
    )
    print(
        "best: "
        f"dynamic_index={report.get('best_dynamic_index')} "
        f"prefix={report.get('best_dynamic_prefix')}"
    )
    for item in report.get("candidates") or []:
        label = item.get("label") or {}
        text = item.get("text") or {}
        if not (
            label.get("ok")
            or text.get("ok")
            or item.get("label_matches")
            or item.get("text_accepted")
        ):
            continue
        print(
            f"- index {item.get('dynamic_index')}: "
            f"variant={item.get('variant')} "
            f"label_ok={label.get('ok')} label_match={item.get('label_matches')} "
            f"text_ok={text.get('ok')} text_accepted={item.get('text_accepted')}"
        )
        print(f"  label_path={item.get('label_path')}")
        print(f"  text_path={item.get('text_path')}")
        print(
            "  label "
            f"name={label.get('name')!r} desc={label.get('description')!r} "
            f"role={label.get('role')!r} states={label.get('states')!r}"
        )
        print(
            "  text  "
            f"name={text.get('name')!r} desc={text.get('description')!r} "
            f"text={text.get('text')!r} role={text.get('role')!r} "
            f"states={text.get('states')!r}"
        )


if __name__ == "__main__":
    main()
