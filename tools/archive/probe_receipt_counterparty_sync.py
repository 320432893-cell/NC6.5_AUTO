import argparse
import ctypes
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_body_table_locator import locate_receipt_body_table  # noqa: E402
from tools.receipt_full_flow_entry import (  # noqa: E402
    find_counterparty_combo,
    read_counterparty_combo_state,
    read_counterparty_selected_option,
)
from tools.receipt_self_made_fill_trial import (  # noqa: E402
    receipt_header_dynamic_prefix,
)
from tools.archive.probe_receipt_counterparty_popup_tree import (  # noqa: E402
    resolve_current_scope,
    strip_handles,
)
from tools.archive.probe_receipt_counterparty_methods import (  # noqa: E402
    activate_and_press,
    action_embedded_customer_option,
    cleanup_all_visible_popups,
    cleanup_popups,
    request_focus,
    root_hwnd,
    select_embedded_customer_option,
)
from core.jab_probe import AccessibleTableCellInfo  # noqa: E402

EXPECTED = "客户"
KNOWN_OPTIONS = {"客户", "部门", "业务员", "供应商"}
DEFAULT_REPAIR_METHODS = (
    "embedded-select-enter,embedded-action-enter,activate-home-enter"
)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Read-only probe for header/detail counterparty synchronization."
    )
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--scope-hwnd", type=int, default=None)
    parser.add_argument("--row", type=int, default=0)
    parser.add_argument("--col", type=int, default=0)
    parser.add_argument("--max-rows", type=int, default=3)
    parser.add_argument(
        "--commit",
        action="store_true",
        help="Attempt repair when lower detail counterparty is blank/non-customer.",
    )
    parser.add_argument(
        "--repair-methods",
        default=DEFAULT_REPAIR_METHODS,
        help=(
            "Comma-separated repair methods: embedded-select-enter, "
            "embedded-action-enter, activate-home-enter, activate-enter, "
            "activate-esc."
        ),
    )
    parser.add_argument("--wait-after-repair", type=float, default=0.25)
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
        "read_only": not args.commit,
        "commit": bool(args.commit),
        "requested": {"row": args.row, "col": args.col},
        "scope": None,
        "before": None,
        "repair": None,
        "after": None,
    }
    try:
        jab.ensure_started()
        if args.commit:
            cleanup_all_visible_popups(jab)
        scope = resolve_current_scope(jab, args.scope_hwnd)
        report["scope"] = scope
        if not scope.get("ok"):
            report["reason"] = scope.get("reason") or "receipt header scope not found"
            return finish(report, args)
        scope_hwnd = scope.get("hwnd")
        dynamic_index = scope.get("dynamic_index")
        before = read_counterparty_snapshot(
            jab,
            dynamic_index=dynamic_index,
            scope_hwnd=scope_hwnd,
            row=args.row,
            col=args.col,
            max_rows=args.max_rows,
        )
        report["before"] = before
        report["ok"] = bool(before.get("ok"))

        if args.commit and should_repair(before):
            report["repair"] = repair_counterparty(
                jab,
                dynamic_index=dynamic_index,
                scope_hwnd=scope_hwnd,
                row=args.row,
                col=args.col,
                max_rows=args.max_rows,
                methods=parse_methods(args.repair_methods),
                wait_after=args.wait_after_repair,
            )
            report["after"] = report["repair"].get("final")
            report["ok"] = bool((report["after"] or {}).get("ok"))

        if not report["ok"]:
            latest = report.get("after") or report.get("before") or {}
            report["reason"] = latest.get("reason") or latest.get("status")
    finally:
        jab.close()
    return finish(report, args)


def read_counterparty_snapshot(jab, dynamic_index, scope_hwnd, row, col, max_rows):
    header = read_header_counterparty(jab, dynamic_index, scope_hwnd)
    detail = read_detail_counterparty(
        jab,
        row=row,
        col=col,
        max_rows=max_rows,
        scope_hwnd=scope_hwnd,
    )
    comparison = compare_counterparty(header, detail)
    return {
        "ok": bool(header.get("ok") and detail.get("ok") and comparison.get("ok")),
        "header": header,
        "detail": detail,
        "comparison": comparison,
        "status": comparison.get("status"),
        "reason": (
            header.get("reason")
            or detail.get("reason")
            or comparison.get("reason")
        ),
    }


def read_header_counterparty(jab, dynamic_index, scope_hwnd):
    result = {
        "ok": False,
        "label": "header.往来对象",
        "dynamic_index": dynamic_index,
        "dynamic_prefix": (
            receipt_header_dynamic_prefix(dynamic_index)
            if dynamic_index is not None
            else None
        ),
    }
    found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
    result["found"] = strip_handles(found)
    if not found.get("ok"):
        result["reason"] = found.get("reason")
        return result
    try:
        state = read_counterparty_combo_state(
            jab, found["vm_id"], found["context"]
        )
        embedded = read_counterparty_selected_option(
            jab, found["vm_id"], found["context"]
        )
        result.update(
            {
                "ok": bool(embedded.get("ok") or state),
                "state": state,
                "embedded": embedded,
                "selected": embedded.get("selected") or "",
                "text": first_non_empty(
                    state.get("description"),
                    state.get("text"),
                    state.get("name"),
                ),
                "effective": first_non_empty(
                    embedded.get("selected"),
                    state.get("description"),
                    state.get("text"),
                    state.get("name"),
                ),
                "window": found.get("window"),
                "source": found.get("source"),
                "path": found.get("path"),
            }
        )
        return result
    finally:
        jab.release_contexts(found["vm_id"], found["owned_contexts"])


def read_detail_counterparty(jab, row, col, max_rows, scope_hwnd):
    located = locate_receipt_body_table(jab, max_rows=max_rows, scope_hwnd=scope_hwnd)
    best = located.get("best")
    result = {
        "ok": False,
        "label": f"detail.row{row}.col{col}",
        "row": row,
        "col": col,
        "located": summarize_located(located),
    }
    if not best:
        result["reason"] = "receipt body table not located"
        return result
    context, vm_id, owned, window = jab.find_context_by_path_once(
        best["path"],
        class_name=(best.get("window") or {}).get("class_name"),
        scope_hwnd=scope_hwnd,
        role="table",
        require_showing=False,
        require_valid_bounds=False,
    )
    if not context:
        result["reason"] = "table context not found by path"
        return result
    try:
        table_info = jab.get_table_info(vm_id, context)
        if not table_info:
            result["reason"] = "table_info not readable"
            return result
        text, is_selected = jab.get_table_cell_text_and_selection(vm_id, context, row, col)
        cell = read_cell_context(jab, vm_id, context, row, col)
        result.update(
            {
                "ok": cell.get("ok"),
                "text": str(text or "").strip(),
                "is_selected": bool(is_selected),
                "cell": cell,
                "table": {
                    "path": best.get("path"),
                    "window": window or best.get("window"),
                    "row_count": int(table_info.rowCount),
                    "col_count": int(table_info.columnCount),
                    "bounds": best.get("bounds"),
                },
                "reason": cell.get("reason"),
            }
        )
        return result
    finally:
        jab.release_contexts(vm_id, owned)


def read_cell_context(jab, vm_id, table_context, row, col):
    if not hasattr(jab.dll, "getAccessibleTableCellInfo"):
        return {"ok": False, "reason": "getAccessibleTableCellInfo unavailable"}
    cell_info = AccessibleTableCellInfo()
    ok = jab.dll.getAccessibleTableCellInfo(
        vm_id,
        table_context,
        row,
        col,
        ctypes.byref(cell_info),
    )
    result = {
        "ok": False,
        "cell_info_ok": bool(ok),
        "accessible_context": int(cell_info.accessibleContext or 0),
        "index": int(cell_info.index),
        "row": int(cell_info.row),
        "col": int(cell_info.column),
        "is_selected": bool(cell_info.isSelected),
    }
    if not ok:
        result["reason"] = "getAccessibleTableCellInfo returned false"
        return result
    if not cell_info.accessibleContext:
        result["reason"] = "cell has no accessibleContext"
        return result
    info = jab.get_context_info(vm_id, cell_info.accessibleContext)
    text = jab.get_text_context_value(vm_id, cell_info.accessibleContext)
    if not info:
        result["reason"] = "cell context info not readable"
        return result
    result.update(
        {
            "ok": True,
            "name": info.name.strip(),
            "description": info.description.strip(),
            "text": str(text or "").strip(),
            "role": info.role_en_US.strip() or info.role.strip(),
            "states": info.states_en_US.strip() or info.states.strip(),
            "bounds": [info.x, info.y, info.width, info.height],
            "children_count": int(info.childrenCount),
        }
    )
    return result


def compare_counterparty(header, detail):
    header_selected = first_non_empty((header.get("embedded") or {}).get("selected"))
    header_text = first_non_empty(
        header.get("text"),
        (header.get("state") or {}).get("description"),
        (header.get("state") or {}).get("text"),
        (header.get("state") or {}).get("name"),
    )
    detail_table_text = first_non_empty(detail.get("text"))
    detail_cell_text = first_non_empty(
        detail.get("text"),
        (detail.get("cell") or {}).get("description"),
        (detail.get("cell") or {}).get("name"),
        (detail.get("cell") or {}).get("text"),
    )
    header_effective = first_non_empty(header_selected, header_text)
    detail_effective = detail_cell_text
    status = classify_counterparty(
        header_selected=header_selected,
        header_text=header_text,
        detail_table_text=detail_table_text,
        detail_text=detail_effective,
    )
    ok = status == "ok"
    return {
        "ok": ok,
        "status": status,
        "header_selected": header_selected,
        "header_text": header_text,
        "header_effective": header_effective,
        "detail_table_text": detail_table_text,
        "detail_text": detail_effective,
        "both_customer": (
            header_effective == EXPECTED and detail_effective == EXPECTED
        ),
        "needs_repair": status in {
            "header-selected-detail-blank",
            "header-customer-detail-non-customer",
            "blank",
            "wrong-or-unknown",
            "mismatch",
        },
        "reason": counterparty_reason(status),
    }


def classify_counterparty(header_selected, header_text, detail_table_text, detail_text):
    header_values = [value for value in (header_selected, header_text) if value]
    detail_values = [value for value in (detail_table_text, detail_text) if value]
    known_header_wrong = next(
        (value for value in header_values if value in KNOWN_OPTIONS and value != EXPECTED),
        "",
    )
    known_detail_wrong = next(
        (value for value in detail_values if value in KNOWN_OPTIONS and value != EXPECTED),
        "",
    )
    if known_header_wrong or known_detail_wrong:
        return "wrong-or-unknown"
    if detail_text == EXPECTED:
        return "ok"
    if header_selected == EXPECTED and not detail_text:
        return "header-selected-detail-blank"
    if (header_selected == EXPECTED or header_text == EXPECTED) and detail_text != EXPECTED:
        return "header-customer-detail-non-customer"
    if not header_values and not detail_values:
        return "blank"
    if header_values or detail_values:
        return "mismatch"
    return "unknown"


def counterparty_reason(status):
    reasons = {
        "ok": None,
        "header-selected-detail-blank": "上方子列表选中客户，但下方明细往来对象为空",
        "header-customer-detail-non-customer": "上方为客户，但下方明细不是客户",
        "blank": "上方和下方都未读到往来对象客户",
        "wrong-or-unknown": "读到非客户的往来对象选项",
        "mismatch": "上方和下方往来对象不一致",
        "unknown": "往来对象状态无法判断",
    }
    return reasons.get(status, "往来对象状态无法判断")


def should_repair(snapshot):
    comparison = (snapshot or {}).get("comparison") or {}
    return bool(comparison.get("needs_repair"))


def parse_methods(text):
    return [item.strip() for item in str(text or "").split(",") if item.strip()]


def repair_counterparty(
    jab,
    dynamic_index,
    scope_hwnd,
    row,
    col,
    max_rows,
    methods,
    wait_after,
):
    report = {
        "attempted": True,
        "methods": methods,
        "attempts": [],
        "final": None,
    }
    for method in methods:
        attempt = run_repair_method(jab, dynamic_index, scope_hwnd, method)
        time.sleep(max(float(wait_after or 0), 0.0))
        after = read_counterparty_snapshot(
            jab,
            dynamic_index=dynamic_index,
            scope_hwnd=scope_hwnd,
            row=row,
            col=col,
            max_rows=max_rows,
        )
        attempt["after"] = after
        attempt["success"] = bool(after.get("ok"))
        report["attempts"].append(attempt)
        cleanup_popups((attempt.get("visible_popups") or []))
        if attempt["success"]:
            report["final"] = after
            report["ok"] = True
            return report
    report["final"] = (
        report["attempts"][-1].get("after") if report["attempts"] else None
    )
    report["ok"] = bool((report["final"] or {}).get("ok"))
    return report


def run_repair_method(jab, dynamic_index, scope_hwnd, method):
    found = find_counterparty_combo(jab, dynamic_index, scope_hwnd=scope_hwnd)
    result = {
        "method": method,
        "found": strip_handles(found),
        "action": None,
        "visible_popups": [],
    }
    if not found.get("ok"):
        result["action"] = {"ok": False, "reason": found.get("reason")}
        return result
    try:
        vm_id = found["vm_id"]
        context = found["context"]
        window_hwnd = ((found.get("window") or {}).get("hwnd"))
        window_root = root_hwnd(window_hwnd) or window_hwnd
        focus = request_focus(jab, vm_id, context)
        if method == "embedded-select-enter":
            action = select_embedded_customer_option(
                jab, vm_id, context, press_enter=True
            )
        elif method == "embedded-action-enter":
            action = action_embedded_customer_option(
                jab, vm_id, context, press_enter=True
            )
        elif method == "activate-home-enter":
            action = activate_and_press(jab, window_root, ["home", "enter"])
        elif method == "activate-enter":
            action = activate_and_press(jab, window_root, ["enter"])
        elif method == "activate-esc":
            action = activate_and_press(jab, window_root, ["esc"])
        else:
            action = {"ok": False, "reason": "unknown repair method"}
        result["action"] = {**action, "request_focus_combo": focus}
        result["visible_popups"] = []
        return result
    except Exception as exc:
        result["action"] = {"ok": False, "error": repr(exc)}
        return result
    finally:
        jab.release_contexts(found["vm_id"], found["owned_contexts"])


def first_non_empty(*values):
    for value in values:
        text = str(value or "").strip()
        if text and text not in {"翸", "ǁ", "|", "||", "ɲ"}:
            return text
    return ""


def summarize_located(located):
    best = located.get("best")
    return {
        "best": {
            "path": best.get("path"),
            "window": best.get("window"),
            "row_count": best.get("row_count"),
            "col_count": best.get("col_count"),
            "score": best.get("score"),
            "reasons": best.get("reasons"),
            "bounds": best.get("bounds"),
            "rows": best.get("rows"),
        }
        if best
        else None,
        "candidate_count": len(located.get("candidates") or []),
    }


def finish(report, args):
    output_dir = ROOT / "logs"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / (
        f"counterparty_sync_probe_{time.strftime('%Y%m%d_%H%M%S')}.json"
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
