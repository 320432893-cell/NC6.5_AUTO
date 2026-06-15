# 职责: T0 探测当前收款单录入页、F3 查询条件页、查询结果页动态 path 前缀是否共用同一页签索引
# 不做什么: 不保存单据，不写 Excel，不修改正式配置
# 允许依赖层: core JAB/config、tools.receipt_self_made_fill_trial、tools.receipt_query_fill/pagination
# 谁不应该 import: 正式流程、core 模块和测试不应 import 本临时探针
# 生命周期: T0 临时探针（删除条件：录入页/F3 查询/结果页前缀复用规则确认并沉淀到正式模块/文档）

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_keyboard_utils import send_virtual_key  # noqa: E402
from tools.receipt_query_fill import wait_after_query_confirm  # noqa: E402
from tools.receipt_query_dynamic_fields import (  # noqa: E402
    find_query_condition_scope,
    set_query_dynamic_text,
)
from tools.receipt_query_pagination_paths import (  # noqa: E402
    resolve_receipt_pagination_paths_dynamic,
)
from tools.receipt_self_made_fill_trial import locate_receipt_header_scope  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--open-query", action="store_true")
    parser.add_argument("--confirm-query", action="store_true")
    parser.add_argument("--org-code", default="A001")
    parser.add_argument("--date-from", default="2026-01-01")
    parser.add_argument("--date-to", default="2026-06-15")
    args = parser.parse_args(argv)

    config = load_config(args.config)
    query_cfg = config["receipt_entry"]["query"]
    jab_cfg = query_cfg["jab"]
    report = {}

    jab = JABOperator(config)
    try:
        jab.ensure_started()
        report["header_scope"] = locate_receipt_header_scope(jab)
        if args.open_query:
            report["query_window_opened"] = open_query_by_f3(jab, query_cfg, jab_cfg)
            report["query_condition_scope"] = find_query_condition_scope(jab, jab_cfg)
    finally:
        jab.close()

    if args.confirm_query:
        jab = JABOperator(config)
        try:
            jab.ensure_started()
            query_scope = find_query_condition_scope(jab, jab_cfg)
            report["confirm_query_condition_scope"] = query_scope
            if not query_scope.get("ok"):
                raise RuntimeError(f"query condition scope not found: {query_scope}")
            report["confirm_query_fill"] = {
                "finance_org": set_query_dynamic_text(
                    jab,
                    jab_cfg,
                    query_scope,
                    "finance_org",
                    args.org_code,
                ),
                "document_date_from": set_query_dynamic_text(
                    jab,
                    jab_cfg,
                    query_scope,
                    "document_date_from",
                    args.date_from,
                ),
                "document_date_to": set_query_dynamic_text(
                    jab,
                    jab_cfg,
                    query_scope,
                    "document_date_to",
                    args.date_to,
                ),
            }
            confirm_ok = jab.do_action_by_path(
                jab_cfg["confirm_button_path"],
                title=jab_cfg["dialog_title"],
                class_name=jab_cfg["dialog_class"],
                role="push button",
                action_name="单击",
                wait=float(jab_cfg.get("confirm_wait", 0.0)),
                timeout=float(jab_cfg.get("confirm_timeout", 1.0)),
                require_showing=True,
            )
            report["confirm_query_ok"] = bool(confirm_ok)
            report["result_wait"] = wait_after_query_confirm(jab, query_cfg)
            report["result_pagination_scope"] = (
                resolve_receipt_pagination_paths_dynamic(
                    jab,
                    query_cfg,
                )
            )
        finally:
            jab.close()

    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    return 0


def open_query_by_f3(jab, query_cfg, jab_cfg):
    existing = jab.wait_window_by_title(
        jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        timeout=float(query_cfg.get("existing_dialog_timeout", 0.1)),
        include_children=bool(query_cfg.get("dialog_include_children", True)),
        visible_only=bool(query_cfg.get("dialog_visible_only", True)),
    )
    if existing:
        return True
    send_virtual_key(0x72)
    return bool(
        jab.wait_window_by_title(
            jab_cfg["dialog_title"],
            class_name=jab_cfg["dialog_class"],
            timeout=float(query_cfg.get("open_timeout", 5)),
            include_children=bool(query_cfg.get("dialog_include_children", True)),
            visible_only=bool(query_cfg.get("dialog_visible_only", True)),
        )
    )


if __name__ == "__main__":
    raise SystemExit(main())
