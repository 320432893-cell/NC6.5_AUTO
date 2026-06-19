import argparse
import json
import sys
from pathlib import Path

# Allow running directly (python tools/validate_config.py config.json) as well
# as importing as tools.validate_config: ensure the repo root is importable so
# the `tools.` sub-module imports below resolve in both modes.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Boundary: validate_config is the single validation truth source and dispatcher.
# It keeps the top-level traversal here and delegates the two largest receipt
# sub-trees to dedicated modules. Shared low-level helpers live in
# validate_config_primitives so the dispatcher can import the sub-validators
# without forming an import cycle.
from tools.validate_config_accounts import _validate_receipt_accounts  # noqa: E402
from tools.validate_config_primitives import (  # noqa: E402
    _enum,
    _iso_date,
    _non_negative_int,
    _non_negative_number,
    _optional_enum,
    _positive_int,
    _require,
    _require_non_empty_str,
)
from tools.validate_config_query import _validate_receipt_query  # noqa: E402

# 保存策略/触发/激活枚举的 SSOT 在 core.voucher_constants(core 与 jab_batch 共用)。
# 这里直接复用，避免与 core 运行时分支、jab_batch argparse choices 三处口径漂移。
from core.voucher_constants import (  # noqa: E402
    HOTKEY_ACTIVATE_POLICIES,
    SAVE_STRATEGIES,
    SAVE_TRIGGERS,
)

DUPE_MATCH_POLICIES = {"stop", "skip"}
OPEN_QUERY_METHODS = {"hotkey", "jab_action"}
RECEIPT_VALIDATION_MODES = {"strict", "skip_invalid_rows"}


def load_json(path):
    with Path(path).open(encoding="utf-8") as file:
        return json.load(file)


def validate_config(config):
    errors = []
    _require(config, "excel_path", str, errors)
    _require(config, "sheet_my", str, errors)
    _require(config, "jab", dict, errors)
    batch_cfg = _require(config, "jab_batch", dict, errors)
    if not isinstance(batch_cfg, dict):
        return errors

    jab_cfg = config.get("jab", {})
    if isinstance(jab_cfg, dict):
        _require(jab_cfg, "dll_path", str, errors, prefix="jab")
        for key in ("startup_wait", "search_timeout", "max_depth", "max_children"):
            _non_negative_number(jab_cfg, key, errors, prefix="jab")
        for key in ("amount_col", "partner_col", "selection_col"):
            _non_negative_int(jab_cfg, key, errors, prefix="jab")

    for key in ("key_col", "amount_out_col", "partner_out_col", "result_col"):
        _positive_int(batch_cfg, key, errors, prefix="jab_batch")
    for key in ("generated_voucher_col", "generated_date_col"):
        _non_negative_int(batch_cfg, key, errors, prefix="jab_batch")

    _enum(batch_cfg, "save_strategy", SAVE_STRATEGIES, errors, prefix="jab_batch")
    _enum(batch_cfg, "save_trigger", SAVE_TRIGGERS, errors, prefix="jab_batch")
    _enum(
        batch_cfg,
        "hotkey_activate_policy",
        HOTKEY_ACTIVATE_POLICIES,
        errors,
        prefix="jab_batch",
    )
    _optional_enum(
        batch_cfg,
        "duplicate_match_policy",
        DUPE_MATCH_POLICIES,
        errors,
        prefix="jab_batch",
    )
    for key in (
        "max_batch_size",
        "save_wait",
        "wait_between_save_batches",
        "save_success_timeout",
        "voucher_record_timeout",
    ):
        _non_negative_number(batch_cfg, key, errors, prefix="jab_batch")
    _positive_int(batch_cfg, "generated_voucher_max", errors, prefix="jab_batch")
    _iso_date(batch_cfg, "generated_date_value", errors, prefix="jab_batch")

    open_query = _require(batch_cfg, "open_query", dict, errors, prefix="jab_batch")
    if isinstance(open_query, dict):
        _enum(open_query, "method", OPEN_QUERY_METHODS, errors, prefix="open_query")
        _require(open_query, "key", str, errors, prefix="open_query")
        _require(open_query, "dialog_title", str, errors, prefix="open_query")
        _require(open_query, "dialog_class", str, errors, prefix="open_query")
        if open_query.get("click_mode") == "bounds":
            errors.append("open_query.click_mode=bounds is not allowed")
        for key in ("timeout", "activate_timeout", "process_timeout"):
            _non_negative_number(open_query, key, errors, prefix="open_query")

    steps = _require(
        batch_cfg, "switch_generated_steps", list, errors, prefix="jab_batch"
    )
    if isinstance(steps, list):
        for index, step in enumerate(steps):
            step_prefix = f"switch_generated_steps[{index}]"
            if not isinstance(step, dict):
                errors.append(f"{step_prefix} must be an object")
                continue
            _require(step, "path", str, errors, prefix=step_prefix)
            if "set_text" not in step and "action" not in step:
                errors.append(f"{step_prefix} must define set_text or action")
            for key in ("timeout", "guard_timeout", "wait"):
                _non_negative_number(step, key, errors, prefix=step_prefix)

    receipt_cfg = config.get("receipt_entry")
    if receipt_cfg is not None:
        _validate_receipt_entry(receipt_cfg, errors)

    return errors


def _validate_receipt_entry(receipt_cfg, errors):
    if not isinstance(receipt_cfg, dict):
        errors.append("receipt_entry must be an object")
        return

    _require(receipt_cfg, "state_label", str, errors, prefix="receipt_entry")
    if "schema_version" in receipt_cfg:
        _positive_int(receipt_cfg, "schema_version", errors, prefix="receipt_entry")
    excel_cfg = receipt_cfg.get("excel")
    if excel_cfg is not None:
        _validate_receipt_excel(excel_cfg, errors)
    query_cfg = receipt_cfg.get("query")
    if query_cfg is not None:
        _validate_receipt_query(query_cfg, errors)
    candidate_cfg = receipt_cfg.get("candidate_check")
    if candidate_cfg is not None:
        _validate_receipt_candidate_check(candidate_cfg, errors)
    validation_policy = receipt_cfg.get("validation_policy")
    if validation_policy is not None:
        _validate_receipt_validation_policy(validation_policy, errors)
    organizations = _require(
        receipt_cfg,
        "finance_organizations",
        list,
        errors,
        prefix="receipt_entry",
    )
    banks = receipt_cfg.get("banks")
    if banks is not None and not isinstance(banks, list):
        errors.append("receipt_entry.banks must be a list")
    accounts = _require(receipt_cfg, "accounts", list, errors, prefix="receipt_entry")
    detail_entry_policy = receipt_cfg.get("detail_entry_policy")
    if detail_entry_policy is not None:
        _validate_receipt_detail_entry_policy(detail_entry_policy, errors)

    _validate_receipt_accounts(organizations, banks, accounts, errors)


def _validate_receipt_detail_entry_policy(policy, errors):
    prefix = "receipt_entry.detail_entry_policy"
    if not isinstance(policy, dict):
        errors.append(f"{prefix} must be an object")
        return
    for key in ("main_line_order", "fee_line_order"):
        if key in policy:
            value = policy.get(key)
            if not isinstance(value, list) or not all(
                isinstance(item, str) and item.strip() for item in value
            ):
                errors.append(f"{prefix}.{key} must be a list of non-empty strings")
    for key in ("fee_add_row_hotkey", "extra_blank_row_delete_hotkey"):
        if key in policy:
            _require_non_empty_str(policy, key, errors, prefix=prefix)
    if "fee_clear_account" in policy and not isinstance(
        policy.get("fee_clear_account"), bool
    ):
        errors.append(f"{prefix}.fee_clear_account must be bool")


def _validate_receipt_excel(excel_cfg, errors):
    if not isinstance(excel_cfg, dict):
        errors.append("receipt_entry.excel must be an object")
        return

    for key in (
        "path",
        "sheet_name",
        "date_column",
        "payer_name_column",
        "raw_amount_column",
        "bank_column",
        "currency_column",
        "customer_code_column",
        "organization_column",
        "nc_done_column",
    ):
        _require_non_empty_str(excel_cfg, key, errors, prefix="receipt_entry.excel")
    if "fee_column" in excel_cfg:
        _require_non_empty_str(
            excel_cfg, "fee_column", errors, prefix="receipt_entry.excel"
        )
    if "result_sheet_name" in excel_cfg:
        _require_non_empty_str(
            excel_cfg, "result_sheet_name", errors, prefix="receipt_entry.excel"
        )
    _positive_int(excel_cfg, "header_row", errors, prefix="receipt_entry.excel")
    _positive_int(excel_cfg, "start_row", errors, prefix="receipt_entry.excel")
    if (
        isinstance(excel_cfg.get("header_row"), int)
        and isinstance(excel_cfg.get("start_row"), int)
        and excel_cfg["start_row"] <= excel_cfg["header_row"]
    ):
        errors.append("receipt_entry.excel.start_row must be greater than header_row")
    _iso_date(excel_cfg, "start_date", errors, prefix="receipt_entry.excel")


def _validate_receipt_candidate_check(candidate_cfg, errors):
    if not isinstance(candidate_cfg, dict):
        errors.append("receipt_entry.candidate_check must be an object")
        return

    _non_negative_int(
        candidate_cfg,
        "recent_months",
        errors,
        prefix="receipt_entry.candidate_check",
    )
    if candidate_cfg.get("from_date") not in (None, ""):
        _iso_date(
            candidate_cfg, "from_date", errors, prefix="receipt_entry.candidate_check"
        )
    value = candidate_cfg.get("only_blank_status")
    if not isinstance(value, bool):
        errors.append("receipt_entry.candidate_check.only_blank_status must be bool")


def _validate_receipt_validation_policy(policy, errors):
    if not isinstance(policy, dict):
        errors.append("receipt_entry.validation_policy must be an object")
        return
    _optional_enum(
        policy,
        "mode",
        RECEIPT_VALIDATION_MODES,
        errors,
        prefix="receipt_entry.validation_policy",
    )
    if "skip_invalid_rows" in policy and not isinstance(
        policy.get("skip_invalid_rows"), bool
    ):
        errors.append("receipt_entry.validation_policy.skip_invalid_rows must be bool")


def main():
    parser = argparse.ArgumentParser(description="Validate nc_auto_v2 config semantics")
    parser.add_argument("config", nargs="?", default="config.json")
    args = parser.parse_args()

    errors = validate_config(load_json(args.config))
    if errors:
        print("config validation failed:")
        for error in errors:
            print(f"- {error}")
        raise SystemExit(1)
    print("config validation passed")


if __name__ == "__main__":
    main()
