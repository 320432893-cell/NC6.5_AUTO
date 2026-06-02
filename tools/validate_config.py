import argparse
import json
from datetime import date
from pathlib import Path


SAVE_STRATEGIES = {"single", "bottom_up", "safe_batch_by_pending_row"}
SAVE_TRIGGERS = {"jab_button", "hotkey"}
HOTKEY_ACTIVATE_POLICIES = {"always", "first", "foreground_guard"}
DUPE_MATCH_POLICIES = {"stop", "skip"}
OPEN_QUERY_METHODS = {"hotkey", "jab_action"}


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
    excel_cfg = receipt_cfg.get("excel")
    if excel_cfg is not None:
        _validate_receipt_excel(excel_cfg, errors)
    query_cfg = receipt_cfg.get("query")
    if query_cfg is not None:
        _validate_receipt_query(query_cfg, errors)
    candidate_cfg = receipt_cfg.get("candidate_check")
    if candidate_cfg is not None:
        _validate_receipt_candidate_check(candidate_cfg, errors)
    organizations = _require(
        receipt_cfg,
        "finance_organizations",
        list,
        errors,
        prefix="receipt_entry",
    )
    accounts = _require(receipt_cfg, "accounts", list, errors, prefix="receipt_entry")

    org_codes = set()
    org_short_names = {}
    if isinstance(organizations, list):
        for index, organization in enumerate(organizations):
            prefix = f"receipt_entry.finance_organizations[{index}]"
            if not isinstance(organization, dict):
                errors.append(f"{prefix} must be an object")
                continue
            code = _require_non_empty_str(organization, "code", errors, prefix=prefix)
            _require_non_empty_str(organization, "name", errors, prefix=prefix)
            short_name = _require_non_empty_str(
                organization, "short_name", errors, prefix=prefix
            )
            if code:
                if code in org_codes:
                    errors.append(f"{prefix}.code must be unique, got {code!r}")
                org_codes.add(code)
                org_short_names[code] = short_name

    account_keys = set()
    if isinstance(accounts, list):
        for index, account in enumerate(accounts):
            prefix = f"receipt_entry.accounts[{index}]"
            if not isinstance(account, dict):
                errors.append(f"{prefix} must be an object")
                continue
            org_code = _require_non_empty_str(
                account, "organization_code", errors, prefix=prefix
            )
            org_short_name = _require_non_empty_str(
                account, "organization_short_name", errors, prefix=prefix
            )
            account_label = _require_non_empty_str(
                account, "account_label", errors, prefix=prefix
            )
            account_no = _require_non_empty_str(
                account, "account_no", errors, prefix=prefix
            )
            if org_code and org_code not in org_codes:
                errors.append(
                    f"{prefix}.organization_code must reference finance_organizations, "
                    f"got {org_code!r}"
                )
            expected_short_name = org_short_names.get(org_code)
            if expected_short_name and org_short_name != expected_short_name:
                errors.append(
                    f"{prefix}.organization_short_name must match organization "
                    f"{org_code!r}, got {org_short_name!r}"
                )
            key = (org_code, account_label, account_no)
            if key in account_keys:
                errors.append(
                    f"{prefix} duplicates account mapping "
                    f"{org_code!r}/{account_label!r}/{account_no!r}"
                )
            account_keys.add(key)


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
        "organization_column",
        "nc_done_column",
    ):
        _require_non_empty_str(excel_cfg, key, errors, prefix="receipt_entry.excel")
    _positive_int(excel_cfg, "header_row", errors, prefix="receipt_entry.excel")
    _iso_date(excel_cfg, "start_date", errors, prefix="receipt_entry.excel")


def _validate_receipt_query(query_cfg, errors):
    if not isinstance(query_cfg, dict):
        errors.append("receipt_entry.query must be an object")
        return

    for key in (
        "date_from",
        "date_to",
        "finance_org_field",
        "finance_org_operator",
        "document_date_field",
        "document_date_operator",
    ):
        _require_non_empty_str(query_cfg, key, errors, prefix="receipt_entry.query")
    if query_cfg.get("date_from") != "{today}":
        _iso_date(query_cfg, "date_from", errors, prefix="receipt_entry.query")
    if query_cfg.get("date_to") != "{today}":
        _iso_date(query_cfg, "date_to", errors, prefix="receipt_entry.query")

    result_columns = _require(
        query_cfg,
        "result_columns",
        dict,
        errors,
        prefix="receipt_entry.query",
    )
    if isinstance(result_columns, dict):
        for key in ("document_date", "original_amount", "customer"):
            _require_non_empty_str(
                result_columns,
                key,
                errors,
                prefix="receipt_entry.query.result_columns",
            )

    jab_cfg = query_cfg.get("jab")
    if jab_cfg is not None:
        _validate_receipt_query_jab(jab_cfg, errors)


def _validate_receipt_query_jab(jab_cfg, errors):
    if not isinstance(jab_cfg, dict):
        errors.append("receipt_entry.query.jab must be an object")
        return

    for key in ("dialog_title", "dialog_class", "confirm_button_path"):
        _require_non_empty_str(jab_cfg, key, errors, prefix="receipt_entry.query.jab")

    fields = _require(
        jab_cfg,
        "fields",
        dict,
        errors,
        prefix="receipt_entry.query.jab",
    )
    if not isinstance(fields, dict):
        return

    for key in ("finance_org", "document_date"):
        field = _require(
            fields,
            key,
            dict,
            errors,
            prefix="receipt_entry.query.jab.fields",
        )
        if isinstance(field, dict):
            _validate_receipt_query_jab_field(
                field,
                errors,
                prefix=f"receipt_entry.query.jab.fields.{key}",
                range_field=key == "document_date",
            )

    for key in ("original_amount", "customer"):
        field = fields.get(key)
        if field is not None:
            if not isinstance(field, dict):
                errors.append(f"receipt_entry.query.jab.fields.{key} must be an object")
                continue
            _validate_receipt_query_jab_field(
                field,
                errors,
                prefix=f"receipt_entry.query.jab.fields.{key}",
                range_field=key == "original_amount",
            )


def _validate_receipt_query_jab_field(field, errors, prefix, range_field=False):
    for key in ("label", "operator"):
        _require_non_empty_str(field, key, errors, prefix=prefix)
    if range_field:
        for key in ("from_text_path", "to_text_path"):
            _require_non_empty_str(field, key, errors, prefix=prefix)
    else:
        _require_non_empty_str(field, "text_path", errors, prefix=prefix)


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


def _require(mapping, key, expected_type, errors, prefix=""):
    label = f"{prefix}.{key}" if prefix else key
    if key not in mapping:
        errors.append(f"{label} is required")
        return None
    value = mapping[key]
    if not isinstance(value, expected_type):
        errors.append(f"{label} must be {expected_type.__name__}")
        return None
    return value


def _require_non_empty_str(mapping, key, errors, prefix=""):
    value = _require(mapping, key, str, errors, prefix=prefix)
    if isinstance(value, str) and not value.strip():
        label = f"{prefix}.{key}" if prefix else key
        errors.append(f"{label} must be non-empty")
    return value.strip() if isinstance(value, str) else None


def _enum(mapping, key, allowed, errors, prefix=""):
    label = f"{prefix}.{key}" if prefix else key
    value = mapping.get(key)
    if value not in allowed:
        errors.append(f"{label} must be one of {sorted(allowed)}, got {value!r}")


def _optional_enum(mapping, key, allowed, errors, prefix=""):
    if key not in mapping:
        return
    _enum(mapping, key, allowed, errors, prefix=prefix)


def _positive_int(mapping, key, errors, prefix=""):
    label = f"{prefix}.{key}" if prefix else key
    value = mapping.get(key)
    if not isinstance(value, int) or value < 1:
        errors.append(f"{label} must be a positive integer")


def _non_negative_int(mapping, key, errors, prefix=""):
    label = f"{prefix}.{key}" if prefix else key
    value = mapping.get(key)
    if not isinstance(value, int) or value < 0:
        errors.append(f"{label} must be a non-negative integer")


def _non_negative_number(mapping, key, errors, prefix=""):
    label = f"{prefix}.{key}" if prefix else key
    value = mapping.get(key)
    if not isinstance(value, int | float) or value < 0:
        errors.append(f"{label} must be a non-negative number")


def _iso_date(mapping, key, errors, prefix=""):
    label = f"{prefix}.{key}" if prefix else key
    value = mapping.get(key)
    if value in (None, ""):
        return
    try:
        date.fromisoformat(str(value))
    except ValueError:
        errors.append(f"{label} must be YYYY-MM-DD")


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
