import argparse
import json
from datetime import date
from pathlib import Path


SAVE_STRATEGIES = {"single", "safe_batch_by_pending_row"}
SAVE_TRIGGERS = {"jab_button"}
DUPE_MATCH_POLICIES = {"stop", "skip"}
OPEN_QUERY_METHODS = {"hotkey", "jab_action"}
ACCOUNT_INPUT_STRATEGIES = {"detail_first", "reference_first"}
ACCOUNT_SUCCESS_RULES = {"non_empty", "exact_or_candidate"}
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
    _optional_enum(
        batch_cfg,
        "duplicate_match_policy",
        DUPE_MATCH_POLICIES,
        errors,
        prefix="jab_batch",
    )
    for key in (
        "pending_generate_batch_size",
        "max_batch_size",
        "save_wait",
        "wait_between_save_batches",
        "save_success_timeout",
        "voucher_record_timeout",
        "state_wait_timeout",
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
    if "excel_text_field_mappings" in receipt_cfg:
        _validate_receipt_excel_text_field_mappings(
            receipt_cfg.get("excel_text_field_mappings"),
            errors,
        )

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

    bank_ids = set()
    if isinstance(banks, list):
        for index, bank in enumerate(banks):
            prefix = f"receipt_entry.banks[{index}]"
            if not isinstance(bank, dict):
                errors.append(f"{prefix} must be an object")
                continue
            bank_id = _require_non_empty_str(bank, "id", errors, prefix=prefix)
            _require_non_empty_str(bank, "name", errors, prefix=prefix)
            aliases = bank.get("aliases", [])
            if aliases is not None and not isinstance(aliases, list):
                errors.append(f"{prefix}.aliases must be a list")
                aliases = []
            if "enabled" in bank and not isinstance(bank.get("enabled"), bool):
                errors.append(f"{prefix}.enabled must be bool")
            if bank_id:
                if bank_id in bank_ids:
                    errors.append(f"{prefix}.id must be unique, got {bank_id!r}")
                bank_ids.add(bank_id)

    account_keys = set()
    account_ids = set()
    account_lookup_owners = {}
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
            account_id = account.get("id")
            if account_id is not None:
                account_id = _require_non_empty_str(
                    account, "id", errors, prefix=prefix
                )
            account_label = _require_non_empty_str(
                account, "account_label", errors, prefix=prefix
            )
            account_no = _require_non_empty_str(
                account, "account_no", errors, prefix=prefix
            )
            header_currency_code = _require_non_empty_str(
                account, "header_currency_code", errors, prefix=prefix
            )
            if header_currency_code and header_currency_code not in {"USD", "CNY"}:
                errors.append(
                    f"{prefix}.header_currency_code must be USD or CNY, "
                    f"got {header_currency_code!r}"
                )
            bank_id = account.get("bank_id")
            if bank_id is not None:
                bank_id = _require_non_empty_str(
                    account, "bank_id", errors, prefix=prefix
                )
                if bank_id and bank_id not in bank_ids:
                    errors.append(
                        f"{prefix}.bank_id must reference receipt_entry.banks, got {bank_id!r}"
                    )
            for list_key in ("aliases", "excel_bank_aliases"):
                values = account.get(list_key, [])
                if values is not None and not isinstance(values, list):
                    errors.append(f"{prefix}.{list_key} must be a list")
                    continue
                for item_index, item in enumerate(values or []):
                    if not isinstance(item, str) or not item.strip():
                        errors.append(
                            f"{prefix}.{list_key}[{item_index}] must be a non-empty string"
                        )
            if "enabled" in account and not isinstance(account.get("enabled"), bool):
                errors.append(f"{prefix}.enabled must be bool")
            candidate_map = account.get("nc_candidates_by_currency")
            if candidate_map is not None:
                _validate_candidate_map(candidate_map, errors, prefix)
            entry_policy = account.get("entry_policy")
            if entry_policy is not None:
                _validate_account_entry_policy(entry_policy, errors, prefix)
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
            if account_id:
                if account_id in account_ids:
                    errors.append(f"{prefix}.id must be unique, got {account_id!r}")
                account_ids.add(account_id)

            lookup_labels = [account_label]
            lookup_labels.extend(account.get("aliases") or [])
            lookup_labels.extend(account.get("excel_bank_aliases") or [])
            owner = account_id or f"{org_code}/{account_label}/{account_no}"
            for label in lookup_labels:
                lookup_key = _receipt_lookup_key(label)
                if not lookup_key:
                    continue
                previous = account_lookup_owners.get(lookup_key)
                if previous and previous != owner:
                    errors.append(
                        f"{prefix} lookup label {label!r} conflicts with account {previous!r}"
                    )
                account_lookup_owners[lookup_key] = owner


def _validate_candidate_map(candidate_map, errors, prefix):
    if not isinstance(candidate_map, dict):
        errors.append(f"{prefix}.nc_candidates_by_currency must be an object")
        return
    for currency, candidates in candidate_map.items():
        if not isinstance(currency, str) or not currency.strip():
            errors.append(
                f"{prefix}.nc_candidates_by_currency key must be non-empty string"
            )
        if not isinstance(candidates, list) or not candidates:
            errors.append(
                f"{prefix}.nc_candidates_by_currency[{currency!r}] must be a non-empty list"
            )
            continue
        for index, candidate in enumerate(candidates):
            if not isinstance(candidate, str) or not candidate.strip():
                errors.append(
                    f"{prefix}.nc_candidates_by_currency[{currency!r}][{index}] "
                    "must be a non-empty string"
                )


def _validate_account_entry_policy(entry_policy, errors, prefix):
    if not isinstance(entry_policy, dict):
        errors.append(f"{prefix}.entry_policy must be an object")
        return
    _optional_enum(
        entry_policy,
        "account_input",
        ACCOUNT_INPUT_STRATEGIES,
        errors,
        prefix=f"{prefix}.entry_policy",
    )
    _optional_enum(
        entry_policy,
        "success_rule",
        ACCOUNT_SUCCESS_RULES,
        errors,
        prefix=f"{prefix}.entry_policy",
    )
    if "fallback_reference" in entry_policy:
        errors.append(
            f"{prefix}.entry_policy.fallback_reference is deprecated; "
            "account input must use detail_first"
        )


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


def _validate_receipt_excel_text_field_mappings(mappings, errors):
    prefix = "receipt_entry.excel_text_field_mappings"
    allowed_nc_fields = {"商务领款备忘", "备注"}
    if not isinstance(mappings, list):
        errors.append(f"{prefix} must be a list")
        return
    for index, item in enumerate(mappings):
        item_prefix = f"{prefix}[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{item_prefix} must be an object")
            continue
        _require_non_empty_str(item, "excel_column", errors, prefix=item_prefix)
        nc_field = _require_non_empty_str(item, "nc_field", errors, prefix=item_prefix)
        if nc_field and nc_field not in allowed_nc_fields:
            errors.append(
                f"{item_prefix}.nc_field must be one of {sorted(allowed_nc_fields)!r}"
            )


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
    for key in ("open_key", "main_title"):
        if key in query_cfg:
            _require_non_empty_str(query_cfg, key, errors, prefix="receipt_entry.query")
    for key in ("open_timeout", "activate_timeout", "open_wait", "result_wait"):
        if key in query_cfg:
            _non_negative_number(query_cfg, key, errors, prefix="receipt_entry.query")
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

    result_column_indexes = query_cfg.get("result_column_indexes")
    if result_column_indexes is not None:
        if not isinstance(result_column_indexes, dict):
            errors.append("receipt_entry.query.result_column_indexes must be an object")
        else:
            for key in (
                "document_no",
                "document_date",
                "original_amount",
                "customer",
                "payer_name",
            ):
                _non_negative_int(
                    result_column_indexes,
                    key,
                    errors,
                    prefix="receipt_entry.query.result_column_indexes",
                )
            if result_column_indexes.get("original_amount") == 8:
                errors.append(
                    "receipt_entry.query.result_column_indexes.original_amount "
                    "uses NC/JAB zero-based indexes; observed receipt amount column "
                    "is 6, not Excel-style 8"
                )

    page_guard = query_cfg.get("page_guard")
    if page_guard is not None:
        _validate_receipt_page_guard(page_guard, errors)

    result_guard = query_cfg.get("result_guard")
    if result_guard is not None:
        _validate_receipt_result_guard(result_guard, errors)

    pagination = query_cfg.get("pagination")
    if pagination is not None:
        if not isinstance(pagination, dict):
            errors.append("receipt_entry.query.pagination must be an object")
        else:
            _positive_int(
                pagination,
                "page_size",
                errors,
                prefix="receipt_entry.query.pagination",
            )
            for key in (
                "page_label_path",
                "page_size_text_path",
                "next_page_button_path",
                "window_class",
            ):
                _require_non_empty_str(
                    pagination, key, errors, prefix="receipt_entry.query.pagination"
                )
            for key in (
                "pager_scope_timeout",
                "wait_before_page_size",
                "wait_after_page_size",
                "wait_before_read",
                "wait_after_page_read",
                "stability_timeout",
                "stability_interval",
                "next_action_timeout",
                "wait_after_next",
            ):
                _non_negative_number(
                    pagination, key, errors, prefix="receipt_entry.query.pagination"
                )
            if "next_bounds_timeout" in pagination:
                errors.append(
                    "receipt_entry.query.pagination.next_bounds_timeout is not allowed"
                )
            _positive_int(
                pagination,
                "stability_required",
                errors,
                prefix="receipt_entry.query.pagination",
            )
            if "wait_before_page_size_stable" in pagination and not isinstance(
                pagination.get("wait_before_page_size_stable"), bool
            ):
                errors.append(
                    "receipt_entry.query.pagination.wait_before_page_size_stable must be bool"
                )

    jab_cfg = query_cfg.get("jab")
    if jab_cfg is not None:
        _validate_receipt_query_jab(jab_cfg, errors)


def _validate_receipt_page_guard(page_guard, errors):
    if not isinstance(page_guard, dict):
        errors.append("receipt_entry.query.page_guard must be an object")
        return
    if "enabled" in page_guard and not isinstance(page_guard.get("enabled"), bool):
        errors.append("receipt_entry.query.page_guard.enabled must be bool")
    if "state_label_timeout" in page_guard:
        _non_negative_number(
            page_guard,
            "state_label_timeout",
            errors,
            prefix="receipt_entry.query.page_guard",
        )
    if "state_label_require_showing" in page_guard and not isinstance(
        page_guard.get("state_label_require_showing"), bool
    ):
        errors.append(
            "receipt_entry.query.page_guard.state_label_require_showing must be bool"
        )
    if "visible_only" in page_guard and not isinstance(
        page_guard.get("visible_only"), bool
    ):
        errors.append("receipt_entry.query.page_guard.visible_only must be bool")


def _validate_receipt_result_guard(result_guard, errors):
    if not isinstance(result_guard, dict):
        errors.append("receipt_entry.query.result_guard must be an object")
        return
    if "enabled" in result_guard and not isinstance(result_guard.get("enabled"), bool):
        errors.append("receipt_entry.query.result_guard.enabled must be bool")
    if "document_type_column" in result_guard:
        _non_negative_int(
            result_guard,
            "document_type_column",
            errors,
            prefix="receipt_entry.query.result_guard",
        )
    if "document_type" in result_guard:
        _require_non_empty_str(
            result_guard,
            "document_type",
            errors,
            prefix="receipt_entry.query.result_guard",
        )
    if "max_samples" in result_guard:
        _positive_int(
            result_guard,
            "max_samples",
            errors,
            prefix="receipt_entry.query.result_guard",
        )
    if "name_column_must_not_equal_document_type" in result_guard and not isinstance(
        result_guard.get("name_column_must_not_equal_document_type"), bool
    ):
        errors.append(
            "receipt_entry.query.result_guard."
            "name_column_must_not_equal_document_type must be bool"
        )
    blocked_keywords = result_guard.get("blocked_keywords")
    if blocked_keywords is not None:
        if not isinstance(blocked_keywords, list):
            errors.append(
                "receipt_entry.query.result_guard.blocked_keywords must be a list"
            )
        else:
            for index, value in enumerate(blocked_keywords):
                if not isinstance(value, str) or not value.strip():
                    errors.append(
                        "receipt_entry.query.result_guard.blocked_keywords"
                        f"[{index}] must be a non-empty string"
                    )


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
                near_label_field=key == "finance_org",
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


def _validate_receipt_query_jab_field(
    field, errors, prefix, range_field=False, near_label_field=False
):
    for key in ("label", "operator"):
        _require_non_empty_str(field, key, errors, prefix=prefix)
    if near_label_field:
        if "text_path" in field:
            _require_non_empty_str(field, "text_path", errors, prefix=prefix)
    elif range_field:
        for key in ("from_text_path", "to_text_path"):
            _require_non_empty_str(field, key, errors, prefix=prefix)
    else:
        _require_non_empty_str(field, "text_path", errors, prefix=prefix)
    for key in ("focus_before_set", "focus_click_mode", "focus_wait", "focus_timeout"):
        if key in field:
            errors.append(f"{prefix}.{key} is not allowed")
    for key in ("timeout",):
        if key in field:
            _non_negative_number(field, key, errors, prefix=prefix)


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


def _receipt_lookup_key(value):
    return "".join(str(value or "").strip().casefold().split())


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
