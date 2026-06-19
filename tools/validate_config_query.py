"""Receipt-entry query sub-tree validation.

Boundary: this module owns the `receipt_entry.query` sub-tree validation
(query fields/result columns/pagination/page_guard/result_guard/jab.fields).
It is called by the top-level dispatcher in `validate_config` and only depends
on the shared primitives in `validate_config_primitives` to avoid an import
cycle with the dispatcher. Error messages and append order here are part of the
single validation contract and must stay byte-for-byte stable.
"""

from tools.validate_config_primitives import (
    _non_negative_int,
    _non_negative_number,
    _positive_int,
    _require,
    _require_non_empty_str,
    _iso_date,
)


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
