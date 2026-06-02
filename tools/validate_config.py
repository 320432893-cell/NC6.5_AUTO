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

    return errors


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
