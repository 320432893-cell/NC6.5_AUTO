"""Shared validation primitives for config validation modules.

Boundary: this module holds the low-level, dependency-free validation helpers
(`_require`/`_enum`/`_iso_date`/...) shared by `validate_config`,
`validate_config_query`, and `validate_config_accounts`. It must not import any
of those modules, so it sits at the bottom of the import graph and breaks the
cycle that would otherwise form between the dispatcher and the sub-validators.
"""

from datetime import date


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
