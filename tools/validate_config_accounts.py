"""Receipt-entry finance orgs / banks / accounts cross-validation.

Boundary: this module owns the uniqueness + foreign-key + lookup-conflict
cross-validation over `receipt_entry.finance_organizations`, `.banks`, and
`.accounts`. It is invoked by the top-level dispatcher in `validate_config`
after that dispatcher has resolved the three lists (so the `is required`/type
errors keep their original position). It depends only on the shared primitives
in `validate_config_primitives` to avoid an import cycle with the dispatcher.
Error messages and append order are part of the single validation contract and
must stay byte-for-byte stable.
"""

from tools.validate_config_primitives import (
    _optional_enum,
    _receipt_lookup_key,
    _require_non_empty_str,
)

ACCOUNT_INPUT_STRATEGIES = {"detail_first", "reference_first"}
ACCOUNT_SUCCESS_RULES = {"non_empty", "exact_or_candidate"}


def _validate_receipt_accounts(organizations, banks, accounts, errors):
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
