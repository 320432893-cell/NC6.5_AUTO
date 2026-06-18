from tools.validate_config import validate_config


def base_config():
    return {
        "excel_path": "unused.xlsx",
        "sheet_my": "Sheet1",
        "jab": {
            "dll_path": "unused.dll",
            "startup_wait": 0.5,
            "search_timeout": 5.0,
            "max_depth": 25,
            "max_children": 1000,
            "amount_col": 4,
            "partner_col": 3,
            "selection_col": 0,
        },
        "jab_batch": {
            "key_col": 1,
            "amount_out_col": 1,
            "partner_out_col": 2,
            "result_col": 3,
            "generated_voucher_col": 22,
            "generated_date_col": 18,
            "save_strategy": "single",
            "save_trigger": "jab_button",
            "hotkey_activate_policy": "always",
            "duplicate_match_policy": "stop",
            "max_batch_size": 50,
            "save_wait": 0.5,
            "wait_between_save_batches": 0.0,
            "save_success_timeout": 8.0,
            "voucher_record_timeout": 8.0,
            "generated_voucher_max": 9999,
            "generated_date_value": "2026-05-31",
            "open_query": {
                "method": "hotkey",
                "key": "f3",
                "dialog_title": "查询",
                "dialog_class": "SunAwtDialog",
                "timeout": 5,
                "activate_timeout": 1,
                "process_timeout": 4,
            },
            "switch_generated_steps": [
                {
                    "path": "0",
                    "action": "单击",
                    "timeout": 1,
                    "guard_timeout": 1,
                    "wait": 0,
                },
            ],
        },
    }


def test_validate_receipt_entry_accepts_account_mapping():
    config = base_config()
    config["receipt_entry"] = {
        "state_label": "收款单录入",
        "excel": {
            "path": "payments.xlsx",
            "sheet_name": "💸Payments来款通知",
            "header_row": 1,
            "start_row": 2,
            "result_sheet_name": "收款单自动化结果",
            "start_date": "2026-01-01",
            "date_column": "到款日期",
            "payer_name_column": "🟪银行来款名",
            "raw_amount_column": "🟪原始金额",
            "bank_column": "银行",
            "currency_column": "币种",
            "customer_code_column": "客户编码",
            "fee_column": "手续费",
            "organization_column": "主体名称",
        },
        "validation_policy": {
            "mode": "strict",
            "skip_invalid_rows": False,
        },
        "query": {
            "date_from": "2026-01-01",
            "date_to": "{today}",
            "open_key": "f3",
            "main_title": "Yonyou UClient",
            "main_class": "YonyouUWnd",
            "open_timeout": 5,
            "activate_timeout": 5,
            "open_wait": 0.8,
            "finance_org_field": "收款财务组织",
            "finance_org_operator": "等于",
            "document_date_field": "单据日期",
            "document_date_operator": "介于",
            "jab": {
                "dialog_title": "查询条件",
                "dialog_class": "SunAwtDialog",
                "confirm_button_path": "0.0.1.0.2.1.0",
                "fields": {
                    "finance_org": {
                        "label": "收款财务组织",
                        "operator": "等于",
                        "timeout": 2.0,
                    },
                    "document_date": {
                        "label": "单据日期",
                        "operator": "介于",
                        "from_text_path": "0.0.1",
                        "to_text_path": "0.0.2",
                    },
                },
            },
            "result_columns": {
                "document_date": "单据日期",
                "original_amount": "原币金额",
                "customer": "客户",
            },
            "result_column_indexes": {
                "document_no": 0,
                "document_date": 1,
                "customer": 4,
                "original_amount": 6,
                "payer_name": 4,
            },
        },
        "candidate_check": {
            "recent_months": 2,
            "from_date": None,
        },
        "finance_organizations": [
            {
                "code": "A001",
                "name": "上海移为通信技术股份有限公司",
                "short_name": "移为",
            },
        ],
        "accounts": [
            {
                "organization_code": "A001",
                "organization_short_name": "移为",
                "account_label": "大陆花旗",
                "account_no": "1783854003",
                "header_currency_code": "USD",
            },
        ],
    }

    assert validate_config(config) == []


def test_validate_receipt_entry_accepts_extensible_bank_account_schema():
    config = base_config()
    config["receipt_entry"] = {
        "schema_version": 2,
        "state_label": "收款单录入",
        "finance_organizations": [
            {
                "code": "A001",
                "name": "上海移为通信技术股份有限公司",
                "short_name": "移为",
            },
        ],
        "banks": [
            {"id": "cmb", "name": "招商银行", "aliases": ["招行"], "enabled": True},
        ],
        "detail_entry_policy": {
            "main_line_order": [
                "business_type",
                "account",
                "subject",
                "amount",
                "settlement",
            ],
            "fee_line_order": ["business_type", "subject", "amount", "settlement"],
            "fee_add_row_hotkey": "ctrl+i",
            "fee_clear_account": True,
            "extra_blank_row_delete_hotkey": "ctrl+d",
        },
        "accounts": [
            {
                "id": "cmb_a001",
                "enabled": True,
                "organization_code": "A001",
                "organization_short_name": "移为",
                "bank_id": "cmb",
                "display_name": "移为-招行",
                "account_label": "大陆招行",
                "account_no": "FTE1219165931831",
                "header_currency_code": "USD",
                "excel_bank_aliases": ["招商"],
                "nc_candidates_by_currency": {
                    "人民币": ["FTE1219165931831RMB", "FTE1219165931831"],
                    "美元": ["FTE1219165931831USD", "FTE1219165931831"],
                },
                "entry_policy": {
                    "account_input": "detail_first",
                    "success_rule": "non_empty",
                },
            },
        ],
    }

    assert validate_config(config) == []


def test_validate_receipt_entry_rejects_conflicting_account_aliases():
    config = base_config()
    config["receipt_entry"] = {
        "state_label": "收款单录入",
        "finance_organizations": [
            {
                "code": "A001",
                "name": "上海移为通信技术股份有限公司",
                "short_name": "移为",
            },
        ],
        "accounts": [
            {
                "id": "first",
                "organization_code": "A001",
                "organization_short_name": "移为",
                "account_label": "招行",
                "account_no": "1",
                "header_currency_code": "USD",
            },
            {
                "id": "second",
                "organization_code": "A001",
                "organization_short_name": "移为",
                "account_label": "招商银行",
                "account_no": "2",
                "header_currency_code": "USD",
                "excel_bank_aliases": ["招行"],
            },
        ],
    }

    assert validate_config(config) == [
        "receipt_entry.accounts[1] lookup label '招行' conflicts with account 'first'"
    ]


def test_validate_receipt_entry_rejects_one_based_nc_amount_column():
    config = base_config()
    config["receipt_entry"] = {
        "state_label": "收款单录入",
        "query": {
            "date_from": "2026-01-01",
            "date_to": "{today}",
            "finance_org_field": "收款财务组织",
            "finance_org_operator": "等于",
            "document_date_field": "单据日期",
            "document_date_operator": "介于",
            "jab": {
                "dialog_title": "查询条件",
                "dialog_class": "SunAwtDialog",
                "confirm_button_path": "0.0.1.0.2.1.0",
                "fields": {
                    "finance_org": {
                        "label": "收款财务组织",
                        "operator": "等于",
                        "text_path": "0.0.1.0.1.0.0.1.0.0.0.0.0.1.0.1.1.2.0.0.0.0",
                    },
                    "document_date": {
                        "label": "单据日期",
                        "operator": "介于",
                        "from_text_path": "0.0.1",
                        "to_text_path": "0.0.2",
                    },
                },
            },
            "result_columns": {
                "document_date": "单据日期",
                "original_amount": "原币金额",
                "customer": "客户",
            },
            "result_column_indexes": {
                "document_no": 0,
                "document_date": 1,
                "customer": 2,
                "original_amount": 8,
                "payer_name": 2,
            },
        },
    }

    assert (
        "receipt_entry.query.result_column_indexes.original_amount "
        "uses NC/JAB zero-based indexes; observed receipt amount column "
        "is 6, not Excel-style 8"
    ) in validate_config(config)


def test_validate_receipt_entry_allows_finance_org_text_path():
    config = base_config()
    config["receipt_entry"] = {
        "state_label": "收款单录入",
        "query": {
            "date_from": "2026-01-01",
            "date_to": "{today}",
            "finance_org_field": "收款财务组织",
            "finance_org_operator": "等于",
            "document_date_field": "单据日期",
            "document_date_operator": "介于",
            "jab": {
                "dialog_title": "查询条件",
                "dialog_class": "SunAwtDialog",
                "confirm_button_path": "0.0.1.0.2.1.0",
                "fields": {
                    "finance_org": {
                        "label": "收款财务组织",
                        "operator": "等于",
                        "text_path": "0.0.0",
                    },
                    "document_date": {
                        "label": "单据日期",
                        "operator": "介于",
                        "from_text_path": "0.0.1",
                        "to_text_path": "0.0.2",
                    },
                },
            },
            "result_columns": {
                "document_date": "单据日期",
                "original_amount": "原币金额",
                "customer": "客户",
            },
        },
        "finance_organizations": [
            {
                "code": "A001",
                "name": "上海移为通信技术股份有限公司",
                "short_name": "移为",
            },
        ],
        "accounts": [
            {
                "organization_code": "A001",
                "organization_short_name": "移为",
                "account_label": "大陆花旗",
                "account_no": "1783854003",
                "header_currency_code": "USD",
            },
        ],
    }

    assert validate_config(config) == []


def test_validate_receipt_entry_rejects_unknown_account_organization():
    config = base_config()
    config["receipt_entry"] = {
        "state_label": "收款单录入",
        "finance_organizations": [
            {
                "code": "A001",
                "name": "上海移为通信技术股份有限公司",
                "short_name": "移为",
            },
        ],
        "accounts": [
            {
                "organization_code": "A003",
                "organization_short_name": "移航",
                "account_label": "移航招行",
                "account_no": "755927177210901",
                "header_currency_code": "USD",
            },
        ],
    }

    assert validate_config(config) == [
        "receipt_entry.accounts[0].organization_code must reference "
        "finance_organizations, got 'A003'"
    ]


def test_validate_receipt_entry_rejects_invalid_start_row_and_policy():
    config = base_config()
    config["receipt_entry"] = {
        "state_label": "收款单录入",
        "excel": {
            "path": "payments.xlsx",
            "sheet_name": "💸Payments来款通知",
            "header_row": 1,
            "start_row": 1,
            "start_date": "2026-01-01",
            "date_column": "到款日期",
            "payer_name_column": "🟪银行来款名",
            "raw_amount_column": "🟪原始金额",
            "bank_column": "银行",
            "currency_column": "币种",
            "customer_code_column": "客户编码",
            "organization_column": "主体名称",
        },
        "validation_policy": {
            "mode": "loose",
            "skip_invalid_rows": "yes",
        },
        "finance_organizations": [],
        "accounts": [],
    }

    errors = validate_config(config)

    assert "receipt_entry.excel.start_row must be greater than header_row" in errors
    assert (
        "receipt_entry.validation_policy.mode must be one of "
        "['skip_invalid_rows', 'strict'], got 'loose'"
    ) in errors
    assert "receipt_entry.validation_policy.skip_invalid_rows must be bool" in errors
