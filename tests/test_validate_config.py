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
            "start_date": "2026-01-01",
            "date_column": "到款日期",
            "payer_name_column": "🟪银行来款名",
            "raw_amount_column": "🟪原始金额",
            "bank_column": "银行",
            "organization_column": "主体名称",
            "nc_done_column": "是否NC已做过",
        },
        "query": {
            "date_from": "2026-01-01",
            "date_to": "{today}",
            "finance_org_field": "收款财务组织",
            "finance_org_operator": "等于",
            "document_date_field": "单据日期",
            "document_date_operator": "介于",
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
            },
        ],
    }

    assert validate_config(config) == [
        "receipt_entry.accounts[0].organization_code must reference "
        "finance_organizations, got 'A003'"
    ]
