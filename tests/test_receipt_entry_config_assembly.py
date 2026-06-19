# 生命周期：持久维护
# 覆盖的业务场景：收款单录入的配置装配：银行/账户别名映射到财务组织与候选
# 依赖的服务/环境：本地 Python，不依赖 NC/GUI/JAB
# 运行方式：.venv/bin/python -m pytest -q tests/test_receipt_entry_config_assembly.py


from tests._receipt_entry_helpers import (
    ReceiptEntryConfig,
    receipt_config,
)


def test_bank_label_maps_to_organization_case_insensitive():
    config = ReceiptEntryConfig(receipt_config())

    organization = config.organization_for_bank("Paypal")

    assert organization is not None
    assert organization.code == "A001"
    assert organization.name == "上海移为通信技术股份有限公司"


def test_extended_account_alias_maps_to_account_and_candidates():
    raw = receipt_config()
    receipt = raw["receipt_entry"]
    receipt["schema_version"] = 2
    receipt["banks"] = [
        {"id": "cmb", "name": "招商银行", "aliases": ["招行"]},
    ]
    receipt["accounts"].append(
        {
            "id": "cmb_a001",
            "enabled": True,
            "organization_code": "A001",
            "organization_short_name": "移为",
            "bank_id": "cmb",
            "account_label": "大陆招行",
            "account_no": "FTE1219165931831",
            "excel_bank_aliases": ["招商", "招行"],
            "nc_candidates_by_currency": {
                "人民币": ["FTE1219165931831RMB"],
                "*": ["FTE1219165931831"],
            },
            "entry_policy": {
                "account_input": "detail_first",
                "success_rule": "non_empty",
            },
        }
    )
    config = ReceiptEntryConfig(raw)

    account = config.account_for_bank("招行")

    assert account is not None
    assert account.id == "cmb_a001"
    assert config.organization_for_bank("招商").code == "A001"
    assert account.nc_candidates("人民币") == [
        "FTE1219165931831RMB",
        "FTE1219165931831",
    ]
