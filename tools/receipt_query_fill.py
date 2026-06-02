import argparse
from datetime import date
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.receipt_entry import parse_date  # noqa: E402
from core.utils import load_config  # noqa: E402


def resolve_today(value):
    return date.today().isoformat() if value == "{today}" else value


def set_text(jab, jab_cfg, path, value):
    return jab.set_text_by_path(
        path,
        value,
        title=jab_cfg["dialog_title"],
        class_name=jab_cfg["dialog_class"],
        role="text",
        timeout=2,
        require_showing=True,
    )


def fill_receipt_query(config, org_code, date_from=None, date_to=None, confirm=False):
    query_cfg = config["receipt_entry"]["query"]
    jab_cfg = query_cfg["jab"]
    fields = jab_cfg["fields"]
    start = date_from or query_cfg["date_from"]
    end = date_to or query_cfg["date_to"]
    start = parse_date(resolve_today(start)).isoformat()
    end = parse_date(resolve_today(end)).isoformat()

    jab = JABOperator(config)
    try:
        steps = [
            (
                "finance_org",
                fields["finance_org"]["text_path"],
                org_code,
            ),
            (
                "document_date_from",
                fields["document_date"]["from_text_path"],
                start,
            ),
            (
                "document_date_to",
                fields["document_date"]["to_text_path"],
                end,
            ),
        ]
        for name, path, value in steps:
            if not set_text(jab, jab_cfg, path, value):
                raise RuntimeError(f"收款查询条件写入失败: {name}={value}")

        if confirm:
            ok = jab.do_action_by_path(
                jab_cfg["confirm_button_path"],
                title=jab_cfg["dialog_title"],
                class_name=jab_cfg["dialog_class"],
                name="确定(Y)",
                role="push button",
                action_name="单击",
                wait=1,
                timeout=2,
                require_showing=True,
            )
            if not ok:
                raise RuntimeError("收款查询确定按钮点击失败")
        return {"organization_code": org_code, "date_from": start, "date_to": end}
    finally:
        jab.close()


def main():
    parser = argparse.ArgumentParser(description="Fill NC receipt query conditions")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--org-code", required=True)
    parser.add_argument("--date-from", default=None)
    parser.add_argument("--date-to", default=None)
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="click query confirm after filling conditions",
    )
    args = parser.parse_args()

    result = fill_receipt_query(
        load_config(args.config),
        org_code=args.org_code,
        date_from=args.date_from,
        date_to=args.date_to,
        confirm=args.confirm,
    )
    print(
        "filled receipt query: "
        f"org={result['organization_code']} "
        f"date_from={result['date_from']} date_to={result['date_to']} "
        f"confirm={args.confirm}"
    )


if __name__ == "__main__":
    raise SystemExit(main())
