import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_batch_processor import JABBatchProcessor
from core.data_handler import DataHandler
from core.logger import log
from core.utils import load_config


def build_parser():
    parser = argparse.ArgumentParser(
        description="JAB 批量生成/回填凭证号工具",
    )
    parser.add_argument(
        "command",
        choices=("plan", "generate", "switch-generated", "backfill", "split-keys"),
        help=(
            "plan=只匹配分批; generate=真实生成并保存; "
            "switch-generated=自动切到已生成列表; backfill=在已生成列表回填凭证号; "
            "split-keys=把金额+对手方拼接列拆成两列"
        ),
    )
    parser.add_argument("--config", default="config.json", help="配置文件路径")
    parser.add_argument("--limit", type=int, default=None, help="仅处理前 N 条 Excel 数据")
    parser.add_argument(
        "--max-batches",
        type=int,
        default=None,
        help="generate 时最多执行几个批次",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="generate 时跳过二次确认",
    )
    return parser


def main():
    args = build_parser().parse_args()
    cfg = load_config(args.config)
    processor = JABBatchProcessor(cfg)

    try:
        if args.command == "plan":
            result = processor.dry_run(limit=args.limit)
            print_summary(result)
            return

        if args.command == "generate":
            if not args.yes:
                print("即将真实点击 NC：多选行、生成、逐张保存。")
                answer = input("确认继续请输入 yes: ").strip().lower()
                if answer != "yes":
                    log.warning("用户取消 generate")
                    return
            saved = processor.generate_and_save(
                limit=args.limit,
                max_batches=args.max_batches,
            )
            print(f"生成保存完成: {saved} 张")
            return

        if args.command == "switch-generated":
            processor.switch_to_generated_list()
            print("已切换到已生成列表")
            return

        if args.command == "backfill":
            updates = processor.backfill_generated_vouchers(limit=args.limit)
            print(f"回填完成: {len(updates)} 行")
            return

        if args.command == "split-keys":
            result = DataHandler(cfg).split_jab_keys_to_columns(limit=args.limit)
            print(
                "拆分完成: "
                f"{result['updates']} 行, "
                f"金额列={result['amount_col']}, 对手方列={result['partner_col']}, "
                f"错误={len(result['errors'])} 行"
            )
            return
    finally:
        processor.close()


def print_summary(result):
    matches = result["matches"]
    issues = result["issues"]
    batches = result["batches"]
    parse_errors = result["parse_errors"]

    print("\nJAB 批量计划")
    print(f"可匹配: {len(matches)}")
    print(f"格式错误: {len(parse_errors)}")
    print(f"未找到/重复: {len(issues)}")
    print(f"批次数: {len(batches)}")

    if issues or parse_errors:
        print("\n问题行:")
        for item in parse_errors[:20]:
            print(f"- Excel行{item['row']}: 格式错误 {item.get('parse_error')}")
        for issue in issues[:20]:
            rows = issue.get("rows", [])
            print(
                f"- Excel行{issue['item']['row']}: {issue['reason']} "
                f"amount={issue['item']['amount']} partner={issue['item']['partner']} "
                f"nc_rows={rows}"
            )
        if len(issues) + len(parse_errors) > 20:
            print("- ...更多问题行见日志")

    if batches:
        print("\n批次:")
        for index, batch in enumerate(batches[:20], start=1):
            print(
                f"- 批次{index}: {len(batch)}条 "
                f"Excel行={[m['item']['row'] for m in batch]} "
                f"NC行={[m['nc_row'] for m in batch]}"
            )
        if len(batches) > 20:
            print("- ...更多批次见日志")


if __name__ == "__main__":
    main()
