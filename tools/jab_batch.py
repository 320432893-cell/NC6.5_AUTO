import argparse
import json
import sys
import time
from datetime import date

from core.paths import base_dir

ROOT = base_dir()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.data_handler import DataHandler  # noqa: E402
from core.jab_batch_processor import JABBatchProcessor  # noqa: E402
from core.logger import log  # noqa: E402
from core.utils import load_config  # noqa: E402


def build_parser():
    parser = argparse.ArgumentParser(
        description="JAB 批量生成/回填凭证号工具",
    )
    parser.add_argument(
        "command",
        choices=(
            "plan",
            "generate",
            "resume-voucher",
            "switch-generated",
            "backfill",
            "split-keys",
        ),
        help=(
            "plan=只匹配分批; generate=真实生成并保存; "
            "resume-voucher=不再点击生成，恢复保存当前制单窗口; "
            "switch-generated=自动切到已生成列表; backfill=在已生成列表回填凭证号; "
            "split-keys=把金额+对手方拼接列拆成两列"
        ),
    )
    parser.add_argument("--config", default="config.json", help="配置文件路径")
    parser.add_argument(
        "--limit", type=int, default=None, help="仅处理前 N 条 Excel 数据"
    )
    parser.add_argument(
        "--start-row",
        type=int,
        default=None,
        help="仅处理 Excel 起始行号，包含该行",
    )
    parser.add_argument(
        "--end-row",
        type=int,
        default=None,
        help="仅处理 Excel 结束行号，包含该行",
    )
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
    parser.add_argument(
        "--perf",
        action="store_true",
        help="记录性能耗时 JSONL 到 logs/perf_*.jsonl",
    )
    parser.add_argument(
        "--perf-label",
        default=None,
        help="性能日志标签，默认使用时间戳",
    )
    parser.add_argument(
        "--generated-date",
        default=None,
        help="已生成列表的目的业务日期，格式 YYYY-MM-DD；不传则使用配置或当天",
    )
    parser.add_argument(
        "--save-trigger",
        choices=("jab_button", "hotkey"),
        default=None,
        help="覆盖保存触发方式：jab_button=JAB按钮；hotkey=Ctrl+S",
    )
    parser.add_argument(
        "--save-strategy",
        choices=(
            "single",
            "bottom_up",
            "safe_batch_by_pending_row",
        ),
        default=None,
        help="覆盖制单保存策略；备选批量用 safe_batch_by_pending_row",
    )
    parser.add_argument(
        "--voucher-order-fallback",
        choices=("strict", "same_count"),
        default=None,
        help="制单表匹配兜底：same_count=制单行数等于本批数量时按行序匹配",
    )
    parser.add_argument(
        "--foreign-currency-rate",
        type=str,
        default=None,
        help="外币到本位币汇率；不传则同名多行自动估计一致汇率",
    )
    parser.add_argument(
        "--on-duplicate",
        choices=("stop", "skip"),
        default=None,
        help="generate 遇到待生成表重复匹配时的处理：stop=暂停；skip=写异常并跳过该行继续",
    )
    parser.add_argument(
        "--hotkey-activate-policy",
        choices=("always", "first", "foreground_guard"),
        default=None,
        help="Ctrl+S 保存前窗口处理：always=每张激活；first=仅首张激活；foreground_guard=只校验前台制单",
    )
    parser.add_argument(
        "--no-backfill-auto-switch",
        action="store_true",
        help="backfill 时不从待生成页自动切到已生成列表，只做状态校验",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="在 stdout 最后一行输出结构化结果信封 JSON（供 GUI 外壳解析）",
    )
    return parser


def main():  # noqa: C901 (intentional single-function dispatch)
    args = build_parser().parse_args()
    if args.generated_date is not None:
        try:
            date.fromisoformat(args.generated_date)
        except ValueError as exc:
            raise SystemExit(
                f"--generated-date 格式必须是 YYYY-MM-DD: {args.generated_date!r}"
            ) from exc
    cfg = load_config(args.config)
    if args.save_strategy is not None:
        cfg.setdefault("jab_batch", {})["save_strategy"] = args.save_strategy
    if args.voucher_order_fallback is not None:
        cfg.setdefault("jab_batch", {})["voucher_order_fallback_mode"] = (
            args.voucher_order_fallback
        )
    if args.foreign_currency_rate is not None:
        cfg.setdefault("jab_batch", {})["foreign_currency_rate"] = (
            args.foreign_currency_rate
        )
    if args.on_duplicate is not None:
        cfg.setdefault("jab_batch", {})["duplicate_match_policy"] = args.on_duplicate
    processor = JABBatchProcessor(
        cfg,
        perf_enabled=args.perf,
        perf_label=args.perf_label,
        command=args.command,
        generated_date_value=args.generated_date,
        save_trigger=args.save_trigger,
        hotkey_activate_policy=args.hotkey_activate_policy,
    )
    state_finished = False
    t_start = time.monotonic()

    # 信封构建辅助 —— 只在 --json 模式下调用，不改任何业务分支
    def _emit_envelope(
        *,
        ok: bool,
        exit_code: int,
        summary: dict,
        items: list,
        error_category: str = "none",
        error_message: str = "",
        can_resume: bool = False,
        resume_command: "str | None" = None,
    ) -> None:
        envelope = {
            "ok": ok,
            "command": args.command,
            "exit_code": exit_code,
            "summary": summary,
            "items": items,
            "error": {"category": error_category, "message": error_message},
            "elapsed_s": round(time.monotonic() - t_start, 3),
            "resumable": {"can_resume": can_resume, "resume_command": resume_command},
        }
        print(json.dumps(envelope, ensure_ascii=False))

    try:
        if args.command == "plan":
            result = processor.dry_run(
                limit=args.limit,
                start_row=args.start_row,
                end_row=args.end_row,
            )
            print_summary(result)
            if args.json_output:
                matches = result["matches"]
                issues = result["issues"]
                parse_errors = result["parse_errors"]
                failed_count = len(issues) + len(parse_errors)
                total = len(matches) + failed_count
                items = []
                for pe in parse_errors:
                    items.append(
                        {
                            "ref": f"Excel行{pe.row}",
                            "outcome": "failed",
                            "reason": f"格式错误: {pe.parse_error}",
                        }
                    )
                for iss in issues:
                    items.append(
                        {
                            "ref": f"Excel行{iss.item.row}",
                            "outcome": "failed",
                            "reason": iss.reason,
                        }
                    )
                for m in matches:
                    items.append(
                        {
                            "ref": f"Excel行{m.item.row}",
                            "outcome": "success",
                            "reason": "",
                        }
                    )
                exit_code = 0 if failed_count == 0 else (2 if len(matches) == 0 else 1)
                _emit_envelope(
                    ok=(exit_code == 0),
                    exit_code=exit_code,
                    summary={
                        "total": total,
                        "succeeded": len(matches),
                        "failed": failed_count,
                        "skipped": 0,
                    },
                    items=items,
                )
            return

        if args.command == "generate":
            if not args.yes:
                print("即将真实点击 NC：多选行、生成、逐张保存。")
                answer = input("确认继续请输入 yes: ").strip().lower()
                if answer != "yes":
                    log.warning("用户取消 generate")
                    processor.run_state.set_stage("generate_cancelled")
                    processor.finish_run_state("cancelled")
                    state_finished = True
                    if args.json_output:
                        _emit_envelope(
                            ok=False,
                            exit_code=3,
                            summary={
                                "total": 0,
                                "succeeded": 0,
                                "failed": 0,
                                "skipped": 0,
                            },
                            items=[],
                            error_category="aborted",
                            error_message="用户取消",
                        )
                    return
            saved = processor.generate_and_save(
                limit=args.limit,
                max_batches=args.max_batches,
                start_row=args.start_row,
                end_row=args.end_row,
            )
            print(f"生成保存完成: {saved} 张")
            if args.json_output:
                _emit_envelope(
                    ok=True,
                    exit_code=0,
                    summary={
                        "total": saved,
                        "succeeded": saved,
                        "failed": 0,
                        "skipped": 0,
                    },
                    items=[],
                    can_resume=True,
                    resume_command="resume-voucher",
                )
            return

        if args.command == "resume-voucher":
            saved = processor.resume_current_voucher_window(
                limit=args.limit,
                start_row=args.start_row,
                end_row=args.end_row,
            )
            print(f"恢复制单窗口保存完成: {saved} 张")
            if args.json_output:
                _emit_envelope(
                    ok=True,
                    exit_code=0,
                    summary={
                        "total": saved,
                        "succeeded": saved,
                        "failed": 0,
                        "skipped": 0,
                    },
                    items=[],
                )
            return

        if args.command == "switch-generated":
            processor.switch_to_generated_list()
            print("已切换到已生成列表")
            if args.json_output:
                _emit_envelope(
                    ok=True,
                    exit_code=0,
                    summary={"total": 1, "succeeded": 1, "failed": 0, "skipped": 0},
                    items=[
                        {"ref": "switch-generated", "outcome": "success", "reason": ""}
                    ],
                )
            return

        if args.command == "backfill":
            updates = processor.backfill_generated_vouchers(
                limit=args.limit,
                start_row=args.start_row,
                end_row=args.end_row,
                auto_switch=not args.no_backfill_auto_switch,
            )
            print(f"回填完成: {len(updates)} 行")
            if args.json_output:
                items = [
                    {"ref": str(ref), "outcome": "success", "reason": ""}
                    for ref in updates
                ]
                succeeded = len(updates)
                _emit_envelope(
                    ok=True,
                    exit_code=0,
                    summary={
                        "total": succeeded,
                        "succeeded": succeeded,
                        "failed": 0,
                        "skipped": 0,
                    },
                    items=items,
                )
            return

        if args.command == "split-keys":
            result = DataHandler(cfg).split_jab_keys_to_columns(limit=args.limit)
            print(
                "拆分完成: "
                f"{result['updates']} 行, "
                f"金额列={result['amount_col']}, 对手方列={result['partner_col']}, "
                f"错误={len(result['errors'])} 行"
            )
            if args.json_output:
                errors = result.get("errors", [])
                succeeded = int(result.get("updates", 0))
                failed = len(errors)
                total = succeeded + failed
                items = [
                    {"ref": str(e), "outcome": "failed", "reason": str(e)}
                    for e in errors
                ]
                exit_code = 0 if failed == 0 else 1
                _emit_envelope(
                    ok=(exit_code == 0),
                    exit_code=exit_code,
                    summary={
                        "total": total,
                        "succeeded": succeeded,
                        "failed": failed,
                        "skipped": 0,
                    },
                    items=items,
                )
            return
    except BaseException as exc:
        status = "aborted" if isinstance(exc, SystemExit) else "failed"
        processor.finish_run_state(status, error=f"{type(exc).__name__}: {exc}")
        state_finished = True
        if args.json_output:
            is_aborted = isinstance(exc, KeyboardInterrupt) or (
                isinstance(exc, SystemExit) and getattr(exc, "code", 0) == 3
            )
            _emit_envelope(
                ok=False,
                exit_code=3 if is_aborted else 4,
                summary={"total": 0, "succeeded": 0, "failed": 0, "skipped": 0},
                items=[],
                error_category="aborted" if is_aborted else "environment",
                error_message=f"{type(exc).__name__}: {exc}",
                can_resume=(args.command == "generate"),
                resume_command=(
                    "resume-voucher" if args.command == "generate" else None
                ),
            )
        raise
    finally:
        if not state_finished:
            processor.finish_run_state("success")
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
            print(f"- Excel行{item.row}: 格式错误 {item.parse_error}")
        for issue in issues[:20]:
            rows = issue.rows
            print(
                f"- Excel行{issue.item.row}: {issue.reason} "
                f"amount={issue.item.amount} partner={issue.item.partner} "
                f"nc_rows={rows}"
            )
        if len(issues) + len(parse_errors) > 20:
            print("- ...更多问题行见日志")

    if batches:
        print("\n批次:")
        for index, batch in enumerate(batches[:20], start=1):
            print(
                f"- 批次{index}: {len(batch)}条 "
                f"Excel行={[m.item.row for m in batch]} "
                f"NC行={[m.nc_row for m in batch]}"
            )
        if len(batches) > 20:
            print("- ...更多批次见日志")


if __name__ == "__main__":
    main()
