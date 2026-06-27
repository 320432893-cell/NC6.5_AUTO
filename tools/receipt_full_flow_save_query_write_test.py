# 职责：现场测试入口，一个文件选择收款单保存/不保存/故障恢复/verify 审查测试
# 不做什么：不复制业务流程逻辑，不绕过 receipt_full_flow_entry.py 的保存确认

from __future__ import annotations

import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_environment import prepare_java_access_bridge  # noqa: E402
from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.receipt_modal_guard import (  # noqa: E402
    collect_visible_java_dialogs,
    recover_cancelable_modal_now,
)
from tools.receipt_full_flow_entry import main as run_full_flow  # noqa: E402
from tools.receipt_full_flow_test_prompts import (  # noqa: E402
    build_interactive_args,
    with_default_excel_path,
)


CONTROLLED_FLAGS = {
    "--save",
    "--query-after-save",
    "--write-plan-sheet",
    "--write-selected-plan-sheet",
    "--pause-after-header-field",
    "--diagnose-header-after-pause",
    "--diagnose-detail-repair",
    "--json",
}

MODE_SAVE_QUERY_WRITE = "1"
MODE_NO_SAVE_VERIFY = "2"
MODE_RECOVERY = "3"
MODE_VERIFY_AUDIT = "4"
MODE_MODAL_ALT_C = "5"
MODE_DETAIL_REPAIR = "6"
MODE_LABELS = {
    MODE_SAVE_QUERY_WRITE: "保存 + 后验查询 + 写 Sheet2",
    MODE_NO_SAVE_VERIFY: "不保存，只跑到保存前并执行 verifier",
    MODE_RECOVERY: "故障恢复诊断：客户后暂停，人工干扰后由失败动作触发 Alt+C 恢复",
    MODE_VERIFY_AUDIT: "verify 审查：不保存，重点观察后台 verifier 和最终报告",
    MODE_MODAL_ALT_C: "只测试已打开 Java 弹窗的 Alt+C 取消恢复",
    MODE_DETAIL_REPAIR: "明细 verifier 自救演练：强制一次 pending 后用缓存 path 修复",
}


def main(argv: list[str] | None = None) -> int:
    user_args = list(sys.argv[1:] if argv is None else argv)
    if not user_args:
        mode = ask_mode()
        user_args = build_interactive_args(
            MODE_LABELS[mode],
            default_start_row="811",
            default_limit="3",
            default_start_delay="2",
        )
    else:
        mode = MODE_SAVE_QUERY_WRITE
    controlled = [arg for arg in user_args if arg in CONTROLLED_FLAGS]
    if controlled:
        print(
            "本脚本由交互模式统一决定保存、查询、写 Sheet2 和诊断参数。"
            f"请移除这些参数: {', '.join(controlled)}"
        )
        return 2
    os.chdir(ROOT)
    jab_ready = prepare_java_access_bridge()
    if not jab_ready.get("ok"):
        print(jab_ready.get("reason"))
        return 4
    if mode == MODE_MODAL_ALT_C:
        return run_modal_alt_c_test(with_default_config(user_args))
    prepared_args = with_default_excel_path(with_default_config(user_args), ROOT)
    return int(run_full_flow([*prepared_args, *mode_flags(mode)]) or 0)


def ask_mode() -> str:
    print()
    print("请选择测试功能，直接回车默认 1：")
    print("1. 保存 + 后验查询 + 写 Sheet2")
    print("2. 不保存，只跑到保存前并执行 verifier")
    print("3. 故障恢复诊断：客户后暂停，人工干扰后由失败动作触发 Alt+C 恢复")
    print("4. verify 审查：不保存，重点观察后台 verifier 和最终报告")
    print("5. 只测试已打开 Java 弹窗的 Alt+C 取消恢复")
    print("6. 明细 verifier 自救演练：强制一次 pending 后用缓存 path 修复")
    value = input("测试功能 [1/2/3/4/5/6]，默认 1: ").strip()
    if value.lower() in {"q", "quit", "exit"}:
        raise SystemExit("用户取消")
    value = value or MODE_SAVE_QUERY_WRITE
    if value not in MODE_LABELS:
        raise SystemExit(f"未知测试功能: {value}")
    return value


def mode_flags(mode: str) -> list[str]:
    if mode == MODE_SAVE_QUERY_WRITE:
        return ["--save", "--query-after-save", "--write-selected-plan-sheet"]
    if mode == MODE_NO_SAVE_VERIFY:
        return []
    if mode == MODE_RECOVERY:
        return ["--pause-after-header-field", "客户", "--diagnose-header-after-pause"]
    if mode == MODE_VERIFY_AUDIT:
        return []
    if mode == MODE_MODAL_ALT_C:
        return []
    if mode == MODE_DETAIL_REPAIR:
        return ["--diagnose-detail-repair"]
    raise ValueError(f"unknown mode: {mode}")


def with_default_config(argv: list[str]) -> list[str]:
    if "--config" in argv:
        return argv
    return ["--config", str(ROOT / "config.json"), *argv]


def run_modal_alt_c_test(argv: list[str]) -> int:
    config_path = config_path_from_args(argv)
    config = load_config(config_path)
    jab = JABOperator(config)
    try:
        jab.ensure_started()
        before = collect_visible_java_dialogs(jab)
        result = recover_cancelable_modal_now(jab, stage="modal-alt-c-test")
        after = collect_visible_java_dialogs(jab)
    finally:
        jab.close()
    print_modal_alt_c_summary(before, result, after)
    return 0 if result.get("ok") else 1


def config_path_from_args(argv: list[str]) -> str:
    for index, item in enumerate(argv):
        if item == "--config" and index + 1 < len(argv):
            return argv[index + 1]
    return str(ROOT / "config.json")


def print_modal_alt_c_summary(before, result, after) -> None:
    before_recoverable = [item for item in before if item.get("cancel_controls")]
    after_recoverable = [item for item in after if item.get("cancel_controls")]
    print("Java 弹窗 Alt+C 恢复测试")
    print(f"结果：{'成功' if result.get('ok') else '失败'}")
    print(f"恢复动作：attempted={result.get('attempted')}")
    print(f"恢复原因：{result.get('reason') or ''}")
    print(
        "弹窗数量："
        f"before={len(before)}, before_cancelable={len(before_recoverable)}, "
        f"after={len(after)}, after_cancelable={len(after_recoverable)}"
    )
    focus = result.get("focus") or {}
    if focus:
        print(
            "聚焦："
            f"ok={focus.get('ok')}, hwnd={focus.get('hwnd')}, "
            f"foreground_before={focus.get('foreground_before')}, "
            f"foreground_after={focus.get('foreground_after')}"
        )
    if before:
        print("恢复前弹窗：")
        for item in before[:5]:
            print(format_dialog_line(item))
    if after:
        print("恢复后弹窗：")
        for item in after[:5]:
            print(format_dialog_line(item))


def format_dialog_line(item) -> str:
    controls = item.get("cancel_controls") or []
    buttons = item.get("buttons") or []
    first_cancel = controls[0] if controls else {}
    return (
        f"- hwnd={item.get('hwnd')} title={item.get('title')!r} "
        f"cancel_count={len(controls)} button_count={len(buttons)} "
        f"cancel_name={first_cancel.get('name')!r} "
        f"cancel_desc={first_cancel.get('description')!r}"
    )


if __name__ == "__main__":
    raise SystemExit(main())
