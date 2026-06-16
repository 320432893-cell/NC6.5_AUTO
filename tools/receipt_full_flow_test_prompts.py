# 职责：给现场全流程测试 wrapper 提供双击可用的交互参数采集
# 不做什么：不调用 NC，不决定保存/查询策略

from __future__ import annotations

import os
from pathlib import Path


DEFAULT_EXCEL_NAME = "📊Pending Order & Credit Management_💸Payments来款通知.xlsx"


def build_interactive_args(
    mode_label: str,
    *,
    default_rows: str = "",
    default_limit: str = "1",
    default_start_delay: str = "0",
) -> list[str]:
    print()
    print(f"收款单全流程现场测试：{mode_label}")
    print("直接回车使用默认值；输入 q 后回车取消。")
    print()
    rows_prompt = "Sheet1 行号，多个用逗号分隔"
    if default_rows:
        rows_prompt += f"，默认 {default_rows}: "
    else:
        rows_prompt += "，默认自动取第一条通过预检行: "
    rows = ask(rows_prompt, default=default_rows)
    limit = ask(f"最多处理几行，默认 {default_limit}: ", default=default_limit)
    start_delay = ask(
        f"启动前等待秒数，默认 {default_start_delay}: ",
        default=default_start_delay,
    )

    args: list[str] = ["--limit", limit, "--start-delay", start_delay]
    if rows:
        if "," in rows:
            args += ["--excel-rows", rows]
        else:
            args += ["--excel-row", rows]
    return args


def with_default_excel_path(argv: list[str], root: Path) -> list[str]:
    if "--excel-path" in argv:
        return argv
    excel_path = find_default_excel_path(root)
    if excel_path:
        return [*argv, "--excel-path", str(excel_path)]
    print()
    print(f"未找到默认 Excel 文件：{DEFAULT_EXCEL_NAME}")
    value = ask("请输入 Excel 完整路径，或输入 q 取消: ")
    return [*argv, "--excel-path", value]


def find_default_excel_path(root: Path) -> Path | None:
    candidates = [
        root / DEFAULT_EXCEL_NAME,
        Path.home() / "Downloads" / DEFAULT_EXCEL_NAME,
        windows_home_downloads() / DEFAULT_EXCEL_NAME,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def windows_home_downloads() -> Path:
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        return Path(userprofile) / "Downloads"
    home = os.environ.get("HOME", "")
    if home.startswith("/home/"):
        username = Path(home).name
        return Path("/mnt/c/Users") / username / "Downloads"
    return Path("C:/Users/Queclink/Downloads")


def ask(prompt: str, default: str = "") -> str:
    value = input(prompt).strip()
    if value.lower() in {"q", "quit", "exit"}:
        raise SystemExit("用户取消")
    return value or default
