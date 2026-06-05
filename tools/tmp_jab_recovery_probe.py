# T0 temporary probe. Delete after NC/JAB startup recovery rules are confirmed.

import argparse
import ctypes
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402
from tools.jab_health_check import check_jab_ready  # noqa: E402
from tools.jab_probe import enum_windows  # noqa: E402

START_DELAY_SECONDS = 2
WM_NULL = 0x0000
WM_SETFOCUS = 0x0007
WM_USER = 0x0400
BM_CLICK = 0x00F5
SMTO_ABORTIFHUNG = 0x0002
SW_RESTORE = 9
WATCH_CLASSES = {"YonyouUWnd", "SunAwtFrame", "SunAwtCanvas", "SunAwtToolkit"}


def print_header(include_status_call):
    print("测试功能：NC Java Access Bridge 注册恢复探测")
    print()
    print("目标：")
    print("1. 先读取当前 JAB 健康状态")
    print("2. 依次执行低风险窗口唤醒动作")
    print("3. 每一步后重新检查 SunAwt 是否变成 isJava/getContext 可用")
    print()
    print("本脚本不会做：写入、点击 NC 业务按钮、键盘输入、保存、暂存、关闭窗口")
    print("会做的窗口动作：等待、前台化 NC Java 窗口、发送 WM_NULL/WM_SETFOCUS")
    if include_status_call:
        print("额外动作：点击隐藏 Access Bridge status 的 Call info 按钮")
    print(f"启动后等待：{START_DELAY_SECONDS} 秒，用来切到 NC 窗口")
    print("=" * 60)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=str(ROOT / "config.json"))
    parser.add_argument("--include-status-call", action="store_true")
    args = parser.parse_args()

    print_header(args.include_status_call)
    print()
    print(f"请在 {START_DELAY_SECONDS} 秒内切到 NC 窗口...")
    time.sleep(START_DELAY_SECONDS)
    print("开始恢复探测。")

    config = load_config(args.config)
    jab = JABOperator(config)
    jab.hide_blank_awt_windows_enabled = False
    try:
        jab.ensure_started()
        report_step(jab, "初始状态")
        if check_jab_ready(jab).get("ok"):
            print("结果：JAB 已可用，不需要恢复动作。")
            return 0

        run_action(jab, "等待消息泵 5 秒", lambda: time.sleep(5.0))
        if check_jab_ready(jab).get("ok"):
            return 0

        run_action(jab, "前台化可见 SunAwtFrame/Yonyou 窗口", activate_nc_windows)
        if check_jab_ready(jab).get("ok"):
            return 0

        run_action(jab, "向 Java 窗口发送 WM_NULL", send_wm_null_to_java_windows)
        if check_jab_ready(jab).get("ok"):
            return 0

        run_action(jab, "向可见 SunAwtCanvas 发送 WM_SETFOCUS", send_focus_to_canvas)
        if check_jab_ready(jab).get("ok"):
            return 0

        if args.include_status_call:
            run_action(
                jab, "点击 Access Bridge status / Call info", click_status_call_info
            )

        final = check_jab_ready(jab)
        print()
        if final.get("ok"):
            print("结果：JAB 已恢复，可继续运行收款单脚本。")
            return 0
        print("结果：本轮低风险动作未恢复 JAB 注册。")
        print(f"原因：{final.get('reason')}")
        print("下一步应验证：退出并重新启动 UClient/NC 后，健康检查是否变为 ok=True。")
        return 1
    finally:
        jab.close()


def run_action(jab, title, action):
    print()
    print("-" * 60)
    print(f"动作：{title}")
    action()
    time.sleep(1.0)
    report_step(jab, f"{title} 后")


def report_step(jab, title):
    health = check_jab_ready(jab)
    print(f"{title}: ok={health.get('ok')} reason={health.get('reason')}")
    for item in (health.get("visible_sunawt") or [])[:5]:
        print(
            "  "
            f"hwnd={item.get('hwnd')} class={item.get('class_name')} "
            f"title={item.get('title') or '<无标题>'} "
            f"isJava={item.get('is_java')} getContext={item.get('get_context_ok')}"
        )


def activate_nc_windows():
    user32 = ctypes.windll.user32
    for hwnd, title, class_name, _pid, visible in enum_windows(include_children=True):
        if not visible:
            continue
        if class_name not in ("SunAwtFrame", "YonyouUWnd"):
            continue
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)


def send_wm_null_to_java_windows():
    user32 = ctypes.windll.user32
    for hwnd, _title, class_name, _pid, _visible in enum_windows(include_children=True):
        if class_name not in WATCH_CLASSES and not class_name.startswith("SunAwt"):
            continue
        send_message_timeout(user32, hwnd, WM_NULL, 0, 0)


def send_focus_to_canvas():
    user32 = ctypes.windll.user32
    for hwnd, _title, class_name, _pid, visible in enum_windows(include_children=True):
        if visible and class_name == "SunAwtCanvas":
            send_message_timeout(user32, hwnd, WM_SETFOCUS, 0, 0)


def click_status_call_info():
    user32 = ctypes.windll.user32
    for hwnd, title, class_name, _pid, _visible in enum_windows(include_children=True):
        if class_name == "Button" and title == "Call info":
            send_message_timeout(user32, hwnd, BM_CLICK, 0, 0)


def send_message_timeout(user32, hwnd, message, wparam, lparam):
    result = ctypes.c_void_p()
    user32.SendMessageTimeoutW(
        hwnd,
        message,
        wparam,
        lparam,
        SMTO_ABORTIFHUNG,
        500,
        ctypes.byref(result),
    )
    return result.value


if __name__ == "__main__":
    raise SystemExit(main())
