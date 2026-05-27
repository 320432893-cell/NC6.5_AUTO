import pyautogui
import keyboard
import json
from pathlib import Path

POSITIONS_TO_COLLECT = {
    "first_amount_cell": "来源金额列第一行单元格（任意数值单元格）",
    "find_next_btn": "查找窗口 - 查找下一个按钮",
    "find_close_btn": "查找窗口 - 关闭按钮",
    "generate_btn": "主界面 - 生成按钮",
    "front_generate": "生成菜单 - 前台生成选项",
    "voucher_num_box": "凭证界面 - 凭证号输入框（77那个框的中心）",
    "return_btn": "凭证界面 - 返回按钮"
}


def collect_positions():
    print("=" * 60)
    print("NC6.5 坐标采集工具")
    print("=" * 60)
    print("\n操作说明：")
    print("1. 将鼠标移动到提示的目标位置")
    print("2. 按空格键记录坐标")
    print("3. 按ESC可以随时退出")
    print("\n采集顺序：")
    print("第一轮：主界面（2个坐标）")
    print("  - 来源金额列单元格")
    print("  - 生成按钮")
    print("\n第二轮：查找窗口（2个坐标）")
    print("  - 手动按 Ctrl+F 打开查找窗口")
    print("  - 查找下一个按钮")
    print("  - 关闭按钮")
    print("\n第三轮：生成菜单（1个坐标）")
    print("  - 关闭查找窗口，点击生成按钮")
    print("  - 前台生成选项")
    print("\n第四轮：凭证界面（2个坐标）")
    print("  - 点击前台生成进入凭证界面")
    print("  - 凭证号输入框")
    print("  - 返回按钮")
    print("\n注意：")
    print("- 采集前请确保NC窗口已最大化")
    print("- 所有坐标都是用于点击，不需要截图")
    print("=" * 60)

    input("\n准备好后按回车开始...")

    positions = {}

    for key, desc in POSITIONS_TO_COLLECT.items():
        print(f"\n请将鼠标移到【{desc}】")
        print("按空格键记录坐标...")

        keyboard.wait("space")
        x, y = pyautogui.position()
        positions[key] = [x, y]

        print(f"✓ 已记录: ({x}, {y})")

        if keyboard.is_pressed("esc"):
            print("\n用户取消")
            return

    config_path = Path("config.json")
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        print("\n警告: config.json 不存在，将创建新文件")
        cfg = {}

    cfg["positions"] = positions

    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=4)

    print("\n" + "=" * 60)
    print("✓ 所有坐标已保存到 config.json")
    print("=" * 60)
    print("\n采集的坐标：")
    for key, pos in positions.items():
        print(f"  {key}: {pos}")


if __name__ == "__main__":
    try:
        collect_positions()
    except KeyboardInterrupt:
        print("\n\n用户中断")
    except Exception as e:
        print(f"\n错误: {e}")