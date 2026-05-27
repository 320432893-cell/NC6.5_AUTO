import time
import math
import pyautogui
from core.logger import log


class TestHelper:
    def __init__(self, config):
        self.cfg = config
        self.positions = config["positions"]

    def test_all_positions(self):
        log.info("开始测试所有坐标点击")
        log.info("每个坐标会显示圆圈标记3秒，请确认位置是否正确")
        log.info("如果位置不对，按ESC中断，重新运行坐标采集工具\n")

        time.sleep(3)

        for key, (x, y) in self.positions.items():
            log.info(f"测试: {key} -> ({x}, {y})")

            pyautogui.moveTo(x, y, duration=0.5)

            for i in range(8):
                angle = i * 45
                dx = int(20 * math.cos(math.radians(angle)))
                dy = int(20 * math.sin(math.radians(angle)))
                pyautogui.moveTo(x + dx, y + dy, duration=0.1)

            pyautogui.moveTo(x, y, duration=0.2)
            time.sleep(2)

        log.info("\n所有坐标测试完成")
        input("如果所有位置都正确，按回车继续。否则按Ctrl+C退出重新采集坐标。")

