import logging
import pyautogui
from pathlib import Path
from datetime import datetime


def setup_logger():
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    filename = log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"

    logger = logging.getLogger("nc_auto")
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler(filename, encoding="utf-8")
    fh.setLevel(logging.DEBUG)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


log = setup_logger()


class ScreenRecorder:
    def __init__(self):
        self.screenshots_dir = Path("logs/screenshots")
        self.screenshots_dir.mkdir(parents=True, exist_ok=True)
        self.enabled = False

    def capture(self, step_name):
        if not self.enabled:
            return

        timestamp = datetime.now().strftime("%H%M%S")
        filename = self.screenshots_dir / f"{timestamp}_{step_name}.png"

        try:
            screenshot = pyautogui.screenshot()
            screenshot.save(filename)
            log.debug(f"截图保存: {filename.name}")
        except Exception as e:
            log.warning(f"截图失败: {e}")