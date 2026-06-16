import json

from core.logger import log
from core.paths import stop_flag_path

try:
    import keyboard
except ModuleNotFoundError:
    keyboard = None


def load_config(path="config.json"):
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["_config_path"] = path
    return cfg


def check_abort():
    # 外部停止信号:控制进程(如 GUI 外壳)落标志文件即请求停止。
    # 与键盘急停走同一批安全检查点,故继承相同的"停在可续跑位置"语义。
    if stop_flag_path().exists():
        log.warning("检测到外部停止标志，脚本中断")
        raise SystemExit("外部停止")
    if keyboard is None:
        return
    if keyboard.is_pressed("space"):
        log.warning("用户按下空格，脚本紧急停止")
        raise SystemExit("用户紧急停止")
    if keyboard.is_pressed("esc"):
        log.warning("用户按下ESC，脚本中断")
        raise SystemExit("用户中断")
