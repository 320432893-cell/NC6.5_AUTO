import json

import keyboard

from core.logger import log


def load_config(path="config.json"):
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["_config_path"] = path
    return cfg


def check_abort():
    if keyboard.is_pressed("space"):
        log.warning("用户按下空格，脚本紧急停止")
        raise SystemExit("用户紧急停止")
    if keyboard.is_pressed("esc"):
        log.warning("用户按下ESC，脚本中断")
        raise SystemExit("用户中断")
