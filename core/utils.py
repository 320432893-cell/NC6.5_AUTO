import json
import ctypes
import pyautogui
import keyboard
from core.logger import log

def ensure_nc_active():
    """确保NC窗口在前台"""
    try:
        import pygetwindow as gw
        windows = gw.getWindowsWithTitle("NC6.5")
        if windows:
            nc_win = windows[0]
            if not nc_win.isActive:
                nc_win.activate()
                log.warning("NC窗口失去焦点，已重新激活")
                import time
                time.sleep(0.5)
                return True
        return False
    except:
        return False
def load_config(path="config.json"):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(cfg, path="config.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=4)


def check_abort():
    if keyboard.is_pressed("space"):
        log.warning("用户按下空格，脚本紧急停止")
        raise SystemExit("用户紧急停止")
    if keyboard.is_pressed("esc"):
        log.warning("用户按下ESC，脚本中断")
        raise SystemExit("用户中断")


def format_amount(val):
    """
    格式化金额用于NC查找：
    - 整数金额（如1600.00）格式化为 "1600"
    - 小数金额（如156.18）格式化为 "156.18"
    """
    amount = float(val)
    # 判断是否为整数
    if amount == int(amount):
        return str(int(amount))
    else:
        return f"{amount:.2f}"


def set_dpi_aware():
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        log.debug("DPI感知已设置")
    except Exception as e:
        log.warning(f"DPI设置失败: {e}")


def check_screen_resolution():
    width, height = pyautogui.size()
    log.info(f"当前屏幕分辨率: {width}x{height}")


def activate_nc_window():
    try:
        import pygetwindow as gw
        windows = gw.getWindowsWithTitle("NC6.5")
        if not windows:
            log.warning("未找到NC窗口，请手动切换")
            return False
        nc_win = windows[0]
        nc_win.activate()
        nc_win.maximize()
        log.info("NC窗口已激活并最大化")
        return True
    except ImportError:
        log.warning("pygetwindow未安装，请手动切换到NC窗口")
        return False
    except Exception as e:
        log.warning(f"激活NC窗口失败: {e}")
        return False


def emergency_recovery():
    log.warning("执行紧急恢复...")
    for _ in range(5):
        pyautogui.press("esc")
        import time
        time.sleep(0.3)
    log.info("紧急恢复完成")


def health_check(gui_operator):
    log.info("执行健康检查...")

    if not gui_operator.wait_for_image("generate_btn.png", timeout=5, desc="主界面"):
        log.error("健康检查失败：主界面不可见")
        return False

    try:
        current_pos = pyautogui.position()
        pyautogui.moveTo(current_pos[0] + 10, current_pos[1] + 10, duration=0.2)
        pyautogui.moveTo(current_pos[0], current_pos[1], duration=0.2)
    except Exception as e:
        log.error(f"健康检查失败：鼠标控制异常 {e}")
        return False

    log.info("✓ 健康检查通过")
    return True
