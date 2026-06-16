import os
import threading
import time

from core import jab_window
from core.jab_control_mixin import JABControlMixin
from core.jab_near_label_mixin import JABNearLabelMixin
from core.jab_path_mixin import JABPathMixin
from core.logger import log
from core.jab_table_mixin import JABTableMixin
from core.utils import check_abort
from tools.jab_probe import (
    configure_jab,
    load_access_bridge,
    run_windows_access_bridge,
)


def take_desktop_screenshot():
    import pyautogui

    return pyautogui.screenshot()


class JABOperator(JABControlMixin, JABNearLabelMixin, JABPathMixin, JABTableMixin):
    """Small Java Access Bridge wrapper for stable NC button/menu actions."""

    def __init__(self, config):
        self.config = config
        jab_cfg = config.get("jab", {})
        self.dll_path = jab_cfg.get("dll_path") or None
        self.startup_wait = jab_cfg.get("startup_wait", 2.0)
        self.search_timeout = jab_cfg.get("search_timeout", 5.0)
        self.max_depth = jab_cfg.get("max_depth", 50)
        self.max_children = jab_cfg.get("max_children", 1000)
        self.menu_wait = jab_cfg.get("menu_wait", 0.5)
        self.amount_col = jab_cfg.get("amount_col", 4)
        self.partner_col = jab_cfg.get("partner_col", 3)
        self.selection_col = jab_cfg.get("selection_col", 0)
        self.save_button_path = jab_cfg.get("save_button_path")
        self.save_button_title = jab_cfg.get("save_button_title", "制单")
        self.save_button_class = jab_cfg.get("save_button_class", "SunAwtDialog")
        self.hide_blank_awt_windows_enabled = jab_cfg.get(
            "hide_blank_awt_windows", True
        )

        self.dll = None
        self.loaded_path = None
        self.stop_pump = None
        self.pump_thread = None
        self.table_cache = {}

    def ensure_started(self):
        if self.dll:
            return

        if os.name != "nt":
            raise RuntimeError("Java Access Bridge must run under Windows Python.")

        self.dll, self.loaded_path = load_access_bridge(self.dll_path)
        configure_jab(self.dll)

        if hasattr(self.dll, "initializeAccessBridge"):
            if not self.dll.initializeAccessBridge():
                raise RuntimeError("initializeAccessBridge returned false.")
        else:
            self.stop_pump = threading.Event()
            self.pump_thread = threading.Thread(
                target=run_windows_access_bridge,
                args=(self.dll, self.stop_pump),
                daemon=True,
            )
            self.pump_thread.start()

        time.sleep(self.startup_wait)
        self.hide_blank_awt_windows()
        log.info(f"JAB 已加载: {self.loaded_path}")

    def close(self):
        self.hide_blank_awt_windows()
        self.clear_table_cache()
        if self.stop_pump:
            self.stop_pump.set()
        if self.pump_thread:
            self.pump_thread.join(timeout=1)

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass

    def hide_blank_awt_windows(self):
        """Force-hide only hidden no-title AWT residue left by JAB/Java.

        No-title small SunAwtWindow instances are also used by real NC popup
        menus, reference dialogs, and dropdowns. Visible small windows must
        never be disabled or moved here.
        """
        return jab_window.hide_blank_awt_windows(self.hide_blank_awt_windows_enabled)

    def close_window_by_title(self, title, class_name=None, wait=None):
        return jab_window.close_window_by_title(
            title,
            class_name=class_name,
            wait=wait,
            menu_wait=self.menu_wait,
            clear_table_cache=self.clear_table_cache,
        )

    def activate_window_by_title(self, title, class_name=None, timeout=None):
        return jab_window.activate_window_by_title(
            title,
            class_name=class_name,
            timeout=timeout,
            search_timeout=self.search_timeout,
        )

    def get_foreground_window_info(self):
        return jab_window.get_foreground_window_info()

    def foreground_window_matches(self, title, class_name=None):
        return jab_window.foreground_window_matches(title, class_name=class_name)

    def wait_window_by_title(
        self,
        title,
        class_name=None,
        timeout=None,
        include_children=False,
        visible_only=True,
        interval=0.2,
    ):
        return jab_window.wait_window_by_title(
            title,
            class_name=class_name,
            timeout=timeout,
            include_children=include_children,
            visible_only=visible_only,
            search_timeout=self.search_timeout,
            interval=interval,
        )

    def window_exists(self, title, class_name=None, include_children=False):
        return jab_window.window_exists(
            title,
            class_name=class_name,
            include_children=include_children,
        )

    def find_window_handle(
        self, title, class_name=None, visible_only=True, include_children=False
    ):
        return jab_window.find_window_handle(
            title,
            class_name=class_name,
            visible_only=visible_only,
            include_children=include_children,
        )

    def press_key(self, key, wait=None):
        import pyautogui

        pyautogui.press(key)
        time.sleep(self.menu_wait if wait is None else wait)

    def press_hotkey(self, *keys, wait=None):
        import pyautogui

        pyautogui.hotkey(*keys)
        time.sleep(self.menu_wait if wait is None else wait)

    def type_text(self, text, interval=0.01, wait=None):
        import pyautogui

        pyautogui.write(str(text), interval=interval)
        time.sleep(0 if wait is None else wait)

    def clipboard_copy(self, text):
        import pyperclip

        pyperclip.copy(str(text))

    def clipboard_paste(self, wait=None):
        self.press_hotkey("ctrl", "v", wait=wait)

    def clipboard_read(self):
        import pyperclip

        return pyperclip.paste()

    def take_screenshot(self):
        return take_desktop_screenshot()

    def do_generate_front(self):
        self.ensure_started()

        log.debug("JAB 点击生成按钮")
        if not self.click_control(
            "生成", roles=("push button",), timeout=self.search_timeout
        ):
            return False

        time.sleep(self.menu_wait)
        check_abort()

        log.debug("JAB 点击前台生成菜单项")
        return self.click_control(
            "前台生成", roles=("menu item",), timeout=self.search_timeout
        )

    def run_named_steps(self, steps):
        self.ensure_started()
        for index, step in enumerate(steps, start=1):
            check_abort()
            name = step["name"] if isinstance(step, dict) else str(step)
            role = step.get("role") if isinstance(step, dict) else None
            wait = (
                float(step.get("wait", self.menu_wait))
                if isinstance(step, dict)
                else self.menu_wait
            )
            timeout = (
                float(step.get("timeout", self.search_timeout))
                if isinstance(step, dict)
                else self.search_timeout
            )
            action = step.get("action") if isinstance(step, dict) else None
            require_showing = (
                step.get("require_showing", True) if isinstance(step, dict) else True
            )
            roles = (role,) if role else ()

            log.info(f"JAB 执行步骤 {index}/{len(steps)}: name={name} role={role}")
            if not self.click_control(
                name,
                roles=roles,
                timeout=timeout,
                action_name=action,
                require_showing=require_showing,
            ):
                raise RuntimeError(f"JAB 步骤失败: {name}")
            time.sleep(wait)

        return True

    def run_named_steps_in_window(
        self,
        steps,
        window_title=None,
        window_class=None,
        visible_only=True,
        scope_hwnd=None,
    ):
        self.ensure_started()
        for index, step in enumerate(steps, start=1):
            check_abort()
            name = step["name"] if isinstance(step, dict) else str(step)
            role = step.get("role") if isinstance(step, dict) else None
            wait = (
                float(step.get("wait", self.menu_wait))
                if isinstance(step, dict)
                else self.menu_wait
            )
            timeout = (
                float(step.get("timeout", self.search_timeout))
                if isinstance(step, dict)
                else self.search_timeout
            )
            action = step.get("action") if isinstance(step, dict) else None
            path = step.get("path") if isinstance(step, dict) else None
            require_showing = (
                step.get("require_showing", True) if isinstance(step, dict) else True
            )
            roles = (role,) if role else ()

            log.info(
                "JAB 执行窗口步骤 "
                f"{index}/{len(steps)}: window={window_title!r}/{window_class!r} "
                f"name={name} role={role}"
            )
            if path:
                thread = self.trigger_action_by_path_async(
                    path,
                    title=window_title,
                    class_name=window_class,
                    name=name,
                    role=role,
                    action_name=action,
                    timeout=float(step.get("return_timeout", 0.2))
                    if isinstance(step, dict)
                    else 0.2,
                    require_showing=require_showing,
                )
                if thread:
                    time.sleep(wait)
                    continue
            if not self.click_control(
                name,
                roles=roles,
                timeout=timeout,
                action_name=action,
                require_showing=require_showing,
                window_title=window_title,
                window_class=window_class,
                visible_only=visible_only,
                scope_hwnd=scope_hwnd,
            ):
                raise RuntimeError(
                    f"JAB 窗口步骤失败: window={window_title!r} name={name}"
                )
            time.sleep(wait)

        return True
