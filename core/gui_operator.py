import time
import pyautogui
import pyperclip
from pathlib import Path
from PIL import Image, ImageChops, ImageStat
from core.jab_operator import JABOperator
from core.logger import log, ScreenRecorder
from core.utils import check_abort, format_amount, emergency_recovery


class GUIOperator:
    def __init__(self, config):
        self.cfg = config
        self.pos = config["positions"]
        self.timing = config["timing"]
        self.retry_cfg = config["retry"]
        self.images_dir = Path("images")
        self.verify_amount = config.get("verify_amount", True)

        self.recorder = ScreenRecorder()
        self.recorder.enabled = config.get("debug_screenshots", False)
        self.jab = JABOperator(config) if config.get("jab", {}).get("enabled", False) else None

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = self.timing.get("global_pause", 0.1)

    def init_ocr(self):
        """初始化检查，当前不使用 OCR"""
        log.info("GUI 操作器初始化完成（图像识别为可选项）")

    def _get_region_from_position(self, key):
        area = self.pos[key]
        if isinstance(area[0], list):
            (x1, y1), (x2, y2) = area
            return (min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))

        x, y = area
        return (max(0, x - 75), max(0, y - 20), 150, 40)

    def _capture_region(self, key):
        region = self._get_region_from_position(key)
        screenshot = pyautogui.screenshot(region=region)
        return screenshot.convert("L"), region

    def _calculate_region_difference_ratio(self, baseline_image, current_image):
        diff_image = ImageChops.difference(baseline_image, current_image)
        stat = ImageStat.Stat(diff_image)
        mean_diff = stat.mean[0] if stat.mean else 0
        return mean_diff / 255

    def wait_for_image(self, image_name, timeout=None, region=None, desc=""):
        timeout = timeout or self.timing["page_timeout"]
        image_path = self.images_dir / image_name

        if self.timing.get("skip_image_verify", False):
            log.debug(f"跳过图像验证: {desc or image_name}")
            time.sleep(self.timing.get("missing_image_wait", 0.2))
            return True

        if not image_path.exists():
            log.warning(f"截图不存在: {image_name}，使用固定等待")
            time.sleep(self.timing.get("missing_image_wait", 2))
            return True

        log.debug(f"等待图像: {desc or image_name}")
        start = time.time()

        while time.time() - start < timeout:
            check_abort()
            try:
                loc = pyautogui.locateOnScreen(
                    str(image_path),
                    confidence=0.8,
                    region=region
                )
                if loc:
                    log.debug(f"检测到: {desc or image_name}")
                    return True
            except Exception as e:
                log.debug(f"图像识别异常: {e}")

            time.sleep(self.timing.get("image_poll_interval", 0.3))

        log.warning(f"等待超时: {desc or image_name}")
        return False

    def click_pos_with_verify(self, key, desc="", verify_image=None):
        """点击并验证是否生效"""
        check_abort()
        x, y = self.pos[key]

        max_retries = self.retry_cfg.get("click_retries", 2)

        for attempt in range(max_retries):
            pyautogui.click(x, y)
            if desc:
                log.debug(f"点击 {desc} ({x}, {y})")
            time.sleep(self.timing["action_delay"])

            # 如果提供了验证图像，检查是否出现
            if verify_image:
                if self.wait_for_image(verify_image, timeout=2, desc=f"验证{desc}"):
                    return True
                else:
                    log.warning(f"{desc} 点击可能未生效，重试 {attempt + 1}/{max_retries}")
            else:
                return True

        raise Exception(f"{desc} 点击失败")

    def click_pos(self, key, desc=""):
        """点击坐标，带鼠标移动轨迹"""
        check_abort()
        x, y = self.pos[key]

        # ===== 新增：先移动到目标位置 =====
        if self.cfg.get("use_mouse_movement", True):
            duration = self.timing.get("mouse_move_duration", 0.1)
            pyautogui.moveTo(x, y, duration=duration)
            time.sleep(0.05)
        # ===== 移动结束 =====

        # 再点击
        pyautogui.click(x, y)
        if desc:
            log.debug(f"点击 {desc} ({x}, {y})")
        time.sleep(self.timing["action_delay"])
    def click_amount_cell(self):
        self.click_pos("first_amount_cell", "来源金额列单元格")

    def find_amount(self, amount_str):
        log.debug(f"查找金额: {amount_str}")

        pyautogui.hotkey("ctrl", "f")
        check_abort()

        if not self.wait_for_image("find_window_title.png", desc="查找窗口"):
            raise Exception("查找窗口未打开")

        # 使用剪贴板粘贴，避免单元格空格干扰
        pyperclip.copy(amount_str)
        time.sleep(0.1)
        check_abort()

        # 焦点已在输入框，直接全选并粘贴
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.05)
        check_abort()
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.2)
        check_abort()

        # 查找下一个
        self.click_pos("find_next_btn", "查找下一个")
        time.sleep(0.5)
        check_abort()

        # 检测是否找到
        if not self.verify_found_amount(amount_str):
            self.click_pos("find_close_btn", "关闭查找")
            raise Exception(f"金额 {amount_str} 在财务系统中没有找到")

        # 关闭查找窗口
        self.click_pos("find_close_btn", "关闭查找")
        time.sleep(self.timing["action_delay"])
    def verify_found_amount(self, expected_amount):
        """
        验证金额是否找到：
        - 如果有 not_found.png 截图，检测是否出现"没有找到！"
        - 如果没有截图，直接认为找到了（降级为不验证）
        """
        if not self.verify_amount:
            log.debug(f"跳过金额查找验证: {expected_amount}")
            return True

        not_found_path = self.images_dir / "not_found.png"

        if not not_found_path.exists():
            # 没有截图，跳过验证，直接认为找到
            log.debug(f"✓ 金额 {expected_amount} 查找成功（未配置验证截图）")
            return True

        # 有截图，快速检测是否出现"没有找到"提示
        found = self.wait_for_image("not_found.png", timeout=0.6, desc="没有找到提示")
        if found:
            log.warning(f"金额 {expected_amount} 没有找到！")
            return False

        log.debug(f"✓ 金额 {expected_amount} 查找成功")
        return True
    def do_generate(self):
        log.debug("点击生成 -> 前台生成")
        if self.jab:
            try:
                if self.jab.do_generate_front():
                    return
                log.warning("JAB 触发生成失败，降级使用坐标点击")
            except Exception as e:
                log.warning(f"JAB 触发生成异常，降级使用坐标点击: {e}")

        self.click_pos("generate_btn", "生成按钮")
        time.sleep(0.5)
        check_abort()
        self.click_pos("front_generate", "前台生成")

    def do_save(self):
        log.debug("等待凭证界面...")

        if not self.wait_for_image("save_btn.png", desc="凭证界面"):
            raise Exception("凭证界面未出现")

        # 等待界面完全加载
        voucher_load_wait = self.timing.get("voucher_load_wait", 0.5)
        time.sleep(voucher_load_wait)
        check_abort()

        log.debug("按 Ctrl+S 保存")
        pyautogui.hotkey("ctrl", "s")
        log.debug("已按 Ctrl+S，默认继续后续流程")
        time.sleep(0.2)
        check_abort()
        return True

    def validate_voucher(self, voucher):
        if not voucher:
            return False, "凭证号为空"
        if not voucher.isdigit():
            return False, f"凭证号包含非数字字符: {voucher}"
        if len(voucher) > 6:
            return False, f"凭证号长度异常: {len(voucher)}位"
        if voucher == "0":
            return False, "凭证号为0"
        return True, ""

    def copy_voucher_num(self):
        for attempt in range(self.retry_cfg["copy_retries"]):
            try:
                pyperclip.copy("")
                time.sleep(0.1)
                check_abort()

                self.click_pos("voucher_num_box", "凭证号框")
                time.sleep(0.2)
                check_abort()

                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.1)
                check_abort()
                pyautogui.hotkey("ctrl", "c")
                time.sleep(0.3)
                check_abort()

                voucher = pyperclip.paste().strip()

                is_valid, error_msg = self.validate_voucher(voucher)
                if is_valid:
                    log.debug(f"复制凭证号: {voucher}")
                    return voucher
                else:
                    log.warning(f"{error_msg}, 重试 {attempt + 1}/{self.retry_cfg['copy_retries']}")
                    time.sleep(0.5)

            except Exception as e:
                log.error(f"复制凭证号异常: {e}")
                time.sleep(0.5)

        log.error("多次尝试后仍无法获取有效凭证号")
        return None


    def do_return(self):
        log.debug("点击返回")
        self.click_pos("return_btn", "返回按钮")

        if not self.wait_for_image("generate_btn.png", desc="主界面"):
            log.warning("返回主界面超时，继续执行")

        time.sleep(self.timing["action_delay"])

    def process_one(self, item, data_handler):
        """处理单条记录，带详细日志"""
        amount_str = format_amount(item["amount"])
        row = item["row"]

        log.info(f"{'=' * 50}")
        log.info(f"开始处理: 行{row}, 金额{amount_str}")
        log.info(f"{'=' * 50}")
        for attempt in range(self.retry_cfg["max_retries"]):
            try:
                log.debug(f"[步骤1/6] 点击来源金额列")
                self.click_amount_cell()
                self.recorder.capture(f"row{row}_00_start")
                log.debug(f"[步骤2/6] 查找金额 {amount_str}")
                self.find_amount(amount_str)
                self.recorder.capture(f"row{row}_01_found")
                log.debug(f"[步骤3/6] 点击生成")
                self.do_generate()
                self.recorder.capture(f"row{row}_02_generated")
                log.debug(f"[步骤4/6] 保存凭证")
                if not self.do_save():
                    raise Exception("保存失败")
                self.recorder.capture(f"row{row}_03_saved")
                log.debug(f"[步骤5/6] 复制凭证号")
                voucher = self.copy_voucher_num()
                if not voucher:
                    raise Exception("凭证号为空或格式错误")
                self.recorder.capture(f"row{row}_04_copied")
                log.info(f"✅ 凭证号: {voucher}")

                log.debug(f"[步骤6/6] 校验凭证号并写入Excel")
                # 获取上一行凭证号进行对比
                previous_voucher = data_handler.get_previous_voucher(row)
                if previous_voucher and previous_voucher == voucher:
                    raise Exception(
                        f"当前凭证号 {voucher} 与上一行凭证号相同，上一条可能保存失败，脚本已停止"
                    )

                data_handler.save_voucher(row, voucher)
                self.do_return()
                time.sleep(self.timing["between_tasks"])

                log.info(f"行{row} 处理成功")
                return True
            except Exception as e:
                log.error(f"第 {attempt + 1} 次尝试失败: {e}")
                emergency_recovery()
                if not self.wait_for_image("generate_btn.png", timeout=10, desc="恢复主界面"):
                    log.error("无法恢复到主界面")
                    input("请手动返回主界面后按回车继续，或Ctrl+C退出...")
                if attempt == self.retry_cfg["max_retries"] - 1:
                    data_handler.save_voucher(row, str(e), "fail")
                    return False
                time.sleep(1)
        return False
