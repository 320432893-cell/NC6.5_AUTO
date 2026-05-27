import sys
import time
from pathlib import Path
from core.logger import log
from core.utils import (
    load_config, check_abort, set_dpi_aware,
    check_screen_resolution, activate_nc_window, health_check,
    format_amount
)
from core.data_handler import DataHandler
from core.gui_operator import GUIOperator
from core.test_helper import TestHelper


def check_amount_contains(queue, finance_amounts):
    """检测金额子串包含关系：NC模糊搜索会导致搜到错误记录"""
    # 所有财务金额的格式化字符串
    all_strs = [format_amount(a) for a in finance_amounts]
    risky = []

    for item in queue:
        s = format_amount(item["amount"])
        # 检查是否有其他金额字符串包含当前金额（当前是别人的子串）
        containers = [a for a in all_strs if a != s and s in a]
        if containers:
            risky.append(item)
            log.warning(
                f"行{item['row']} 金额{s} 被包含于其他金额({', '.join(containers[:3])}{'...' if len(containers) > 3 else ''})，NC搜索可能误匹配"
            )

    return risky


def validate_config(cfg):
    required = ["excel_path", "positions", "timing", "retry"]
    for key in required:
        if key not in cfg:
            raise ValueError(f"配置文件缺少必需项: {key}")

    for key, pos in cfg["positions"].items():
        if pos == [0, 0]:
            raise ValueError(f"坐标 {key} 未采集，请先运行 collect_positions.py")

    excel_path = Path(cfg["excel_path"])
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel文件不存在: {excel_path}")

    log.info("配置验证通过")


def run_tests(cfg):
    tester = TestHelper(cfg)

    while True:
        print("\n" + "=" * 60)
        print("测试模式")
        print("=" * 60)
        print("\n可选测试项：")
        print("1. 测试所有坐标位置（推荐首次运行）")
        print("2. 跑两条后暂停测试")
        print("3. 跳过测试，直接运行")

        choice = input("\n请选择 (1-3): ").strip()

        if choice == "1":
            tester.test_all_positions()
            return None
        elif choice == "2":
            log.info("启用两条暂停测试")
            return 2
        elif choice == "3":
            log.info("跳过测试")
            return None

        log.warning("无效选择，请输入 1、2 或 3")


def main():
    print("\n" + "=" * 60)
    print("NC6.5 自动生成凭证脚本")
    print("=" * 60)

    try:
        set_dpi_aware()
        check_screen_resolution()

        cfg = load_config()
        log.info("配置加载成功")

        validate_config(cfg)

        data_handler = DataHandler(cfg)
        gui_operator = GUIOperator(cfg)

        log.info("开始加载数据...")
        my_data = data_handler.load_my_data()
        finance_amounts = data_handler.load_finance_data()

        queue, duplicates, not_found = data_handler.check_duplicates(
            my_data, finance_amounts
        )

        if duplicates:
            data_handler.mark_issues(duplicates, "重复")
        if not_found:
            data_handler.mark_issues(not_found, "未找到")

        # 检测金额子串包含关系
        contains_risk = check_amount_contains(queue, finance_amounts)
        if contains_risk:
            data_handler.mark_issues(contains_risk, "金额包含风险")
            risk_rows = {item["row"] for item in contains_risk}
            queue = [item for item in queue if item["row"] not in risk_rows]

        if not queue:
            log.warning("没有可操作的数据")
            return

        pause_after = run_tests(cfg)

        progress = data_handler.load_progress()
        start_index = 0

        if progress:
            log.info(f"检测到上次进度: {progress['current']}/{progress['total']}")
            resume = input("是否从上次中断处继续？(y/n): ").strip().lower()
            if resume == 'y':
                start_index = progress['current']
            else:
                data_handler.clear_progress()

        print("\n" + "=" * 60)
        print(f"即将处理 {len(queue) - start_index} 条记录")
        print(f"重复/未找到: {len(duplicates) + len(not_found)} 条需人工处理")
        if contains_risk:
            print(f"金额包含风险: {len(contains_risk)} 条已移出队列，需人工处理")
        print("=" * 60)
        print("\n提示：")
        print("- 请确保NC主界面已打开并最大化")
        print("- 运行过程中按空格可紧急停止，ESC也可中断")
        print("- 日志保存在 logs/ 目录")

        input("\n按回车开始...")

        log.info("初始化 GUI 操作器")
        gui_operator.init_ocr()

        activate_nc_window()

        print("\n3秒后开始，请切换到NC窗口...")
        for i in range(3, 0, -1):
            print(f"{i}...")
            time.sleep(1)

        log.info("开始批量处理")
        success = 0
        fail = 0
        health_check_interval = cfg.get("health_check_interval", 10)
        processed_this_run = 0

        for i, item in enumerate(queue[start_index:], start=start_index):
            check_abort()

            if health_check_interval > 0 and i > 0 and i % health_check_interval == 0:
                if not health_check(gui_operator):
                    log.error("健康检查失败，暂停")
                    input("请检查NC窗口状态，按回车继续...")

            log.info(f"--- 进度: {i+1}/{len(queue)} ---")

            if gui_operator.process_one(item, data_handler):
                success += 1
            else:
                fail += 1

            data_handler.save_progress(i + 1, len(queue))
            processed_this_run += 1

            if pause_after and processed_this_run >= pause_after:
                print("\n已完成两条暂停测试。")
                choice = input("按回车继续处理剩余数据，输入 q 退出: ").strip().lower()
                if choice == "q":
                    log.info("两条暂停测试结束，已保留当前进度")
                    return
                pause_after = None

        data_handler.clear_progress()

        print("\n" + "=" * 60)
        log.info(f"处理完成！成功: {success}, 失败: {fail}")
        log.info(f"重复/未找到: {len(duplicates) + len(not_found)} 条需人工处理")
        log.info(f"结果已保存到: {cfg['excel_path']}")
        print("=" * 60)

    except KeyboardInterrupt:
        log.warning("用户中断")
    except Exception as e:
        log.error(f"程序异常: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
