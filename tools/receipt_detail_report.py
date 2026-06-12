# 职责：打印收款单明细正式测试入口的人类可读报告
# 不做什么：不执行 JAB/GUI 动作，不读取 Excel，不决定明细写入规则
# 允许依赖层：tools.jab_health_check、tools.receipt_account_reference_try、tools.receipt_detail_rows
# 谁不应该 import：底层 JAB operator、Sheet 写入、收款匹配模块不应 import

from tools.jab_health_check import print_jab_health_failure
from tools.receipt_account_reference_try import STOP_HOTKEY
from tools.receipt_detail_rows import ADD_FEE_ROW_HOTKEY

KEYBOARD_INPUT_COMMIT_KEY = "Right"


def print_header(account_no, args, start_delay_seconds):
    if args.cleanup_extra_rows_only:
        print("测试功能：收款单明细多余行清理")
    elif args.fee_only:
        print("测试功能：收款单手续费行从增行开始完整试写")
    else:
        print("测试功能：收款单明细主行正式入口试写")
    print()
    print("测试数据来源：")
    print(f"1. 银行标签：{args.bank_label}（来自 config.json 映射）")
    print(f"2. 收款银行账户：{account_no}")
    if args.fee_only:
        print(f"3. 手续费行：手续费 / 660305 / {args.fee_amount} / 网银")
    else:
        print("3. 明细主行：货款 / 账号 / 1002 / 1090 / 网银")
    print()
    print("前置条件：")
    print("1. NC 已停在收款单自制录入界面")
    print("2. 当前明细主行已存在；手续费模式要求主行已写好")
    print("3. 当前没有打开参照窗口或提示框")
    print()
    print("本脚本会做：")
    if args.cleanup_extra_rows_only:
        print("1. 定位 25 列明细表，并读取当前行数")
        print("2. 删除第 1 行以外的多余行")
        print("3. 删除后再次读取明细表")
    elif args.fee_only:
        print("1. 定位 25 列明细表，并读取当前行数")
        print(f"2. 前台守卫通过后发送 {ADD_FEE_ROW_HOTKEY} 新增手续费行")
        print("3. 在新增后的最后一行写入：手续费、660305、金额、网银")
        print("4. 写完后统一读回明细表关键列；失败字段最多修复 3 次")
    else:
        print("1. 读取表头【收款银行账户】状态")
        print("2. 定位 25 列明细表")
        print("3. 用 JAB 选中目标单元格后键盘写入明细主行；明细不写币种")
        print("4. 写完后统一读回明细表关键列；失败字段最多修复 3 次")
    print()
    print("不会做：保存、暂存、关闭收款单")
    if not args.fee_only:
        print(f"增行规则：只允许手续费非零分支使用 {ADD_FEE_ROW_HOTKEY}；主行不增行")
    print("说明：每个字段都会先用 JAB selection API 选中 row/col，再 F2/Ctrl+A/输入。")
    print(
        f"说明：提交使用方向键 {KEYBOARD_INPUT_COMMIT_KEY}，不使用 Tab，避免误触发增行。"
    )
    print("兼容性：明细字段定位不再依赖列宽、窗口大小或横向滚动坐标。")
    print(f"紧急停止：按 {STOP_HOTKEY}")
    print(f"启动后等待：{start_delay_seconds} 秒，用来切到 NC 窗口")
    print("=" * 60)


def print_table_snapshot(title, snapshot):
    print(title)
    if not snapshot.get("ok"):
        print(f"  失败：{snapshot.get('reason')}")
        return
    print(f"  明细表：{snapshot.get('row_count')} 行 x {snapshot.get('col_count')} 列")
    rows = snapshot.get("rows") or []
    for row in rows[:3]:
        row_no = int(row.get("row_index", 0)) + 1
        cells = row.get("cells") or {}
        print(
            f"  第 {row_no} 行关键列："
            f"业务类型={cells.get('1')!r}, "
            f"币种={cells.get('3')!r}, "
            f"账户={cells.get('4')!r}, "
            f"科目={cells.get('5')!r}, "
            f"金额={cells.get('7')!r}, "
            f"结算={cells.get('11')!r}"
        )


def print_table_candidates(candidates):
    if not candidates:
        print("  未发现任何 JAB 表格候选。")
        return
    print("  候选表摘要：")
    for item in candidates[:5]:
        print(
            "  - "
            f"{item.get('row_count')} 行 x {item.get('col_count')} 列，"
            f"score={item.get('score')}，"
            f"窗口={((item.get('window') or {}).get('title') or '<无标题>')}"
        )
        print(f"    原因：{item.get('reasons')}")
        rows = item.get("rows") or []
        if rows:
            cells = rows[0].get("cells") or {}
            print(f"    第 1 行关键列：{cells}")


def print_fill_summary(steps):
    print("明细写入结果：")
    for step in steps:
        name = step.get("name")
        value = step.get("value")
        ok = "成功" if step.get("ok") else "失败"
        actual = step.get("actual")
        attempts = step.get("attempts") or []
        elapsed = sum(float(item.get("seconds") or 0) for item in attempts)
        print(
            f"  {name}: {ok} | 期望={value!r} | 实际={actual!r} | "
            f"尝试={len(attempts)} | 用时={elapsed:.3f}s"
        )
        if step.get("partial_success"):
            print("    注意：本次已有字段写入成功，失败后需要人工核对当前明细行。")
        if step.get("target"):
            print(f"    目标单元格：{step.get('target')}")
        if step.get("geometry"):
            geometry = step.get("geometry") or {}
            print(
                "    表格依据："
                f"bounds={geometry.get('table_bounds')} "
                f"行数={geometry.get('row_count')} 列数={geometry.get('col_count')}"
            )
        if not step.get("ok"):
            print(f"    原因：{step.get('reason')}")


def print_summary(report):
    print()
    print("测试结果：")
    if report.get("stopped_by_hotkey"):
        print(f"已停止：检测到紧急停止键 {STOP_HOTKEY}。")
        print(f"停止位置：{report.get('failed_step')}")
        return

    if report.get("exception"):
        print(f"脚本异常：{report.get('exception')}")
        print(f"原因：{report.get('reason')}")
        return

    if report.get("failed_step") == "jab-health-check":
        print("JAB 启动状态：")
        health = report.get("jab_health") or {}
        if isinstance(health, dict):
            print_jab_health_failure(health)
        print("失败：当前不能读取 NC JAB 控件树，未执行表头检查和明细写入。")
        return

    header = report.get("header_account") or {}
    print("表头账户检查：")
    if header.get("accepted"):
        print(f"  已读到账户字段：{header.get('text') or header.get('description')}")
    else:
        print("  未确认读到账户字段；本轮仍继续尝试明细写入。")

    failed_step = report.get("failed_step")
    if failed_step == "locate-body-table":
        print("明细表定位：")
        print("  失败：没有找到符合收款单明细特征的 25 列表。")
        print("  本轮已阻塞停止，没有尝试写入明细。")
        print_table_candidates(report.get("table_candidates") or [])
        print("失败：请确认当前停在【收款单自制录入界面】，且没有参照窗口/提示框遮挡。")
        return

    print_table_snapshot("写入前明细表：", report.get("before_table") or {})
    _print_fee_row_add(report)
    if report.get("fill_steps") is not None:
        print_fill_summary(report.get("fill_steps") or [])
    _print_fee_account_clear(report)
    _print_extra_row_delete(report)
    if report.get("after_table") is not None:
        print_table_snapshot("写入后明细表：", report.get("after_table") or {})

    if report.get("ok"):
        if report.get("mode") == "cleanup-only":
            print("成功：明细第 1 行以外的多余行已清理；本次没有保存、没有暂存。")
        elif report.get("mode") == "fee-only":
            print("成功：手续费行已增行并写入，通过表格读回校验。")
        else:
            print(
                "成功：明细主行字段已按 JAB 单元格选中+键盘输入写入，并通过表格读回校验。"
            )
    else:
        print("失败：至少一个明细字段没有写入成功；本次没有保存、没有暂存。")
    _print_timings(report)


def _print_fee_row_add(report):
    if report.get("fee_row_add") is None:
        return
    add_row = report.get("fee_row_add") or {}
    print("手续费增行：")
    print(
        f"  {add_row.get('hotkey')}: "
        f"{'成功' if add_row.get('ok') else '失败'} | "
        f"行数 {add_row.get('before_rows')} -> {add_row.get('after_rows')}"
    )
    pressed = add_row.get("pressed") or {}
    if pressed.get("mode"):
        print(f"  发送方式：{pressed.get('mode')}")
    if add_row.get("overwrite_guard"):
        guard = add_row.get("overwrite_guard") or {}
        print(
            "  覆盖守卫："
            f"{'通过' if guard.get('ok') else '失败'} | "
            f"空行={guard.get('empty')} 已是手续费={guard.get('already_fee')}"
        )
    if not add_row.get("ok"):
        print(f"  原因：{add_row.get('reason')}")


def _print_fee_account_clear(report):
    if report.get("fee_account_clear") is None:
        return
    clear = report.get("fee_account_clear") or {}
    print("手续费账户清空：")
    if clear.get("skipped"):
        print("  跳过：手续费行账户本来就是空。")
    else:
        print(
            f"  {'成功' if clear.get('ok') else '失败'} | "
            f"清空前={clear.get('before')!r} | 清空后={clear.get('after')!r}"
        )
        if not clear.get("ok"):
            print(f"  原因：{clear.get('reason')}")


def _print_extra_row_delete(report):
    if report.get("extra_row_delete") is None:
        return
    delete = report.get("extra_row_delete") or {}
    print("多余空行删除：")
    if delete.get("skipped"):
        print("  跳过：未发现需要删除的多余行。")
    else:
        print(
            f"  {'成功' if delete.get('ok') else '失败'} | "
            f"行数 {delete.get('before_rows')} -> {delete.get('after_rows')}"
        )
        if delete.get("partial_success"):
            print("  注意：删行已有部分成功，失败后保留当前行数，需要人工核对。")
        if not delete.get("ok"):
            print(f"  原因：{delete.get('reason')}")


def _print_timings(report):
    if report.get("total_seconds") is not None:
        print(f"总用时：{float(report.get('total_seconds') or 0):.3f}s")
    timings = report.get("timings") or []
    if timings:
        print("阶段计时：")
        for item in timings:
            print(f"  {item.get('name')}: {float(item.get('seconds') or 0):.3f}s")
