# 职责：打开或复用收款单查询条件窗口。
# 不做什么：不填写查询条件、不读取结果表、不做分页匹配。
# 允许依赖层：core JAB 操作对象、收款单查询配置。
# 谁不应该 import：core 层模块不应 import。

import time


def ensure_query_window(jab, config, query_cfg, jab_cfg, skip_open=False):
    title = jab_cfg["dialog_title"]
    class_name = jab_cfg["dialog_class"]
    timeout = float(query_cfg.get("open_timeout", query_cfg.get("timeout", 5)))
    existing_timeout = float(query_cfg.get("existing_dialog_timeout", 0.1))
    existing = jab.wait_window_by_title(
        title,
        class_name=class_name,
        timeout=existing_timeout,
        include_children=bool(query_cfg.get("dialog_include_children", True)),
        visible_only=bool(query_cfg.get("dialog_visible_only", True)),
        interval=float(query_cfg.get("window_poll_interval", 0.05)),
    )
    if existing or skip_open:
        return bool(existing)

    batch_open_query = (config.get("jab_batch") or {}).get("open_query") or {}
    main_title = query_cfg.get("main_title", batch_open_query.get("main_title", ""))
    main_class = query_cfg.get("main_class", batch_open_query.get("main_class"))
    if main_title:
        maximize = bool(query_cfg.get("maximize_main_window", True))
        activate = getattr(jab, "maximize_window_by_title", None) if maximize else None
        activate = activate or jab.activate_window_by_title
        activate(
            main_title,
            class_name=main_class,
            timeout=float(query_cfg.get("activate_timeout", 5)),
        )
    open_key = query_cfg.get("open_key", batch_open_query.get("key", "f3"))
    deadline = time.perf_counter() + timeout
    interval = float(query_cfg.get("window_poll_interval", 0.05))
    press_interval = float(query_cfg.get("open_key_retry_interval", 0.2))
    next_press_at = 0.0
    while time.perf_counter() < deadline:
        now = time.perf_counter()
        if now >= next_press_at:
            jab.press_key(open_key, wait=0.0)
            next_press_at = now + press_interval
        opened = jab.wait_window_by_title(
            title,
            class_name=class_name,
            timeout=interval,
            include_children=bool(query_cfg.get("dialog_include_children", True)),
            visible_only=bool(query_cfg.get("dialog_visible_only", True)),
            interval=interval,
        )
        if opened:
            return True
    return False
