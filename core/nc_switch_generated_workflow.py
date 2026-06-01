import subprocess
import sys
import time
from pathlib import Path

from core.errors import JABActionError, JABControlNotFound, WorkflowStateError
from core.logger import log


class NCSwitchGeneratedWorkflow:
    def __init__(self, processor):
        self.processor = processor

    def __getattr__(self, name):
        return getattr(self.processor, name)

    def switch_to_generated_list(self):
        with self.perf.span("switch_generated_total"):
            self.perf.event(
                "switch_generated_date",
                generated_date_value=self.generated_date_value,
            )
            self.run_state.set_stage("switch_detect_state")
            state = self.get_nc_workflow_state()
            if state == "voucher_empty":
                close_cfg = self.batch_cfg.get("close_voucher_window", {})
                with self.perf.span("switch_close_empty_voucher_window"):
                    self.jab.close_window_by_title(
                        close_cfg.get("title", self.voucher_window_title),
                        class_name=close_cfg.get(
                            "class_name", self.voucher_window_class
                        ),
                        wait=float(close_cfg.get("wait", 0.5)),
                    )
                state = self.get_nc_workflow_state()

            if state != "parent_ready":
                raise WorkflowStateError(
                    f"当前不能切换正式单据，必须先回到制单父界面: state={state}"
                )

            open_query = self.batch_cfg.get("open_query", {})
            query_method = open_query.get("method")
            self.run_state.set_stage(
                "switch_find_query_window",
                query_method=query_method,
                nc_state=state,
            )
            query_hwnd = self.find_query_window(open_query, timeout=0.5)
            if query_method == "jab_action":
                self.run_state.set_stage("switch_open_query_jab_action")
                self.record_event(
                    "event_query_open_start",
                    method="jab_action",
                    existing=bool(query_hwnd),
                )
                with self.perf.span("switch_open_query_jab_action"):
                    if query_hwnd:
                        log.info("查询窗口已存在，跳过 JAB 查询入口触发")
                    else:
                        try:
                            self.open_query_with_jab_action(open_query)
                        except RuntimeError:
                            if open_query.get("fallback_method") != "hotkey":
                                raise
                            log.warning("JAB 查询入口失败，回退 F3 热键")
                        query_hwnd = self.find_query_window(
                            open_query,
                            timeout=float(open_query.get("timeout", 5)),
                        )
                    if not query_hwnd and open_query.get("fallback_method") == "hotkey":
                        log.warning("JAB 查询入口未打开查询窗口，回退 F3 热键")
                        self.open_query_with_hotkey(open_query)
                        query_hwnd = self.find_query_window(
                            open_query,
                            timeout=float(open_query.get("timeout", 5)),
                        )
                    if not query_hwnd:
                        raise JABControlNotFound("未检测到查询窗口")
                self.record_transition(
                    "query_opened",
                    from_state="pending",
                    to_state="query_open",
                    method="jab_action",
                )
            elif query_method == "hotkey":
                self.run_state.set_stage("switch_open_query_hotkey")
                self.record_event(
                    "event_query_open_start",
                    method="hotkey",
                    existing=bool(query_hwnd),
                )
                with self.perf.span("switch_open_query"):
                    if not query_hwnd:
                        self.open_query_with_hotkey(open_query)
                        query_hwnd = self.find_query_window(
                            open_query,
                            timeout=float(open_query.get("timeout", 5)),
                        )
                    if not query_hwnd:
                        raise JABControlNotFound("按快捷键后未检测到查询窗口")
                self.record_transition(
                    "query_opened",
                    from_state="pending",
                    to_state="query_open",
                    method="hotkey",
                )
            elif query_method:
                raise WorkflowStateError(f"不支持的 open_query.method: {query_method}")

            steps = self.batch_cfg.get("switch_generated_steps", [])
            if not steps:
                raise WorkflowStateError(
                    "未配置 switch_generated_steps，暂不能自动切换到已生成列表"
                )
            try:
                self.run_state.set_stage("switch_run_query_steps")
                self.require_page_state(
                    "query_open",
                    items=None,
                    command="switch-generated",
                )
                self.record_event(
                    "event_query_confirm_start",
                    steps=len(steps),
                    generated_date_value=self.generated_date_value,
                )
                with self.perf.span("switch_run_steps", steps=len(steps)):
                    self.run_switch_generated_steps(
                        open_query,
                        steps,
                        query_method,
                        query_hwnd=query_hwnd,
                    )
                self.record_transition(
                    "query_confirmed",
                    from_state="query_open",
                    to_state="loading",
                    generated_date_value=self.generated_date_value,
                )
            except RuntimeError:
                if (
                    query_method == "jab_action"
                    and open_query.get("fallback_method") == "hotkey"
                ):
                    log.warning("JAB 查询窗口步骤失败，回退 F3 重新打开查询窗口")
                    self.open_query_with_hotkey(open_query)
                    query_hwnd = self.find_query_window(
                        open_query,
                        timeout=float(open_query.get("timeout", 5)),
                    )
                    if not query_hwnd:
                        raise JABControlNotFound("F3 回退后未检测到查询窗口")
                    with self.perf.span("switch_run_steps_fallback", steps=len(steps)):
                        self.run_switch_generated_steps(
                            open_query,
                            steps,
                            "hotkey",
                            query_hwnd=query_hwnd,
                        )
                else:
                    raise
            with self.perf.span("switch_generated_snapshot"):
                self.run_state.set_stage("switch_verify_generated_snapshot")
                state = self.require_page_state(
                    "generated",
                    items=None,
                    command="switch-generated",
                )
            self.record_transition(
                "generated_list_loaded",
                from_state="loading",
                to_state="generated",
                rows=state.table.get("row_count", 0) if state.table else 0,
            )
            rows = state.table.get("row_count", 0) if state.table else 0
            with_voucher = state.table.get("voucher_values", []) if state.table else []
            log.info(
                "已切换到已生成列表: "
                f"rows={rows} sample_voucher_count={len(with_voucher)}"
            )
            self.run_state.update_counts(generated_snapshot_rows=rows)
            self.run_state.set_stage("switch_generated_done")
            return True

    def run_switch_generated_steps(
        self,
        open_query,
        steps,
        query_method,
        query_hwnd=None,
    ):
        dialog_title = open_query.get("dialog_title", "查询")
        dialog_class = open_query.get("dialog_class", "SunAwtDialog")

        try:
            for index, step in enumerate(steps, start=1):
                step_name = self.get_switch_step_name(step, index)
                if isinstance(step, dict) and step.get("runner") == "subprocess":
                    subprocess_step = {
                        **step,
                        "title": dialog_title,
                        "class_name": dialog_class,
                    }
                    with self.perf.span(
                        f"switch_step_{step_name}",
                        step_index=index,
                        path=step.get("path"),
                        control_name=step.get("name"),
                        role=step.get("role"),
                        runner="subprocess",
                    ):
                        self.run_jab_action_subprocess(subprocess_step)
                        time.sleep(float(step.get("wait", 0.2)))
                else:
                    with self.perf.span(
                        f"switch_step_{step_name}",
                        step_index=index,
                        path=step.get("path") if isinstance(step, dict) else None,
                        control_name=step.get("name")
                        if isinstance(step, dict)
                        else None,
                        role=step.get("role") if isinstance(step, dict) else None,
                    ):
                        self.run_query_window_step(
                            step,
                            dialog_title=dialog_title,
                            dialog_class=dialog_class,
                        )
            return True
        except RuntimeError as exc:
            if query_method == "jab_action":
                raise
            log.warning(
                "查询窗口内步骤失败，回退全局控件查找: "
                f"window={dialog_title!r}/{dialog_class!r}, error={exc}"
            )
            self.jab.run_named_steps(steps)
            return True

    def get_switch_step_name(self, step, index):
        if not isinstance(step, dict):
            return f"{index}"
        if step.get("perf_name"):
            return str(step["perf_name"])
        name = step.get("name")
        set_text = step.get("set_text")
        path = step.get("path")
        if name == "正式单据":
            return "formal_action"
        if name == "确定":
            return "confirm_action"
        if set_text == "{generated_date_value}" and path:
            if str(path).endswith(".1.0.0"):
                return "date_from"
            if str(path).endswith(".1.2.0"):
                return "date_to"
            return "date_text"
        return f"{index}"

    def run_query_window_step(self, step, dialog_title, dialog_class):
        path = step.get("path") if isinstance(step, dict) else None
        if not path:
            self.jab.run_named_steps_in_window(
                [step],
                window_title=dialog_title,
                window_class=dialog_class,
                visible_only=True,
            )
            return

        wait_info = self.jab.wait_context_by_path(
            path,
            title=dialog_title,
            class_name=dialog_class,
            name=step.get("name"),
            role=step.get("role"),
            require_showing=bool(step.get("require_showing", False)),
            timeout=float(step.get("guard_timeout", step.get("timeout", 3))),
        )
        if not wait_info:
            raise JABControlNotFound(
                f"查询窗口控件未就绪: path={path} name={step.get('name')}"
            )

        set_text = step.get("set_text")
        if set_text == "{generated_date_value}":
            set_text = self.generated_date_value
        if set_text is not None:
            ok = self.jab.set_text_by_path(
                path,
                set_text,
                title=dialog_title,
                class_name=dialog_class,
                name=step.get("name"),
                role=step.get("role"),
                wait=float(step.get("wait", 0.0)),
                timeout=float(step.get("timeout", 1)),
                require_showing=bool(step.get("require_showing", False)),
            )
        else:
            ok = self.jab.do_action_by_path(
                path,
                title=dialog_title,
                class_name=dialog_class,
                name=step.get("name"),
                role=step.get("role"),
                action_name=step.get("action"),
                click_mode=step.get("click_mode"),
                wait=float(step.get("wait", 0.0)),
                timeout=float(step.get("timeout", 1)),
                require_showing=bool(step.get("require_showing", False)),
            )
        if not ok:
            raise JABActionError(
                f"查询窗口步骤失败: path={path} name={step.get('name')}"
            )

    def find_query_window(self, open_query, timeout):
        return self.jab.wait_window_by_title(
            open_query.get("dialog_title", "查询"),
            class_name=open_query.get("dialog_class", "SunAwtDialog"),
            timeout=timeout,
            include_children=bool(open_query.get("dialog_include_children", True)),
            visible_only=bool(open_query.get("dialog_visible_only", True)),
        )

    def open_query_with_jab_action(self, open_query):
        if open_query.get("runner") == "subprocess":
            self.run_jab_action_subprocess(open_query)
            return

        if (
            bool(open_query.get("wait_dialog", True))
            and open_query.get("click_mode") != "bounds"
        ):
            thread = self._trigger_query_action_path(open_query)
            if not thread:
                raise JABControlNotFound(
                    f"JAB 查询入口未找到: path={open_query.get('path')}"
                )
            return

        if not self._do_query_action_path(open_query):
            raise JABActionError(f"JAB 查询入口执行失败: path={open_query.get('path')}")

    def _trigger_query_action_path(self, open_query):
        path = open_query.get("path")
        kwargs = {
            "name": open_query.get("name"),
            "role": open_query.get("role"),
            "action_name": open_query.get("action"),
            "timeout": float(open_query.get("return_timeout", 0.2)),
            "require_showing": bool(open_query.get("require_showing", False)),
        }
        attempts = [
            {
                "title": open_query.get("window_title"),
                "class_name": open_query.get("window_class"),
            },
            {
                "title": open_query.get("main_title"),
                "class_name": open_query.get("main_class"),
            },
            {"title": None, "class_name": None},
        ]
        for attempt in attempts:
            thread = self.jab.trigger_action_by_path_async(path, **kwargs, **attempt)
            if thread:
                return thread
        return None

    def _do_query_action_path(self, open_query):
        path = open_query.get("path")
        kwargs = {
            "name": open_query.get("name"),
            "role": open_query.get("role"),
            "action_name": open_query.get("action"),
            "click_mode": open_query.get("click_mode"),
            "wait": float(open_query.get("wait", 0.8)),
            "timeout": float(open_query.get("timeout", 5)),
            "require_showing": bool(open_query.get("require_showing", False)),
        }
        attempts = [
            {
                "title": open_query.get("window_title"),
                "class_name": open_query.get("window_class"),
            },
            {
                "title": open_query.get("main_title"),
                "class_name": open_query.get("main_class"),
            },
            {"title": None, "class_name": None},
        ]
        for attempt in attempts:
            if self.jab.do_action_by_path(path, **kwargs, **attempt):
                return True
        return False

    def run_jab_action_subprocess(self, step):
        path = step.get("path")
        if not path:
            raise JABControlNotFound("JAB 子进程动作缺少 path")

        timeout = float(step.get("process_timeout", 1.0))
        action_timeout = float(step.get("timeout", 3.0))
        action_wait = float(step.get("action_wait", 0.0))
        script = Path(__file__).resolve().parents[1] / "tools" / "jab_action_once.py"
        cmd = [
            sys.executable,
            str(script),
            "--config",
            self.config_path,
            "--path",
            str(path),
            "--timeout",
            str(action_timeout),
            "--wait",
            str(action_wait),
        ]

        title = step.get("title", step.get("window_title"))
        class_name = step.get("class_name", step.get("window_class"))
        if title is not None:
            cmd.extend(["--title", str(title)])
        if class_name is not None:
            cmd.extend(["--class-name", str(class_name)])
        if step.get("name") is not None:
            cmd.extend(["--name", str(step.get("name"))])
        if step.get("role") is not None:
            cmd.extend(["--role", str(step.get("role"))])
        if step.get("action") is not None:
            cmd.extend(["--action", str(step.get("action"))])
        click_mode = step.get("click_mode")
        if click_mode is not None:
            cmd.extend(["--click-mode", str(click_mode)])

        set_text = step.get("set_text")
        if set_text == "{generated_date_value}":
            set_text = self.generated_date_value
        if set_text is not None:
            cmd.extend(["--set-text", str(set_text)])
        if step.get("set_text_near_label") is not None:
            cmd.extend(["--set-text-near-label", str(step.get("set_text_near_label"))])
        if step.get("guard_path") is not None:
            cmd.extend(["--guard-path", str(step.get("guard_path"))])
        if step.get("guard_name") is not None:
            cmd.extend(["--guard-name", str(step.get("guard_name"))])
        if step.get("guard_role") is not None:
            cmd.extend(["--guard-role", str(step.get("guard_role"))])

        if bool(step.get("require_showing", False)):
            cmd.append("--require-showing")

        log.info(
            "启动 JAB action 子进程: "
            f"path={path} name={step.get('name')} "
            f"set_text={set_text is not None} timeout={timeout}"
        )
        try:
            result = subprocess.run(
                cmd,
                cwd=str(Path(__file__).resolve().parents[1]),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired:
            log.warning(
                "JAB action 子进程超时，按 UI 状态继续判断: "
                f"path={path} name={step.get('name')} timeout={timeout}"
            )
            return False

        output = ((result.stdout or "") + (result.stderr or "")).strip()
        if result.returncode != 0:
            log.warning(
                "JAB action 子进程未确认成功，按 UI 状态继续判断: "
                f"returncode={result.returncode} output={output[:500]}"
            )
            return False

        log.info(f"JAB action 子进程完成: output={output[:500]}")
        return True

    def open_query_with_hotkey(self, open_query):
        self.jab.activate_window_by_title(
            open_query.get("main_title", ""),
            class_name=open_query.get("main_class"),
            timeout=float(open_query.get("timeout", 5)),
        )
        self.jab.press_key(
            open_query.get("key", "f3"),
            wait=float(open_query.get("wait", 0.8)),
        )

    def get_nc_workflow_state(self):
        voucher_exists = self.jab.window_exists(
            self.voucher_window_title,
            class_name=self.voucher_window_class,
        )
        if voucher_exists:
            tables = self.jab.read_window_table_cells(
                self.voucher_window_title,
                max_rows=20,
                max_cols=13,
            )
            voucher_tables = [
                table
                for table in tables
                if table["row_count"] > 0 and table["col_count"] == 13
            ]
            if voucher_tables:
                rows = sum(table["row_count"] for table in voucher_tables)
                log.info(f"NC 状态: 制单子窗口打开，制单表 rows={rows}")
                return "voucher_open"
            log.info("NC 状态: 制单子窗口打开，但制单表为空/不可读")
            return "voucher_empty"

        main_title = self.batch_cfg.get("open_query", {}).get("main_title", "")
        main_class = self.batch_cfg.get("open_query", {}).get("main_class")
        if not main_title or self.jab.window_exists(main_title, class_name=main_class):
            log.info("NC 状态: 已回到制单父界面")
            return "parent_ready"

        log.warning("NC 状态: 未检测到制单子窗口，也未确认父界面")
        return "unknown"
