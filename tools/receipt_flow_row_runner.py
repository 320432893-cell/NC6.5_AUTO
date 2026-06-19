# 职责：编排单行收款单流程——开单/表头/明细/手续费/校验修复/可选保存
# 不做什么：不做 CLI 解析,不读取 Excel,不做批次汇总,不默认触发保存
# 允许依赖层：core JAB/明细组件 + tools.receipt_flow_{entry_state,save,detail_repair,header_diag,report}
# 谁不应该 import：core 层模块不应 import；底层 flow 子模块不应反向 import 本模块

import sys
import threading
import time
from dataclasses import asdict, dataclass, field

from core.receipt_business_constants import (
    RECEIPT_FEE_BUSINESS_TYPE,
    RECEIPT_FEE_SUBJECT,
    RECEIPT_MAIN_BUSINESS_TYPE,
    RECEIPT_MAIN_SUBJECT,
    RECEIPT_SETTLEMENT,
)
from core.runtime_mode import is_engine_mode
from tools.receipt_detail_rows import StepTimer, run_fee_only
from tools.receipt_flow_detail_repair import (
    force_one_detail_field_pending,
    repair_detail_pipeline_failures,
)
from tools.receipt_flow_entry_state import (
    build_header_scope_for_followup,
    extract_entry_anchor_path,
    extract_entry_dynamic_index,
    extract_entry_scope_hwnd,
    resolve_body_table_by_dynamic_prefix,
    run_with_jab_lock,
)
from tools.receipt_flow_header_diag import (
    diagnose_written_header_fields,
    read_customer_name_after_header,
    summarize_header_failure,
)
from tools.receipt_flow_report import serializable


class _FlowNamespace:
    # 按调用时从已加载的入口模块取属性：让测试对 tools.receipt_full_flow_entry 上
    # JABOperator / DetailPipelineVerifier / recover_cancelable_modal_now /
    # fill_header / write_detail_line_by_screen / delete_extra_row_if_present /
    # read_body_table / wait_header_account_description / save_receipt_by_ctrl_s /
    # wait_receipt_header_anchor_in_current_canvas / run_receipt_new_probe[_with_jab]
    # 的 monkeypatch 与拆分前一致地生效，且不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_full_flow_entry"], name)


_flow = _FlowNamespace()


# 阶段间共享的可变状态载体：把原 run_one_row 内层闭包/局部变量集中持有，
# 让 _run_*_stage(state, ...) 顺序流转而不必在阶段间传一长串参数。
@dataclass
class RowRunState:
    config: object
    row: object
    save_enabled: bool
    recorder: object
    pause_after_header_field: object
    diagnose_header_after_pause: bool
    diagnose_detail_repair: bool

    business: object = None
    jab: object = None
    jab_lock: object = None
    timings: object = None
    flow_started_at: float = 0.0
    row_report: dict = field(default_factory=dict)
    current_stage: dict = field(default_factory=lambda: {"name": ""})
    modal_events: list = field(default_factory=list)

    # 阶段间中间产物（开单结果/表头步骤）传递
    open_step: dict = None
    header_steps: list = None

    # 定位/表头锚点跨阶段传递
    entry_scope_hwnd: object = None
    entry_dynamic_index: object = None
    entry_anchor_path: object = None
    located: object = None

    # verifier 及其任务 id 列表（跨阶段提交/等待/修复）
    pipeline_verifier: object = None
    pipeline_field_task_ids: list = field(default_factory=list)
    pipeline_field_tasks: dict = field(default_factory=dict)
    pipeline_snapshot_task_ids: list = field(default_factory=list)
    pipeline_row_count_task_id: object = None
    expected_detail_rows: int = 1

    def stage(self, stage, **fields):
        self.current_stage["name"] = stage
        if self.recorder is not None:
            self.recorder.set_stage(stage, **fields)

    def event(self, name, **fields):
        if self.recorder is not None:
            self.recorder.event(name, **fields)

    def recover_modal_after_failure(self):
        result = _flow.recover_cancelable_modal_now(
            self.jab,
            stage=self.current_stage.get("name") or "",
        )
        if result.get("attempted"):
            self.modal_events.append(result)
        return result

    def submit_detail_verify(self, row_index, field, business_values, _step):
        task_id = self.pipeline_verifier.submit_field(
            row_index,
            field,
            business_values,
        )
        self.pipeline_field_task_ids.append(task_id)
        self.pipeline_field_tasks[task_id] = {
            "row_index": int(row_index),
            "field": dict(field),
            "business": business_values,
        }
        return task_id


def run_one_row(
    config,
    row,
    save_enabled=False,
    recorder=None,
    pause_after_header_field=None,
    diagnose_header_after_pause=False,
    diagnose_detail_repair=False,
):
    state = RowRunState(
        config=config,
        row=row,
        save_enabled=save_enabled,
        recorder=recorder,
        pause_after_header_field=pause_after_header_field,
        diagnose_header_after_pause=diagnose_header_after_pause,
        diagnose_detail_repair=diagnose_detail_repair,
    )
    state.timings = StepTimer()
    state.flow_started_at = time.perf_counter()
    state.row_report = {
        "excel_row": row.row,
        "plan_row": serializable(asdict(row)),
        "steps": [],
        "save_enabled": bool(save_enabled),
    }
    state.business = business_from_plan_row(row)
    state.jab = _flow.JABOperator(config)
    state.jab_lock = threading.RLock()
    try:
        result = _run_open_stage(state)
        if result is not None:
            return result
        result = _run_locate_anchor_stage(state)
        if result is not None:
            return result
        result = _run_header_stage(state)
        if result is not None:
            return result
        result = _run_customer_readback_stage(state)
        if result is not None:
            return result
        result = _run_body_locate_stage(state)
        if result is not None:
            return result
        result = _run_detail_main_stage(state)
        if result is not None:
            return result
        result = _run_fee_stage(state)
        if result is not None:
            return result
        result = _run_row_count_verify_stage(state)
        if result is not None:
            return result
        result = _run_account_readback_stage(state)
        if result is not None:
            return result
        result = _run_save_stage(state)
        if result is not None:
            return result
        state.row_report["ok"] = True
        state.row_report["timings"] = state.timings.items
        return state.row_report
    except Exception as exc:
        state.row_report["exception"] = {
            "type": type(exc).__name__,
            "message": str(exc),
            "stage": state.current_stage.get("name") or "",
        }
        state.event(
            "row-exception",
            excel_row=row.row,
            stage=state.current_stage.get("name") or "",
            error=f"{type(exc).__name__}: {exc}",
        )
        return fail(
            state.row_report,
            "exception",
            state.timings,
            f"{type(exc).__name__}: {exc}",
        )
    finally:
        if state.pipeline_verifier is not None:
            state.pipeline_verifier.close(timeout=0.2)
        state.row_report["modal_recovery"] = {"events": state.modal_events}
        state.jab.close()


def _run_open_stage(state):
    # 阶段1 开单：启动 JAB → 进入收款单自制录入态。失败提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    timings.measure("jab.ensure-started", state.jab.ensure_started)
    state.stage("开单", excel_row=row.row)
    open_step = timings.measure(
        "open.self-made",
        run_with_jab_lock,
        state.jab_lock,
        _flow.open_self_made_entry,
        state.config,
        state.jab,
    )
    row_report["steps"].append({"name": "open-self-made", **open_step})
    if not open_step.get("ok"):
        state.event("open-failed", excel_row=row.row, error=open_step.get("reason"))
        return fail(row_report, "open-self-made", timings, open_step.get("reason"))
    row_report["modal_recovery"] = {"events": state.modal_events}
    state.open_step = open_step
    return None


def _run_locate_anchor_stage(state):
    # 阶段2 表头锚点：从开单结果解析 scope_hwnd/dynamic_index/anchor_path，
    # 必要时在当前 canvas 重试锚点；解析不到财务组织前缀则提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    open_step = state.open_step
    entry_scope_hwnd = extract_entry_scope_hwnd(open_step)
    entry_dynamic_index = extract_entry_dynamic_index(open_step)
    entry_anchor_path = extract_entry_anchor_path(open_step)
    if entry_scope_hwnd and entry_dynamic_index is None:
        anchor_retry = timings.measure(
            "header.anchor-retry-current-canvas",
            run_with_jab_lock,
            state.jab_lock,
            _flow.wait_receipt_header_anchor_in_current_canvas,
            state.jab,
            entry_scope_hwnd,
            timeout=1.2,
            interval=0.2,
        )
        row_report["entry_header_anchor_retry"] = anchor_retry
        if anchor_retry.get("ok"):
            entry_dynamic_index = anchor_retry.get("dynamic_index")
            entry_anchor_path = anchor_retry.get("label_path") or entry_anchor_path
    row_report["entry_scope_hwnd"] = entry_scope_hwnd
    row_report["entry_dynamic_index"] = entry_dynamic_index
    row_report["entry_anchor_path"] = entry_anchor_path
    row_report["locator_policy"] = {
        "header": (
            "财务组织用控件名粘贴+Enter并确认中文；其它表头字段优先用动态前缀"
            "+稳定后缀 path，path 失败才单字段语义兜底"
        ),
        "body": "明细表用校正后的动态前缀直接拼稳定 path，校验失败才现场扫描",
    }
    if not entry_scope_hwnd or entry_dynamic_index is None:
        reason = "当前 canvas 未解析到财务组织(O) 前缀，停止；不走语义兜底"
        state.event("header-anchor-failed", excel_row=row.row, error=reason)
        return fail(row_report, "header-anchor", timings, reason)
    state.entry_scope_hwnd = entry_scope_hwnd
    state.entry_dynamic_index = entry_dynamic_index
    state.entry_anchor_path = entry_anchor_path
    return None


def _run_header_stage(state):
    # 阶段3 表头：填表头字段（可选诊断暂停 after_field 钩子），任一字段失败提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    entry_scope_hwnd = state.entry_scope_hwnd
    entry_dynamic_index = state.entry_dynamic_index
    entry_anchor_path = state.entry_anchor_path
    state.stage("表头", excel_row=row.row)
    header_pause_reports = []
    header_steps_so_far_labels = []

    def after_header_field(label, _value, step):
        if label and label not in header_steps_so_far_labels:
            header_steps_so_far_labels.append(label)
        if state.pause_after_header_field != label:
            return None
        # 引擎模式(子进程无 TTY)下不打印旁白、不 input()，避免阻塞挂起；CLI 诊断保留交互。
        if not is_engine_mode():
            print(
                f"诊断暂停：表头字段 [{label}] 已写入。"
                "可以现在人工打开干扰窗口或清理已输入字段；回车后继续检查。"
            )
            input("完成干扰后按回车继续: ")
        report = {
            "ok": True,
            "paused_after": label,
            "field_path": step.get("path"),
        }
        if state.diagnose_header_after_pause:
            report["header_readback"] = diagnose_written_header_fields(
                state.jab,
                list(header_steps_so_far_labels),
                step.get("dynamic_index"),
                step.get("dynamic_prefix"),
                entry_scope_hwnd,
            )
            report["ok"] = all(
                item.get("present") for item in report["header_readback"]
            )
            if not report["ok"]:
                report["reason"] = "暂停恢复后检测到已写表头字段为空或不可读"
        header_pause_reports.append(report)
        return report

    if state.pause_after_header_field:
        header_steps = timings.measure(
            "header.fill",
            run_with_jab_lock,
            state.jab_lock,
            _flow.fill_header,
            state.jab,
            state.business,
            scope_hwnd=entry_scope_hwnd,
            dynamic_index=entry_dynamic_index,
            anchor_path=entry_anchor_path,
            recover_after_failure=state.recover_modal_after_failure,
            after_field=after_header_field,
        )
    else:
        header_steps = timings.measure(
            "header.fill",
            run_with_jab_lock,
            state.jab_lock,
            _flow.fill_header,
            state.jab,
            state.business,
            after_field=after_header_field,
            scope_hwnd=entry_scope_hwnd,
            dynamic_index=entry_dynamic_index,
            anchor_path=entry_anchor_path,
            recover_after_failure=state.recover_modal_after_failure,
        )
    if header_pause_reports:
        row_report["header_pause_diagnostics"] = header_pause_reports
    row_report["header_steps"] = header_steps
    if any(not step.get("ok") for step in header_steps):
        header_error = summarize_header_failure(header_steps)
        state.event("header-fill-failed", excel_row=row.row, error=header_error)
        return fail(row_report, "header-fill", timings, header_error)
    state.header_steps = header_steps
    return None


def _run_customer_readback_stage(state):
    # 阶段4 客户回读：从 JAB 后端读回客户名称确认；未确认则提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    customer_name = timings.measure(
        "header.customer-name-readback",
        run_with_jab_lock,
        state.jab_lock,
        read_customer_name_after_header,
        state.jab,
        state.header_steps,
        state.entry_dynamic_index,
        state.entry_scope_hwnd,
    )
    row_report["customer_name_readback"] = customer_name
    row_report["nc_customer_name"] = str(customer_name.get("value") or "").strip()
    if not row_report["nc_customer_name"]:
        reason = customer_name.get("reason") or "客户名称未确认"
        state.event(
            "header-customer-readback-failed",
            excel_row=row.row,
            error=reason,
        )
        return fail(row_report, "header-customer-name", timings, reason)
    return None


def _run_body_locate_stage(state):
    # 阶段5a 定位：用动态前缀解析明细表 path，并启动后台 pipeline verifier。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    state.stage("明细主行", excel_row=row.row)
    located = timings.measure(
        "body.locate",
        run_with_jab_lock,
        state.jab_lock,
        resolve_body_table_by_dynamic_prefix,
        state.jab,
        state.entry_dynamic_index,
        state.entry_scope_hwnd,
    )
    row_report["body_locate"] = {
        "source": located.get("source"),
        "cache_hit": located.get("cache_hit"),
        "fallback_used": located.get("fallback_used"),
        "status": located.get("status"),
        "seconds": located.get("seconds"),
        "started_offset_seconds": located.get("started_offset_seconds"),
    }
    row_report["table_candidates"] = located.get("candidates", [])[:5]
    if not located.get("best"):
        state.event("locate-table-failed", excel_row=row.row, error="未定位到明细表")
        return fail(row_report, "locate-body-table", timings, "未定位到明细表")
    row_report["before_table"] = {
        "ok": True,
        "skipped": True,
        "reason": "明细表 path 已定位；后台 pipeline verifier 负责预热 path 和并发读回",
    }
    state.located = located
    state.pipeline_verifier = _flow.DetailPipelineVerifier(
        state.config,
        located,
        flow_started_at=state.flow_started_at,
        jab=state.jab,
        jab_lock=state.jab_lock,
    )
    state.pipeline_verifier.start()
    return None


def _run_detail_main_stage(state):
    # 阶段5b 明细主行：写主行各字段（after_field 提交 verifier 任务）+ 提交快照；
    # 任一字段写入失败则提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    detail_steps = timings.measure(
        "detail.main-line",
        run_with_jab_lock,
        state.jab_lock,
        _flow.write_detail_line_by_screen,
        state.jab,
        state.business,
        state.located,
        after_field=state.submit_detail_verify,
        recover_after_failure=state.recover_modal_after_failure,
    )
    row_report["detail_steps"] = detail_steps
    state.pipeline_snapshot_task_ids.append(
        state.pipeline_verifier.submit_snapshot(
            "after-main-line",
            max_rows=3,
            min_matches=len(detail_steps),
        )
    )
    if not all(step.get("ok") for step in detail_steps):
        state.event("detail-main-failed", excel_row=row.row, error="明细主行写入失败")
        return fail(row_report, "detail-main-line", timings, "明细主行写入失败")
    return None


def _run_fee_stage(state):
    # 阶段6 手续费：fee>0 时新增手续费行并清账户+删多余行（提交快照/行数任务）；
    # fee==0 时删主行后多余行并提交行数=1。任一步骤失败提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    if row.fee > 0:
        state.stage("手续费", excel_row=row.row)
        row_report["extra_row_delete"] = {
            "ok": True,
            "skipped": True,
            "reason": "手续费非 0，保留主行后自动带出的第 2 行给手续费覆盖",
        }
        add_row, fee_steps, clear_account, delete_extra = timings.measure(
            "detail.fee-line",
            run_with_jab_lock,
            state.jab_lock,
            run_fee_only,
            state.jab,
            state.located,
            str(row.fee),
            after_field=state.submit_detail_verify,
            recover_after_failure=state.recover_modal_after_failure,
        )
        row_report["fee_row_add"] = add_row
        row_report["fee_steps"] = fee_steps
        state.pipeline_snapshot_task_ids.append(
            state.pipeline_verifier.submit_snapshot(
                "after-fee-line",
                max_rows=4,
            )
        )
        row_report["fee_account_clear"] = clear_account
        row_report["fee_extra_row_delete"] = delete_extra
        if delete_extra.get("ok"):
            state.pipeline_row_count_task_id = state.pipeline_verifier.submit_row_count(
                2
            )
        if (
            not add_row.get("ok")
            or not all(step.get("ok") for step in fee_steps)
            or not clear_account.get("ok")
            or not delete_extra.get("ok")
        ):
            state.event("fee-line-failed", excel_row=row.row, error="手续费行处理失败")
            return fail(row_report, "detail-fee-line", timings, "手续费行处理失败")
    else:
        row_report["extra_row_delete"] = timings.measure(
            "detail.delete-extra-after-main",
            run_with_jab_lock,
            state.jab_lock,
            _flow.delete_extra_row_if_present,
            state.jab,
            state.located,
            1,
        )
        if not row_report["extra_row_delete"].get("ok"):
            state.event(
                "delete-extra-failed",
                excel_row=row.row,
                error="主行后多余行删除失败",
            )
            return fail(
                row_report,
                "detail-delete-extra-after-main",
                timings,
                "主行后多余行删除失败",
            )
        row_report["fee_skipped"] = {
            "ok": True,
            "reason": "手续费为 0，跳过手续费行",
        }
        state.pipeline_row_count_task_id = state.pipeline_verifier.submit_row_count(1)
    return None


def _run_row_count_verify_stage(state):
    # 阶段7 行数/字段验证：等待最后字段+行数 verifier 任务，必要时触发修复并复核；
    # 仍不通过则做整表读 fallback 并提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    pipeline_verifier = state.pipeline_verifier
    state.expected_detail_rows = 2 if row.fee > 0 else 1
    pipeline_wait_ids = []
    if state.pipeline_field_task_ids:
        pipeline_wait_ids.append(state.pipeline_field_task_ids[-1])
    if state.pipeline_row_count_task_id:
        pipeline_wait_ids.append(state.pipeline_row_count_task_id)
    pipeline_wait_started = time.perf_counter()
    row_report["detail_pipeline_verify"] = pipeline_verifier.wait(
        pipeline_wait_ids,
        timeout=2.0,
    )
    if state.diagnose_detail_repair:
        row_report["detail_pipeline_verify_before_repair_drill"] = dict(
            row_report["detail_pipeline_verify"]
        )
        row_report["detail_pipeline_verify"] = force_one_detail_field_pending(
            row_report["detail_pipeline_verify"],
            state.pipeline_field_task_ids,
        )
    row_report["detail_pipeline_state"] = verifier_snapshot(pipeline_verifier)
    timings.add(
        "detail.pipeline-final-wait",
        time.perf_counter() - pipeline_wait_started,
    )
    row_report["detail_pipeline_snapshots"] = state.pipeline_snapshot_task_ids
    detail_pipeline_ok = bool(row_report["detail_pipeline_verify"].get("ok"))
    if not detail_pipeline_ok:
        repair_report = timings.measure(
            "detail.pipeline-repair",
            repair_detail_pipeline_failures,
            state.jab,
            state.jab_lock,
            state.located,
            pipeline_verifier,
            row_report["detail_pipeline_verify"],
            state.pipeline_field_tasks,
            state.pipeline_row_count_task_id,
            state.expected_detail_rows,
            state.entry_scope_hwnd,
            state.recover_modal_after_failure,
        )
        row_report["detail_pipeline_repair"] = repair_report
        if repair_report.get("snapshot_task_id"):
            state.pipeline_snapshot_task_ids.append(repair_report["snapshot_task_id"])
        repair_wait_ids = repair_report.get("wait_ids") or []
        if repair_wait_ids:
            repair_wait_started = time.perf_counter()
            row_report["detail_pipeline_verify_after_repair"] = (
                pipeline_verifier.wait(
                    repair_wait_ids,
                    timeout=2.0,
                )
            )
            row_report["detail_pipeline_state_after_repair"] = verifier_snapshot(
                pipeline_verifier
            )
            timings.add(
                "detail.pipeline-repair-wait",
                time.perf_counter() - repair_wait_started,
            )
            detail_pipeline_ok = bool(
                row_report["detail_pipeline_verify_after_repair"].get("ok")
            )
    if not detail_pipeline_ok:
        row_report["after_table"] = timings.measure(
            "body.read-after-fallback",
            run_with_jab_lock,
            state.jab_lock,
            _flow.read_body_table,
            state.jab,
            "after_detail_fill",
            state.entry_scope_hwnd,
        )
        state.event(
            "pipeline-verify-failed",
            excel_row=row.row,
            error="后台明细验证未通过，已执行整表读 fallback",
        )
        return fail(
            row_report,
            "detail-pipeline-verify",
            timings,
            "后台明细验证未通过，已执行整表读 fallback",
        )
    row_report["after_table"] = {
        "ok": True,
        "skipped": True,
        "reason": "后台 pipeline verifier 已覆盖最后字段与最终行数，跳过同步整表读",
    }
    return None


def _run_account_readback_stage(state):
    # 阶段8 账户回读：从 JAB 后端读回表头收款银行账户；未读回只记 warning 不阻断。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    account_check = timings.measure(
        "header.account-readback-after-detail",
        run_with_jab_lock,
        state.jab_lock,
        _flow.wait_header_account_description,
        state.jab,
        0.0,
        scope=build_header_scope_for_followup(
            state.entry_scope_hwnd,
            state.entry_dynamic_index,
        ),
    )
    row_report["header_account"] = account_check
    if not account_check.get("accepted"):
        row_report["header_account_readback_warning"] = {
            "ok": False,
            "reason": "表头收款银行账户未从 JAB 后端读回；明细账号已由后台 pipeline 校验，继续保存/后验查询闭包",
            "account_check": account_check,
        }
        state.event(
            "account-readback-warning",
            excel_row=row.row,
            warning="表头收款银行账户未从 JAB 后端读回，继续执行",
        )
    return None


def _run_save_stage(state):
    # 阶段9 保存：save_enabled 时 Ctrl+S（失败做弹窗恢复后重试一次）；
    # no-save 模式则记录 skipped。保存失败提前返回 fail。
    row = state.row
    timings = state.timings
    row_report = state.row_report
    if state.save_enabled:
        state.stage("保存", excel_row=row.row)
        save_result = timings.measure(
            "save.ctrl-s",
            run_with_jab_lock,
            state.jab_lock,
            _flow.save_receipt_by_ctrl_s,
            state.jab,
            state.entry_scope_hwnd,
        )
        if not save_result.get("ok"):
            recovery = timings.measure(
                "save.modal-recovery-after-failure",
                state.recover_modal_after_failure,
            )
            if recovery.get("attempted") and recovery.get("ok"):
                save_result = timings.measure(
                    "save.ctrl-s-retry-after-modal",
                    run_with_jab_lock,
                    state.jab_lock,
                    _flow.save_receipt_by_ctrl_s,
                    state.jab,
                    state.entry_scope_hwnd,
                )
                save_result["retried_after_modal_recovery"] = True
                save_result["modal_recovery"] = recovery
            else:
                save_result["modal_recovery"] = recovery
        row_report["save"] = save_result
        if not save_result.get("ok"):
            state.event(
                "save-failed", excel_row=row.row, error=save_result.get("reason")
            )
            return fail(row_report, "save", timings, save_result.get("reason"))
    else:
        row_report["save"] = {
            "ok": True,
            "skipped": True,
            "reason": "no-save 模式：已停在保存前，未触发 Ctrl+S",
        }
    return None


def open_self_made_entry(config, jab=None):
    opened = (
        _flow.run_receipt_new_probe_with_jab(jab)
        if jab is not None
        else _flow.run_receipt_new_probe()
    )
    if opened.get("ok"):
        return opened
    opened["reason"] = opened.get("reason") or "未能进入收款单自制录入态"
    return opened


def verifier_snapshot(verifier):
    if verifier is None or not hasattr(verifier, "snapshot"):
        return None
    return verifier.snapshot()


def business_from_plan_row(row):
    return {
        "finance_org_code": row.organization_code,
        "finance_org_name": row.organization_name,
        "document_date": row.receipt_date.isoformat(),
        "customer_code": row.customer_code,
        "currency": row.currency,
        "header_currency_code": row.header_currency_code or row.currency,
        "bank_label": row.bank,
        "bank_account": row.account_no,
        "amount": str(row.raw_amount),
        "fee": str(row.fee),
        "has_fee": row.fee > 0,
        "settlement": RECEIPT_SETTLEMENT,
        "main_subject": RECEIPT_MAIN_SUBJECT,
        "main_business_type": RECEIPT_MAIN_BUSINESS_TYPE,
        "fee_subject": RECEIPT_FEE_SUBJECT,
        "fee_business_type": RECEIPT_FEE_BUSINESS_TYPE,
    }


def fail(row_report, failed_step, timings, reason):
    row_report.update(
        {
            "ok": False,
            "failed_step": failed_step,
            "reason": reason,
            "timings": timings.items,
        }
    )
    return row_report
