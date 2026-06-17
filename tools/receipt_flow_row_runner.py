# 职责：编排单行收款单流程——开单/表头/明细/手续费/校验修复/可选保存
# 不做什么：不做 CLI 解析,不读取 Excel,不做批次汇总,不默认触发保存
# 允许依赖层：core JAB/明细组件 + tools.receipt_flow_{entry_state,save,detail_repair,header_diag,report}
# 谁不应该 import：core 层模块不应 import；底层 flow 子模块不应反向 import 本模块

import sys
import threading
import time
from dataclasses import asdict

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


def run_one_row(
    config,
    row,
    save_enabled=False,
    recorder=None,
    pause_after_header_field=None,
    diagnose_header_after_pause=False,
    diagnose_detail_repair=False,
):
    current_stage = {"name": ""}

    def _stage(stage, **fields):
        current_stage["name"] = stage
        if recorder is not None:
            recorder.set_stage(stage, **fields)

    def _event(name, **fields):
        if recorder is not None:
            recorder.event(name, **fields)

    timings = StepTimer()
    flow_started_at = time.perf_counter()
    row_report = {
        "excel_row": row.row,
        "plan_row": serializable(asdict(row)),
        "steps": [],
        "save_enabled": bool(save_enabled),
    }
    business = business_from_plan_row(row)
    jab = _flow.JABOperator(config)
    jab_lock = threading.RLock()
    pipeline_verifier = None
    modal_events = []
    try:
        timings.measure("jab.ensure-started", jab.ensure_started)
        _stage("开单", excel_row=row.row)
        open_step = timings.measure(
            "open.self-made",
            run_with_jab_lock,
            jab_lock,
            _flow.open_self_made_entry,
            config,
            jab,
        )
        row_report["steps"].append({"name": "open-self-made", **open_step})
        if not open_step.get("ok"):
            _event("open-failed", excel_row=row.row, error=open_step.get("reason"))
            return fail(row_report, "open-self-made", timings, open_step.get("reason"))
        row_report["modal_recovery"] = {"events": modal_events}

        def recover_modal_after_failure():
            result = _flow.recover_cancelable_modal_now(
                jab,
                stage=current_stage.get("name") or "",
            )
            if result.get("attempted"):
                modal_events.append(result)
            return result

        entry_scope_hwnd = extract_entry_scope_hwnd(open_step)
        entry_dynamic_index = extract_entry_dynamic_index(open_step)
        entry_anchor_path = extract_entry_anchor_path(open_step)
        if entry_scope_hwnd and entry_dynamic_index is None:
            anchor_retry = timings.measure(
                "header.anchor-retry-current-canvas",
                run_with_jab_lock,
                jab_lock,
                _flow.wait_receipt_header_anchor_in_current_canvas,
                jab,
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
            _event("header-anchor-failed", excel_row=row.row, error=reason)
            return fail(row_report, "header-anchor", timings, reason)
        _stage("表头", excel_row=row.row)
        header_pause_reports = []
        header_steps_so_far_labels = []

        def after_header_field(label, _value, step):
            if label and label not in header_steps_so_far_labels:
                header_steps_so_far_labels.append(label)
            if pause_after_header_field != label:
                return None
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
            if diagnose_header_after_pause:
                report["header_readback"] = diagnose_written_header_fields(
                    jab,
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

        if pause_after_header_field:
            header_steps = timings.measure(
                "header.fill",
                run_with_jab_lock,
                jab_lock,
                _flow.fill_header,
                jab,
                business,
                scope_hwnd=entry_scope_hwnd,
                dynamic_index=entry_dynamic_index,
                anchor_path=entry_anchor_path,
                recover_after_failure=recover_modal_after_failure,
                after_field=after_header_field,
            )
        else:
            header_steps = timings.measure(
                "header.fill",
                run_with_jab_lock,
                jab_lock,
                _flow.fill_header,
                jab,
                business,
                after_field=after_header_field,
                scope_hwnd=entry_scope_hwnd,
                dynamic_index=entry_dynamic_index,
                anchor_path=entry_anchor_path,
                recover_after_failure=recover_modal_after_failure,
            )
        if header_pause_reports:
            row_report["header_pause_diagnostics"] = header_pause_reports
        row_report["header_steps"] = header_steps
        if any(not step.get("ok") for step in header_steps):
            header_error = summarize_header_failure(header_steps)
            _event("header-fill-failed", excel_row=row.row, error=header_error)
            return fail(row_report, "header-fill", timings, header_error)
        customer_name = timings.measure(
            "header.customer-name-readback",
            run_with_jab_lock,
            jab_lock,
            read_customer_name_after_header,
            jab,
            header_steps,
            entry_dynamic_index,
            entry_scope_hwnd,
        )
        row_report["customer_name_readback"] = customer_name
        row_report["nc_customer_name"] = str(customer_name.get("value") or "").strip()
        if not row_report["nc_customer_name"]:
            reason = customer_name.get("reason") or "客户名称未确认"
            _event(
                "header-customer-readback-failed",
                excel_row=row.row,
                error=reason,
            )
            return fail(row_report, "header-customer-name", timings, reason)
        _stage("明细主行", excel_row=row.row)
        located = timings.measure(
            "body.locate",
            run_with_jab_lock,
            jab_lock,
            resolve_body_table_by_dynamic_prefix,
            jab,
            entry_dynamic_index,
            entry_scope_hwnd,
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
            _event("locate-table-failed", excel_row=row.row, error="未定位到明细表")
            return fail(row_report, "locate-body-table", timings, "未定位到明细表")
        row_report["before_table"] = {
            "ok": True,
            "skipped": True,
            "reason": "明细表 path 已定位；后台 pipeline verifier 负责预热 path 和并发读回",
        }
        pipeline_verifier = _flow.DetailPipelineVerifier(
            config,
            located,
            flow_started_at=flow_started_at,
            jab=jab,
            jab_lock=jab_lock,
        )
        pipeline_verifier.start()
        pipeline_field_task_ids = []
        pipeline_field_tasks = {}
        pipeline_snapshot_task_ids = []
        pipeline_row_count_task_id = None

        def submit_detail_verify(row_index, field, business_values, _step):
            task_id = pipeline_verifier.submit_field(
                row_index,
                field,
                business_values,
            )
            pipeline_field_task_ids.append(task_id)
            pipeline_field_tasks[task_id] = {
                "row_index": int(row_index),
                "field": dict(field),
                "business": business_values,
            }
            return task_id

        detail_steps = timings.measure(
            "detail.main-line",
            run_with_jab_lock,
            jab_lock,
            _flow.write_detail_line_by_screen,
            jab,
            business,
            located,
            after_field=submit_detail_verify,
            recover_after_failure=recover_modal_after_failure,
        )
        row_report["detail_steps"] = detail_steps
        pipeline_snapshot_task_ids.append(
            pipeline_verifier.submit_snapshot(
                "after-main-line",
                max_rows=3,
                min_matches=len(detail_steps),
            )
        )
        if not all(step.get("ok") for step in detail_steps):
            _event("detail-main-failed", excel_row=row.row, error="明细主行写入失败")
            return fail(row_report, "detail-main-line", timings, "明细主行写入失败")
        if row.fee > 0:
            _stage("手续费", excel_row=row.row)
            row_report["extra_row_delete"] = {
                "ok": True,
                "skipped": True,
                "reason": "手续费非 0，保留主行后自动带出的第 2 行给手续费覆盖",
            }
            add_row, fee_steps, clear_account, delete_extra = timings.measure(
                "detail.fee-line",
                run_with_jab_lock,
                jab_lock,
                run_fee_only,
                jab,
                located,
                str(row.fee),
                after_field=submit_detail_verify,
                recover_after_failure=recover_modal_after_failure,
            )
            row_report["fee_row_add"] = add_row
            row_report["fee_steps"] = fee_steps
            pipeline_snapshot_task_ids.append(
                pipeline_verifier.submit_snapshot(
                    "after-fee-line",
                    max_rows=4,
                )
            )
            row_report["fee_account_clear"] = clear_account
            row_report["fee_extra_row_delete"] = delete_extra
            if delete_extra.get("ok"):
                pipeline_row_count_task_id = pipeline_verifier.submit_row_count(2)
            if (
                not add_row.get("ok")
                or not all(step.get("ok") for step in fee_steps)
                or not clear_account.get("ok")
                or not delete_extra.get("ok")
            ):
                _event("fee-line-failed", excel_row=row.row, error="手续费行处理失败")
                return fail(row_report, "detail-fee-line", timings, "手续费行处理失败")
        else:
            row_report["extra_row_delete"] = timings.measure(
                "detail.delete-extra-after-main",
                run_with_jab_lock,
                jab_lock,
                _flow.delete_extra_row_if_present,
                jab,
                located,
                1,
            )
            if not row_report["extra_row_delete"].get("ok"):
                _event(
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
            pipeline_row_count_task_id = pipeline_verifier.submit_row_count(1)
        expected_detail_rows = 2 if row.fee > 0 else 1
        pipeline_wait_ids = []
        if pipeline_field_task_ids:
            pipeline_wait_ids.append(pipeline_field_task_ids[-1])
        if pipeline_row_count_task_id:
            pipeline_wait_ids.append(pipeline_row_count_task_id)
        pipeline_wait_started = time.perf_counter()
        row_report["detail_pipeline_verify"] = pipeline_verifier.wait(
            pipeline_wait_ids,
            timeout=2.0,
        )
        if diagnose_detail_repair:
            row_report["detail_pipeline_verify_before_repair_drill"] = dict(
                row_report["detail_pipeline_verify"]
            )
            row_report["detail_pipeline_verify"] = force_one_detail_field_pending(
                row_report["detail_pipeline_verify"],
                pipeline_field_task_ids,
            )
        row_report["detail_pipeline_state"] = verifier_snapshot(pipeline_verifier)
        timings.add(
            "detail.pipeline-final-wait",
            time.perf_counter() - pipeline_wait_started,
        )
        row_report["detail_pipeline_snapshots"] = pipeline_snapshot_task_ids
        detail_pipeline_ok = bool(row_report["detail_pipeline_verify"].get("ok"))
        if not detail_pipeline_ok:
            repair_report = timings.measure(
                "detail.pipeline-repair",
                repair_detail_pipeline_failures,
                jab,
                jab_lock,
                located,
                pipeline_verifier,
                row_report["detail_pipeline_verify"],
                pipeline_field_tasks,
                pipeline_row_count_task_id,
                expected_detail_rows,
                entry_scope_hwnd,
                recover_modal_after_failure,
            )
            row_report["detail_pipeline_repair"] = repair_report
            if repair_report.get("snapshot_task_id"):
                pipeline_snapshot_task_ids.append(repair_report["snapshot_task_id"])
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
                jab_lock,
                _flow.read_body_table,
                jab,
                "after_detail_fill",
                entry_scope_hwnd,
            )
            _event(
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
        account_check = timings.measure(
            "header.account-readback-after-detail",
            run_with_jab_lock,
            jab_lock,
            _flow.wait_header_account_description,
            jab,
            0.0,
            scope=build_header_scope_for_followup(
                entry_scope_hwnd,
                entry_dynamic_index,
            ),
        )
        row_report["header_account"] = account_check
        if not account_check.get("accepted"):
            row_report["header_account_readback_warning"] = {
                "ok": False,
                "reason": "表头收款银行账户未从 JAB 后端读回；明细账号已由后台 pipeline 校验，继续保存/后验查询闭包",
                "account_check": account_check,
            }
            _event(
                "account-readback-warning",
                excel_row=row.row,
                warning="表头收款银行账户未从 JAB 后端读回，继续执行",
            )
        if save_enabled:
            _stage("保存", excel_row=row.row)
            save_result = timings.measure(
                "save.ctrl-s",
                run_with_jab_lock,
                jab_lock,
                _flow.save_receipt_by_ctrl_s,
                jab,
                entry_scope_hwnd,
            )
            if not save_result.get("ok"):
                recovery = timings.measure(
                    "save.modal-recovery-after-failure",
                    recover_modal_after_failure,
                )
                if recovery.get("attempted") and recovery.get("ok"):
                    save_result = timings.measure(
                        "save.ctrl-s-retry-after-modal",
                        run_with_jab_lock,
                        jab_lock,
                        _flow.save_receipt_by_ctrl_s,
                        jab,
                        entry_scope_hwnd,
                    )
                    save_result["retried_after_modal_recovery"] = True
                    save_result["modal_recovery"] = recovery
                else:
                    save_result["modal_recovery"] = recovery
            row_report["save"] = save_result
            if not save_result.get("ok"):
                _event(
                    "save-failed", excel_row=row.row, error=save_result.get("reason")
                )
                return fail(row_report, "save", timings, save_result.get("reason"))
        else:
            row_report["save"] = {
                "ok": True,
                "skipped": True,
                "reason": "no-save 模式：已停在保存前，未触发 Ctrl+S",
            }
        row_report["ok"] = True
        row_report["timings"] = timings.items
        return row_report
    except Exception as exc:
        row_report["exception"] = {
            "type": type(exc).__name__,
            "message": str(exc),
            "stage": current_stage.get("name") or "",
        }
        _event(
            "row-exception",
            excel_row=row.row,
            stage=current_stage.get("name") or "",
            error=f"{type(exc).__name__}: {exc}",
        )
        return fail(
            row_report,
            "exception",
            timings,
            f"{type(exc).__name__}: {exc}",
        )
    finally:
        if pipeline_verifier is not None:
            pipeline_verifier.close(timeout=0.2)
        row_report["modal_recovery"] = {"events": modal_events}
        jab.close()


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
        "settlement": "网银",
        "main_subject": "1002",
        "main_business_type": "货款",
        "fee_subject": "660305",
        "fee_business_type": "手续费",
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
