# 职责：收款单表头字段写入门面——fill_header 编排、动态字段提交、剪贴板守护粘贴;
#       并 re-export 已抽出的财务组织专用写法(finance_org)、字段定位 finders(locator)、后端字段纯状态判定(state),对外 API 零变化
# 不做什么：不解析 CLI/不读 Excel,不做探针报告,不做 scope 锚点解析(委托 scope 模块),
#           不再自实现财务组织写法(委托 finance_org)、字段定位(委托 locator)与纯状态判定(委托 state)
# 允许依赖层：core/JAB + tools.receipt_header_{paths,tree,scope,state,finance_org,locator} + tools.receipt_keyboard_utils;
#             被 monkeypatch 的协作者(同/跨模块)经 _trial 代理读 tools.receipt_self_made_flow
# 谁不应该 import：receipt_header_{paths,tree,state,finance_org,locator} 不应 import 本模块(会成环)

import sys

import ctypes  # noqa: F401  # 供搬移函数内 ctypes.* 使用

from tools.receipt_header_paths import (
    HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
    HEADER_SCOPE_ANCHOR_LABEL,
    accepted_text_from_backend,
    clear_receipt_header_path_template_cache,
    infer_header_path_template_from_field,
    receipt_header_dynamic_prefix,
    set_receipt_header_path_template,
)
from tools.receipt_header_scope import resolve_receipt_header_scope

# 抽出的兄弟模块：原文件在此 import 回来并 re-export,保持对外 API(及测试 monkeypatch 面)零变化。
from tools.receipt_header_finance_org import (  # noqa: F401
    confirm_finance_org_accepted,
    probe_finance_org_accepted_text_in_scope,
    set_finance_org_by_legacy_control_name,
)
from tools.receipt_header_locator import (  # noqa: F401
    find_receipt_header_field_by_dynamic_path,
    find_receipt_header_field_by_live_semantic,
    find_receipt_header_field_by_semantic_label,
)
from tools.receipt_header_state import (  # noqa: F401
    backend_field_accepts,
    backend_field_has_written_value,
    context_contains,
    describe_backend_field_state,
)

import time  # noqa: E402,F401  # 供搬移函数内 time.* 使用;测试经 _trial.time.sleep monkeypatch


class _TrialNamespace:
    # 按调用时从已加载的入口模块 tools.receipt_self_made_flow 取属性：
    # 让测试对 trial 上 set_receipt_header_dynamic_field / guarded_paste_header_value /
    # find_receipt_header_field_by_* / set_finance_org_by_legacy_control_name /
    # probe_finance_org_accepted_text_in_scope / locate_receipt_header_scope /
    # find_context_with_window / find_header_label_context_with_window /
    # 及剪贴板/热键工具(get/set/restore_clipboard_text, send_hotkey_ctrl_a/v,
    # foreground_matches_window)的 monkeypatch 与拆分前一致地生效；
    # 不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_self_made_flow"], name)


_trial = _TrialNamespace()


def fill_header(
    jab,
    business,
    after_field=None,
    scope_hwnd=None,
    dynamic_index=None,
    anchor_path=None,
    recover_after_failure=None,
):
    started_at = time.perf_counter()
    steps = []
    scope_started_at = time.perf_counter()
    scope = resolve_receipt_header_scope(jab, scope_hwnd, dynamic_index, anchor_path)
    scope_seconds = round(time.perf_counter() - scope_started_at, 3)
    if not scope.get("ok"):
        return [
            {
                "step": "blocked",
                "reason": "header scope not resolved",
                "scope": scope,
                "header_timing": {
                    "scope_seconds": scope_seconds,
                    "total_seconds": round(time.perf_counter() - started_at, 3),
                },
            }
        ]
    dynamic_index = scope.get("dynamic_index")
    scope_hwnd = scope.get("scope_hwnd")
    learned_header_template = None
    if dynamic_index is not None:
        clear_receipt_header_path_template_cache()
    for field in [
        {
            "label": "财务组织",
            "value": business["finance_org_code"],
            "accepted_text": business.get("finance_org_name"),
            "dynamic_path": True,
        },
        {
            "label": "客户",
            "value": business["customer_code"],
            "dynamic_path": True,
        },
        {
            "label": "单据日期",
            "value": business["document_date"],
            "dynamic_path": True,
        },
        {
            "label": "币种",
            "value": business.get("header_currency_code") or business.get("currency"),
            "dynamic_path": True,
        },
        {
            "label": "结算方式",
            "value": business.get("settlement") or "网银",
            "dynamic_path": True,
        },
    ]:
        label = field["label"]
        value = field["value"]
        if field.get("dynamic_path"):
            field_started_at = time.perf_counter()
            result = _trial.set_receipt_header_dynamic_field(
                jab,
                label,
                value,
                dynamic_index,
                scope_hwnd,
                accepted_text=field.get("accepted_text"),
                recover_after_failure=recover_after_failure,
                path_template=learned_header_template,
            )
            ok = bool(result.get("ok"))
            if ok and label != HEADER_SCOPE_ANCHOR_LABEL:
                inferred_template = infer_header_path_template_from_field(
                    result.get("path"),
                    dynamic_index,
                    label,
                )
                if inferred_template and learned_header_template is None:
                    learned_header_template = inferred_template
                    set_receipt_header_path_template(dynamic_index, inferred_template)
                    result["header_path_template_learned"] = learned_header_template
            field_seconds = round(time.perf_counter() - field_started_at, 3)
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "dynamic_path_commit",
                    "scope": {
                        "mode": scope.get("mode"),
                        "scope_hwnd": scope_hwnd,
                        "dynamic_index": dynamic_index,
                        "dynamic_prefix": scope.get("dynamic_prefix"),
                    },
                    **result,
                    "seconds": field_seconds,
                    "header_timing": {
                        "scope_seconds": scope_seconds,
                        "elapsed_seconds": round(time.perf_counter() - started_at, 3),
                    },
                }
            )
        else:
            ok = False
            steps.append(
                {
                    "step": "header",
                    "label": label,
                    "value": value,
                    "method": "unsupported",
                    "ok": False,
                    "reason": "header field must use dynamic path",
                }
            )
        if not ok:
            steps.append(
                {
                    "step": "blocked",
                    "reason": "header field failed; blocking workflow stopped",
                    "label": label,
                }
            )
            break
        if after_field is not None:
            callback_report = after_field(label, value, steps[-1])
            if callback_report is not None:
                steps[-1]["after_field_callback"] = callback_report
                if not callback_report.get("ok", True):
                    steps.append(
                        {
                            "step": "blocked",
                            "reason": callback_report.get("reason")
                            or "header after-field callback blocked workflow",
                            "label": label,
                        }
                    )
                    break
    return steps


def wait_header_account_description(jab, timeout=5.0, scope=None):
    deadline = time.time() + max(float(timeout or 0), 0.0)
    last = None
    if scope is None:
        scope = _trial.locate_receipt_header_scope(jab)
    if not scope.get("ok"):
        return {"text": None, "description": "", "accepted": False, "scope": scope}
    dynamic_index = scope.get("dynamic_index")
    scope_hwnd = scope.get("scope_hwnd")
    first_attempt = True
    while first_attempt or time.time() < deadline:
        first_attempt = False
        found = _trial.find_receipt_header_field_by_dynamic_path(
            jab,
            "收款银行账户",
            dynamic_index,
            scope_hwnd=scope_hwnd,
            require_showing=False,
            require_valid_bounds=False,
        )
        if found.get("ok"):
            context = found["context"]
            vm_id = found["vm_id"]
            owned_contexts = found["owned_contexts"]
            try:
                info = jab.get_context_info(vm_id, context)
                text = jab.get_text_context_value(vm_id, context)
                desc = info.description.strip() if info else ""
                last = {"text": text, "description": desc, "path": found.get("path")}
                if text or desc:
                    last["accepted"] = True
                    return last
            finally:
                jab.release_contexts(vm_id, owned_contexts)
        if time.time() >= deadline:
            break
        time.sleep(0.3)
    if last is None:
        last = {"text": None, "description": "", "accepted": False}
    else:
        last["accepted"] = False
    return last


def set_receipt_header_dynamic_field(
    jab,
    label,
    value,
    dynamic_index,
    scope_hwnd,
    accepted_text=None,
    recover_after_failure=None,
    path_template=None,
):
    legacy_control_name_attempt = None
    if label == HEADER_SCOPE_ANCHOR_LABEL:
        legacy_control_name_attempt = _trial.set_finance_org_by_legacy_control_name(
            jab,
            value,
            scope_hwnd=scope_hwnd,
            accepted_text=accepted_text,
        )
        return {
            **legacy_control_name_attempt,
            "dynamic_index": dynamic_index,
            "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        }

    def find_by_path():
        path_found = _trial.find_receipt_header_field_by_dynamic_path(
            jab,
            label,
            dynamic_index,
            scope_hwnd=scope_hwnd,
            require_showing=False,
            require_valid_bounds=False,
            path_template=path_template,
        )
        if path_found.get("ok"):
            return path_found
        live_found = _trial.find_receipt_header_field_by_live_semantic(
            jab,
            label,
            scope_hwnd=scope_hwnd,
            timeout=HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
        )
        if live_found.get("ok"):
            live_found["dynamic_path_attempt"] = path_found
            return live_found
        return {
            **live_found,
            "dynamic_path_attempt": path_found,
            "live_semantic_timeout": HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
        }

    recovery_after_find = None
    try:
        found = find_by_path()
    except Exception as exc:
        if recover_after_failure is None:
            raise
        recovery_after_find = recover_after_failure()
        if recovery_after_find.get("attempted") and recovery_after_find.get("ok"):
            try:
                found = find_by_path()
            except Exception as retry_exc:
                return {
                    "ok": False,
                    "stage": "resolve",
                    "exception": f"{type(retry_exc).__name__}: {retry_exc}",
                    "first_exception": f"{type(exc).__name__}: {exc}",
                    "modal_recovery": recovery_after_find,
                }
        else:
            return {
                "ok": False,
                "stage": "resolve",
                "exception": f"{type(exc).__name__}: {exc}",
                "modal_recovery": recovery_after_find,
            }
    if not found.get("ok") and recover_after_failure is not None:
        recovery_after_find = recover_after_failure()
        if recovery_after_find.get("attempted") and recovery_after_find.get("ok"):
            found = find_by_path()
    path_attempt = found
    if not found.get("ok"):
        return {
            "ok": False,
            "stage": "resolve",
            "path_attempt": path_attempt,
            "modal_recovery": recovery_after_find,
            "legacy_control_name_attempt": legacy_control_name_attempt,
        }
    context = found["context"]
    vm_id = found["vm_id"]
    owned_contexts = found["owned_contexts"]
    window_info = found["window"]

    def write_current_context(initial_recovery=None):
        info_before = jab.get_context_info(vm_id, context)
        before = jab.get_text_context_value(vm_id, context)
        modal_recovery = initial_recovery
        guarded_paste = _trial.guarded_paste_header_value(
            jab,
            vm_id,
            context,
            window_info,
            value,
        )
        set_ok = bool(guarded_paste.get("ok"))
        if not set_ok and recover_after_failure is not None:
            recovery_after_set = recover_after_failure()
            modal_recovery = recovery_after_set or modal_recovery
            if recovery_after_set.get("attempted") and recovery_after_set.get("ok"):
                guarded_paste = _trial.guarded_paste_header_value(
                    jab,
                    vm_id,
                    context,
                    window_info,
                    value,
                )
                set_ok = bool(guarded_paste.get("ok"))
        if set_ok and hasattr(jab.dll, "requestFocus"):
            jab.dll.requestFocus(vm_id, context)
        enter_ok = (
            bool(guarded_paste.get("enter_ok"))
            if guarded_paste and guarded_paste.get("ok")
            else False
        )
        info_after = jab.get_context_info(vm_id, context)
        after = jab.get_text_context_value(vm_id, context)
        backend_state = describe_backend_field_state(
            info_after,
            after,
            value=value,
            accepted_text=accepted_text,
        )
        return {
            "ok": bool(set_ok),
            "path": found.get("path"),
            "label_path": found.get("label_path"),
            "dynamic_index": found.get("dynamic_index"),
            "dynamic_prefix": found.get("dynamic_prefix"),
            "source": found.get("source") or "path",
            "path_attempt": path_attempt,
            "dynamic_path_attempt": found.get("dynamic_path_attempt"),
            "text_before": before,
            "description_before": (
                info_before.description.strip() if info_before else None
            ),
            "set_ok": bool(set_ok),
            "guarded_paste": guarded_paste,
            "legacy_control_name_attempt": legacy_control_name_attempt,
            "modal_recovery": modal_recovery,
            "enter_ok": bool(enter_ok),
            "post_write_snapshot": backend_state,
            "accepted_text": accepted_text_from_backend(
                backend_state,
                value,
                accepted_text,
            ),
            "text_after": after,
            "description_after": (
                info_after.description.strip() if info_after else None
            ),
        }

    try:
        try:
            return write_current_context(recovery_after_find)
        except Exception as exc:
            if recover_after_failure is None:
                raise
            recovery_after_exception = recover_after_failure()
            if recovery_after_exception.get(
                "attempted"
            ) and recovery_after_exception.get("ok"):
                try:
                    retried = write_current_context(recovery_after_exception)
                except Exception as retry_exc:
                    return {
                        "ok": False,
                        "stage": "write",
                        "exception": f"{type(retry_exc).__name__}: {retry_exc}",
                        "first_exception": f"{type(exc).__name__}: {exc}",
                        "path": found.get("path"),
                        "label_path": found.get("label_path"),
                        "dynamic_index": found.get("dynamic_index"),
                        "dynamic_prefix": found.get("dynamic_prefix"),
                        "modal_recovery": recovery_after_exception,
                        "legacy_control_name_attempt": legacy_control_name_attempt,
                    }
                retried["retried_after_modal_recovery"] = True
                return retried
            return {
                "ok": False,
                "stage": "write",
                "exception": f"{type(exc).__name__}: {exc}",
                "path": found.get("path"),
                "label_path": found.get("label_path"),
                "dynamic_index": found.get("dynamic_index"),
                "dynamic_prefix": found.get("dynamic_prefix"),
                "modal_recovery": recovery_after_exception,
                "legacy_control_name_attempt": legacy_control_name_attempt,
            }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def guarded_paste_header_value(jab, vm_id, context, window_info, value):
    focus_ok = True
    if hasattr(jab.dll, "requestFocus"):
        focus_ok = bool(jab.dll.requestFocus(vm_id, context))
    if not focus_ok:
        return {
            "ok": False,
            "method": "guarded-clipboard-paste",
            "reason": "JAB requestFocus 失败，未发送剪贴板输入",
            "focus_ok": False,
        }
    guard = _trial.foreground_matches_window({"hwnd": (window_info or {}).get("hwnd")})
    if not guard.get("ok"):
        return {
            **guard,
            "ok": False,
            "method": "guarded-clipboard-paste",
            "reason": f"{guard.get('reason')}，未发送表头剪贴板输入",
            "focus_ok": True,
        }
    old_clipboard = _trial.get_clipboard_text()
    result = None
    try:
        _trial.set_clipboard_text(str(value))
        time.sleep(0.02)
        _trial.send_hotkey_ctrl_a()
        time.sleep(0.02)
        _trial.send_hotkey_ctrl_v()
        time.sleep(0.02)
        jab.press_key("enter", wait=0)
        result = {
            **guard,
            "ok": True,
            "method": "guarded-clipboard-paste",
            "focus_ok": True,
            "enter_ok": True,
            "enter_method": "jab.press_key",
        }
        return result
    finally:
        try:
            clipboard_restored = _trial.restore_clipboard_text(old_clipboard)
        except Exception:
            clipboard_restored = False
        if result is not None:
            result["clipboard_restored"] = clipboard_restored
