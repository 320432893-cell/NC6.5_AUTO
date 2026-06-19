# 职责：收款单表头字段写入——fill_header 编排、动态字段提交、财务组织专用写法、剪贴板粘贴、字段定位与回读校验
# 不做什么：不解析 CLI/不读 Excel,不做探针报告,不做 scope 锚点解析(委托 scope 模块)
# 允许依赖层：core/JAB + tools.receipt_header_{paths,tree,scope} + tools.receipt_keyboard_utils;
#             被 monkeypatch 的协作者(同/跨模块)经 _trial 代理读 tools.receipt_self_made_flow
# 谁不应该 import：receipt_header_{paths,tree} 不应 import 本模块(会成环)

import sys

import ctypes  # noqa: F401  # 供搬移函数内 ctypes.* 使用

from tools.receipt_header_paths import (
    FINANCE_ORG_ACCEPTED_TEXT,
    HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
    HEADER_SCOPE_ANCHOR_LABEL,
    HEADER_SCOPE_ANCHOR_TEXT,
    accepted_text_from_backend,
    build_receipt_header_dynamic_label_path,
    build_receipt_header_dynamic_path,
    build_receipt_header_label_path_from_template,
    build_receipt_header_path_from_template,
    clear_receipt_header_path_template_cache,
    infer_header_path_template_from_field,
    infer_header_text_path_from_label_path,
    receipt_header_dynamic_prefix,
    set_receipt_header_path_template,
)
from tools.receipt_header_scope import resolve_receipt_header_scope

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


def find_receipt_header_field_by_live_semantic(
    jab,
    label,
    scope_hwnd=None,
    timeout=HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT,
):
    found = _trial.find_receipt_header_field_by_semantic_label(
        jab,
        label,
        scope_hwnd=scope_hwnd,
        timeout=timeout,
    )
    if found.get("ok"):
        found["source"] = "semantic-live-after-path-miss"
        return found
    return {
        **found,
        "source": "semantic-live-after-path-miss",
        "timeout": timeout,
    }


def set_finance_org_by_legacy_control_name(
    jab,
    value,
    scope_hwnd=None,
    accepted_text=None,
):
    accepted_text = accepted_text or FINANCE_ORG_ACCEPTED_TEXT
    context, vm_id, owned_contexts, owned_indexes, window_info = (
        _trial.find_context_with_window(
            jab,
            HEADER_SCOPE_ANCHOR_TEXT,
            roles=("text",),
            timeout=1.5,
            require_showing=True,
            window_class="SunAwtCanvas",
            visible_only=True,
            scope_hwnd=scope_hwnd,
        )
    )
    if not context:
        return {
            "ok": False,
            "method": "legacy-control-name-guarded-paste-enter",
            "reason": "control not found",
            "control_name": HEADER_SCOPE_ANCHOR_TEXT,
            "scope_hwnd": scope_hwnd,
        }
    path = "0" + "".join(f".{index}" for index in owned_indexes)
    try:
        info_before = jab.get_context_info(vm_id, context)
        before = jab.get_text_context_value(vm_id, context)
        paste_result = _trial.guarded_paste_header_value(
            jab,
            vm_id,
            context,
            window_info,
            value,
        )
        set_text_result = None
        if not paste_result.get("ok"):
            set_ok = bool(jab.set_text_context(vm_id, context, value))
            set_text_result = {"ok": set_ok, "method": "setTextContents"}
        info_after = jab.get_context_info(vm_id, context)
        after = jab.get_text_context_value(vm_id, context)
        backend_state = describe_backend_field_state(
            info_after,
            after,
            value=value,
            accepted_text=accepted_text,
        )
        acceptance_probe = confirm_finance_org_accepted(
            jab,
            vm_id,
            context,
            expected_text=accepted_text,
            value=value,
            scope_hwnd=scope_hwnd,
        )
        accepted = bool(acceptance_probe.get("accepted"))
        write_ok = bool(paste_result.get("ok") or (set_text_result or {}).get("ok"))
        return {
            "ok": bool(write_ok and accepted),
            "method": "legacy-control-name-guarded-paste-enter",
            "source": "legacy-control-name",
            "control_name": HEADER_SCOPE_ANCHOR_TEXT,
            "path": path,
            "window": window_info,
            "scope_hwnd": scope_hwnd,
            "text_before": before,
            "description_before": (
                info_before.description.strip() if info_before else None
            ),
            "set_ok": write_ok,
            "set_text_ok": bool((set_text_result or {}).get("ok")),
            "guarded_paste": paste_result,
            "set_text_fallback": set_text_result,
            "enter_ok": bool(paste_result.get("enter_ok")),
            "post_write_snapshot": backend_state,
            "acceptance_probe": acceptance_probe,
            "accepted_text": accepted_text_from_backend(
                acceptance_probe.get("matched_snapshot") or backend_state,
                value,
                accepted_text,
            ),
            "text_after": after,
            "description_after": (
                info_after.description.strip() if info_after else None
            ),
        }
    finally:
        jab.release_contexts(vm_id, owned_contexts)


def confirm_finance_org_accepted(
    jab,
    vm_id,
    context,
    expected_text=FINANCE_ORG_ACCEPTED_TEXT,
    value=None,
    scope_hwnd=None,
    timeout=2.0,
    interval=0.15,
):
    deadline = time.time() + max(float(timeout or 0), 0.0)
    attempts = []
    first = True
    while first or time.time() < deadline:
        first = False
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        snapshot = describe_backend_field_state(
            info,
            text,
            value=value,
            accepted_text=expected_text,
        )
        attempts.append(snapshot)
        if snapshot.get("accepted"):
            return {
                "ok": True,
                "accepted": True,
                "expected_text": expected_text,
                "source": "current-context",
                "attempts": attempts,
                "matched_snapshot": snapshot,
            }
        scoped_match = _trial.probe_finance_org_accepted_text_in_scope(
            jab,
            expected_text,
            scope_hwnd=scope_hwnd,
        )
        if scoped_match.get("accepted"):
            return {
                "ok": True,
                "accepted": True,
                "expected_text": expected_text,
                "source": "scope-text-probe",
                "attempts": attempts,
                "matched_snapshot": scoped_match.get("snapshot"),
                "scope_probe": scoped_match,
            }
        if time.time() >= deadline:
            break
        time.sleep(max(float(interval or 0), 0.0))
    return {
        "ok": False,
        "accepted": False,
        "expected_text": expected_text,
        "attempts": attempts,
        "scope_probe": _trial.probe_finance_org_accepted_text_in_scope(
            jab,
            expected_text,
            scope_hwnd=scope_hwnd,
        ),
        "reason": "财务组织未确认解析为中文",
    }


def probe_finance_org_accepted_text_in_scope(
    jab,
    expected_text=FINANCE_ORG_ACCEPTED_TEXT,
    scope_hwnd=None,
):
    context, vm_id, owned_contexts, owned_indexes, window_info = (
        _trial.find_context_with_window(
            jab,
            expected_text,
            roles=(),
            timeout=0.05,
            require_showing=False,
            window_class="SunAwtCanvas",
            visible_only=True,
            scope_hwnd=scope_hwnd,
        )
    )
    if not context:
        return {
            "ok": False,
            "accepted": False,
            "expected_text": expected_text,
            "reason": "scope accepted text not found",
        }
    try:
        info = jab.get_context_info(vm_id, context)
        text = jab.get_text_context_value(vm_id, context)
        snapshot = describe_backend_field_state(
            info,
            text,
            value=None,
            accepted_text=expected_text,
        )
        return {
            "ok": bool(snapshot.get("accepted")),
            "accepted": bool(snapshot.get("accepted")),
            "expected_text": expected_text,
            "path": "0" + "".join(f".{index}" for index in owned_indexes),
            "window": window_info,
            "snapshot": snapshot,
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


def find_receipt_header_field_by_semantic_label(
    jab,
    label,
    scope_hwnd=None,
    timeout=1.5,
):
    label_found = _trial.find_header_label_context_with_window(
        jab,
        label,
        timeout=timeout,
        require_showing=label != HEADER_SCOPE_ANCHOR_LABEL,
        scope_hwnd=scope_hwnd,
    )
    label_context, vm_id, owned_contexts, owned_indexes, window = label_found
    if not label_context:
        return {"ok": False, "label": label, "reason": "semantic label not found"}
    label_path = None
    if owned_indexes:
        label_path = "0" + "".join(f".{index}" for index in owned_indexes)
    jab.release_contexts(vm_id, owned_contexts)
    if not label_path:
        return {"ok": False, "label": label, "reason": "semantic label path missing"}
    text_path = infer_header_text_path_from_label_path(label, label_path)
    if not text_path:
        return {
            "ok": False,
            "label": label,
            "label_path": label_path,
            "reason": "semantic label path cannot infer text path",
        }
    label_window_hwnd = (window or {}).get("hwnd") or scope_hwnd
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        text_path,
        class_name="SunAwtCanvas",
        scope_hwnd=label_window_hwnd,
        role="text",
        require_showing=True,
        require_valid_bounds=False,
    )
    if not context:
        return {
            "ok": False,
            "label": label,
            "label_path": label_path,
            "path": text_path,
            "reason": "semantic inferred text path not found",
            "label_window": window,
            "path_scope_hwnd": label_window_hwnd,
        }
    return {
        "ok": True,
        "label": label,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": label_path,
        "window": window_info,
    }


def find_receipt_header_field_by_dynamic_path(
    jab,
    label,
    dynamic_index,
    scope_hwnd=None,
    require_showing=True,
    require_valid_bounds=True,
    path_template=None,
):
    text_path = (
        build_receipt_header_path_from_template(dynamic_index, label, path_template)
        if path_template
        else build_receipt_header_dynamic_path(dynamic_index, label)
    )
    if not text_path:
        return {
            "ok": False,
            "reason": "header dynamic path not configured",
            "label": label,
            "dynamic_index": dynamic_index,
        }
    context, vm_id, owned_contexts, window_info = jab.find_context_by_path_once(
        text_path,
        class_name="SunAwtCanvas",
        scope_hwnd=scope_hwnd,
        role="text",
        require_showing=require_showing,
        require_valid_bounds=require_valid_bounds,
    )
    if not context:
        return {
            "ok": False,
            "reason": "header dynamic path not found",
            "label": label,
            "path": text_path,
            "dynamic_index": dynamic_index,
        }
    return {
        "ok": True,
        "context": context,
        "vm_id": vm_id,
        "owned_contexts": owned_contexts,
        "path": text_path,
        "label_path": (
            build_receipt_header_label_path_from_template(
                dynamic_index,
                label,
                path_template,
            )
            if path_template
            else build_receipt_header_dynamic_label_path(dynamic_index, label)
        ),
        "window": window_info,
        "dynamic_index": dynamic_index,
        "dynamic_prefix": receipt_header_dynamic_prefix(dynamic_index),
        "path_template": path_template,
    }


def describe_backend_field_state(info, text, value=None, accepted_text=None):
    name = info.name.strip() if info else ""
    description = info.description.strip() if info else ""
    accepted = backend_field_accepts(info, text, value, accepted_text)
    written = backend_field_has_written_value(info, text, value)
    return {
        "accepted": bool(accepted),
        "written": bool(written),
        "unlocked": False,
        "text": text,
        "name": name,
        "description": description,
    }


def backend_field_has_written_value(info, text, value=None):
    expected = str(value).strip() if value is not None else ""
    if not info or not expected:
        return False
    actual_text = str(text or "").strip()
    description = info.description.strip()
    return actual_text == expected or description == expected


def backend_field_accepts(info, text, value=None, accepted_text=None):
    if not info:
        return False
    if accepted_text:
        return context_contains(info, accepted_text)
    expected = str(value).strip() if value is not None else ""
    actual_text = str(text or "").strip()
    description = info.description.strip()
    if expected and (actual_text == expected or description == expected):
        return True
    return bool(description)


def context_contains(info, expected_text):
    if not info or not expected_text:
        return False
    expected = str(expected_text).strip()
    haystack = " ".join(
        part
        for part in (
            info.name.strip(),
            info.description.strip(),
            info.role.strip(),
            info.role_en_US.strip(),
            info.states.strip(),
            info.states_en_US.strip(),
        )
        if part
    )
    return expected in haystack
