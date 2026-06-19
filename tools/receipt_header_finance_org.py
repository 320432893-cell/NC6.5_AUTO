# 职责：收款单表头"财务组织"专用写法——按 legacy 控件名守护粘贴写入、确认中文解析被接受、按 scope 探针回读接受文本
# 不做什么：不做通用动态字段提交(委托 writer),不算 path 模板(委托 paths),不做纯状态判定(委托 state),不读 Excel/不解析 CLI
# 允许依赖层：tools.receipt_header_paths(纯计算/常量) + tools.receipt_header_state(纯状态判定);
#             被 monkeypatch 的协作者(find_context_with_window / guarded_paste_header_value /
#             probe_finance_org_accepted_text_in_scope)经 _trial 代理读 tools.receipt_self_made_flow
# 谁不应该 import：receipt_header_{paths,state} 不应 import 本模块(会成环);writer 经门面 re-export 本模块,不反向被本模块 import

import sys

from tools.receipt_header_paths import (
    FINANCE_ORG_ACCEPTED_TEXT,
    HEADER_SCOPE_ANCHOR_TEXT,
    accepted_text_from_backend,
)
from tools.receipt_header_state import describe_backend_field_state

import time  # noqa: F401  # 供搬移函数内 time.* 使用;测试经 _trial.time.sleep monkeypatch(共享内建 time 单例)


class _TrialNamespace:
    # 按调用时从已加载的入口模块 tools.receipt_self_made_flow 取属性：
    # 让测试对 trial 上 find_context_with_window / guarded_paste_header_value /
    # probe_finance_org_accepted_text_in_scope 的 monkeypatch 与拆分前一致地生效；
    # 不在加载期 import 入口模块以避免成环。
    def __getattr__(self, name):
        return getattr(sys.modules["tools.receipt_self_made_flow"], name)


_trial = _TrialNamespace()


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
