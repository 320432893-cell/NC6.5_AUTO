import ctypes
from typing import Any, cast

from core.jab_operator import JABOperator
from tools.jab_probe import AccessibleContextInfo


def make_info(states, x=1, y=1, width=10, height=10):
    info = AccessibleContextInfo()
    info.states_en_US = states
    info.x = x
    info.y = y
    info.width = width
    info.height = height
    return info


def test_context_info_is_showing_requires_showing_state():
    assert JABOperator.context_info_is_showing(make_info("enabled,visible,showing"))
    assert not JABOperator.context_info_is_showing(make_info("enabled,visible"))


def test_context_info_has_valid_bounds_rejects_mirror_bounds():
    assert JABOperator.context_info_has_valid_bounds(
        make_info("enabled", 406, 548, 70, 15)
    )
    assert not JABOperator.context_info_has_valid_bounds(
        make_info("enabled", -1, -1, -1, -1)
    )
    assert not JABOperator.context_info_has_valid_bounds(
        make_info("enabled", 0, 0, 0, 15)
    )


def test_find_context_by_path_rejects_visible_non_showing_mirror(monkeypatch):
    operator = JABOperator({"jab": {}})
    operator.dll = cast(
        Any,
        FakeDll(make_info("enabled,focusable,visible", -1, -1, -1, -1)),
    )

    monkeypatch.setattr(
        "core.jab_operator.enum_windows",
        lambda include_children=True: [(1, "", "SunAwtCanvas", 2, True)],
    )

    result = operator.find_context_by_path_once(
        "0.0",
        name="2026-06-02",
        role="push button",
        require_showing=True,
        require_valid_bounds=True,
    )

    assert result == (None, None, [], {})


class FakeDll:
    def __init__(self, info):
        self.info = info

    def isJavaWindow(self, hwnd):
        return True

    def getAccessibleContextFromHWND(self, hwnd, vm_id, root_context):
        vm_id._obj.value = 100
        root_context._obj.value = 200
        return True

    def getAccessibleChildFromContext(self, vm_id, context, index):
        return 300

    def getAccessibleContextInfo(self, vm_id, context, info_ref):
        info = ctypes.cast(info_ref, ctypes.POINTER(AccessibleContextInfo)).contents
        info.name = "2026-06-02"
        info.role_en_US = "push button"
        info.states_en_US = self.info.states_en_US
        info.x = self.info.x
        info.y = self.info.y
        info.width = self.info.width
        info.height = self.info.height
        return True
