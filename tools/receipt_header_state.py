# 职责：收款单表头后端字段的纯状态判定——快照已写入/已被接受/上下文文本包含,无任何 JAB I/O
# 不做什么：不调 JAB、不读剪贴板/热键、不解析 path、不做探针(只对已取出的 info/text 做纯逻辑判定)
# 允许依赖层：无(纯 Python,标准库都不需要);不经 _trial 代理(本模块函数不被 monkeypatch)
# 谁不应该 import：本模块不 import 任何 tools.receipt_header_* 模块(保持叶子,任何上层都可安全依赖它)


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
