# 职责：提供 JAB 操作中可复用的金额、文本、路径和控件信息纯判定函数
# 不做什么：不加载 Java Access Bridge，不枚举窗口，不读写控件，不持有 JAB context
# 允许依赖层：标准库和调用方传入的 JAB info 结构对象
# 谁不应该 import：T0 探针脚本不应为临时试验反向依赖本模块

from decimal import Decimal, InvalidOperation


def normalize_amount(value):
    if value is None:
        return None
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def normalize_text(value):
    if value is None:
        return ""
    return "".join(str(value).split())


def text_matches(value, target, match_mode):
    text = normalize_text(value)
    if match_mode == "contains":
        return target in text
    return text == target


def parse_context_path(path):
    try:
        parts = [int(part) for part in str(path).split(".") if part != ""]
    except ValueError as exc:
        raise ValueError(f"Invalid JAB context path: {path!r}") from exc
    if not parts or parts[0] != 0:
        raise ValueError(f"JAB context path must start with 0: {path!r}")
    return parts


def context_info_is_showing(info):
    states = (info.states_en_US.strip() or info.states.strip()).lower()
    return "visible" in states and "showing" in states


def context_info_has_valid_bounds(info):
    return info.x >= 0 and info.y >= 0 and info.width > 0 and info.height > 0
