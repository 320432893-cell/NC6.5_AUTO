# 职责：收款单表头 path/label-path 模板的纯计算——动态前缀、模板套用、索引提取、label 匹配
# 不做什么：不碰 JAB/窗口/键盘，不做字段写入或 scope 解析,无副作用(模板缓存仅函数属性)
# 允许依赖层：仅标准库;本模块是表头工具链的叶子,不依赖其它 receipt_header_* 模块
# 谁不应该 import：无限制(纯函数叶子);但本模块不应反向 import scope/writer/tree/trial

CURRENCY_NAMES = {"USD": "美元", "RMB": "人民币", "CNY": "人民币"}
HEADER_DYNAMIC_PREFIX_BASE = "0.0.1.0.0.0.0"
HEADER_COMMON_SUFFIX_TEMPLATE = "0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.{index}.0"
HEADER_COMMON_LABEL_SUFFIX_TEMPLATE = "0.0.0.1.1.0.0.0.0.1.0.2.0.0.0.0.0.0.0.{index}"
FINANCE_ORG_LABEL_SUFFIX = "0.0.0.1.1.0.0.0.1.1.1.0"
HEADER_LIVE_SEMANTIC_FALLBACK_TIMEOUT = 0.35
HEADER_FORM_TEXT_INDEXES = {
    "单据日期": 5,
    "币种": 13,
    "收款银行账户": 15,
    "客户": 17,
    "结算方式": 31,
}
HEADER_REQUIRED_LABELS = ("财务组织", "客户", "单据日期", "币种", "结算方式")
HEADER_PROBE_LABEL_KEYS = {
    "finance": "财务组织",
    "finance_org": "财务组织",
    "customer": "客户",
    "client": "客户",
    "date": "单据日期",
    "document_date": "单据日期",
    "currency": "币种",
    "settlement": "结算方式",
    "settlement_method": "结算方式",
}
HEADER_SCOPE_ANCHOR_LABEL = "财务组织"
HEADER_SCOPE_ANCHOR_TEXT = "财务组织(O)"
FINANCE_ORG_ACCEPTED_TEXT = "上海移为通信技术股份有限公司"


def build_receipt_header_dynamic_path(dynamic_index, label):
    cached_template = get_receipt_header_path_template(dynamic_index)
    if cached_template:
        path = build_receipt_header_path_from_template(
            dynamic_index,
            label,
            cached_template,
        )
        if path:
            return path
    if label == "财务组织":
        label_path = build_receipt_header_dynamic_label_path(dynamic_index, label)
        return infer_header_text_path_from_label_path(label, label_path)
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    if index is None:
        return None
    suffix = HEADER_COMMON_SUFFIX_TEMPLATE.format(index=index)
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{suffix}"


def build_receipt_header_dynamic_label_path(dynamic_index, label):
    cached_template = get_receipt_header_path_template(dynamic_index)
    if cached_template:
        path = build_receipt_header_label_path_from_template(
            dynamic_index,
            label,
            cached_template,
        )
        if path:
            return path
    if label == "财务组织":
        return (
            f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{FINANCE_ORG_LABEL_SUFFIX}"
        )
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    if index is None:
        return None
    suffix = HEADER_COMMON_LABEL_SUFFIX_TEMPLATE.format(index=index - 1)
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{suffix}"


def receipt_header_dynamic_prefix(dynamic_index):
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}"


def receipt_header_default_path_template():
    return {
        "source": "default-header-template",
        "text_suffix_template": HEADER_COMMON_SUFFIX_TEMPLATE,
        "label_suffix_template": HEADER_COMMON_LABEL_SUFFIX_TEMPLATE,
    }


def set_receipt_header_path_template(dynamic_index, template):
    if dynamic_index is None or not template:
        return None
    cache = getattr(set_receipt_header_path_template, "_cache", None)
    if cache is None:
        cache = {}
        setattr(set_receipt_header_path_template, "_cache", cache)
    cache[int(dynamic_index)] = dict(template)
    return cache[int(dynamic_index)]


def get_receipt_header_path_template(dynamic_index):
    if dynamic_index is None:
        return None
    cache = getattr(set_receipt_header_path_template, "_cache", None) or {}
    return cache.get(int(dynamic_index))


def clear_receipt_header_path_template_cache():
    setattr(set_receipt_header_path_template, "_cache", {})


def build_receipt_header_path_from_template(dynamic_index, label, template):
    if label == "财务组织":
        return None
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    suffix_template = (template or {}).get("text_suffix_template")
    if dynamic_index is None or index is None or not suffix_template:
        return None
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{suffix_template.format(index=index)}"


def build_receipt_header_label_path_from_template(dynamic_index, label, template):
    if label == "财务组织":
        return None
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    suffix_template = (template or {}).get("label_suffix_template")
    if dynamic_index is None or index is None or not suffix_template:
        return None
    return f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}.{suffix_template.format(index=index - 1)}"


def infer_header_path_template_from_field(path, dynamic_index, label):
    index = HEADER_FORM_TEXT_INDEXES.get(label)
    if not path or dynamic_index is None or index is None:
        return None
    prefix = f"{HEADER_DYNAMIC_PREFIX_BASE}.{dynamic_index}."
    suffix = str(path)
    if not suffix.startswith(prefix):
        return None
    suffix = suffix[len(prefix) :]
    marker = f".{index}.0"
    if not suffix.endswith(marker):
        return None
    base = suffix[: -len(marker)]
    if not base:
        return None
    return {
        "source": f"learned-from-{label}",
        "text_suffix_template": f"{base}.{{index}}.0",
        "label_suffix_template": f"{base}.{{index}}",
        "sample_label": label,
        "sample_path": path,
    }


def extract_receipt_header_dynamic_index(path):
    prefix = f"{HEADER_DYNAMIC_PREFIX_BASE}."
    if not path or not path.startswith(prefix):
        return None
    first = path[len(prefix) :].split(".", 1)[0]
    try:
        return int(first)
    except ValueError:
        return None


def infer_header_text_path_from_label_path(label, label_path):
    parts = split_header_path(label_path)
    if not parts:
        return None
    if label == "财务组织":
        if parts[-1] == 0:
            return ".".join(str(part) for part in [*parts[:-1], 2, 1, 0])
        return None
    if parts[-1] % 2 != 0:
        return None
    return ".".join(str(part) for part in [*parts[:-1], parts[-1] + 1, 0])


def split_header_path(path):
    try:
        return [int(part) for part in str(path).split(".") if part != ""]
    except ValueError:
        return []


def accepted_text_from_backend(backend_state, raw_value=None, preferred=None):
    if preferred:
        return str(preferred).strip()
    for key in ("description", "text", "name"):
        text = str((backend_state or {}).get(key) or "").strip()
        if text and text != str(raw_value or "").strip():
            return text
    return ""


def header_label_text_matches(info, label):
    expected = str(label or "").strip()
    if not expected:
        return False
    texts = (
        info.name.strip(),
        info.description.strip(),
    )
    for text in texts:
        if not text:
            continue
        normalized = text.replace("（", "(").replace("）", ")")
        if normalized == expected or normalized.startswith(f"{expected}("):
            return True
    return False


def header_scope_anchor_text_matches(info):
    if not info:
        return False
    for text in (info.name.strip(), info.description.strip()):
        if text.strip() == HEADER_SCOPE_ANCHOR_TEXT:
            return True
    return False
