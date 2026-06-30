import json
from dataclasses import asdict, dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from core.errors import TableMatchError
from core.models import ExcelVoucherItem, PendingMatch
from core.paths import logs_dir


PLAN_CACHE_NAME = "voucher_precheck_plan.json"


@dataclass(frozen=True)
class VoucherPlanCache:
    excel_path: str
    sheet: str
    has_header: bool
    limit: int | None
    start_row: int | None
    end_row: int | None
    rows: list[dict[str, Any]]
    created_at: str


def plan_cache_path() -> Path:
    return logs_dir() / PLAN_CACHE_NAME


def normalize_range_value(value):
    if value is None:
        return None
    return int(value)


def item_to_plan_row(match: PendingMatch) -> dict[str, Any]:
    item = match.item
    return {
        "excel_row": item.row,
        "nc_row": match.nc_row,
        "amount": str(item.amount) if isinstance(item.amount, Decimal) else item.amount,
        "partner": item.partner,
    }


def write_voucher_plan_cache(
    *,
    config: dict,
    limit,
    start_row,
    end_row,
    matches: list[PendingMatch],
) -> Path:
    cache = VoucherPlanCache(
        excel_path=str(config.get("excel_path", "")),
        sheet=str(config.get("sheet_my", "")),
        has_header=bool(config.get("has_header", True)),
        limit=normalize_range_value(limit),
        start_row=normalize_range_value(start_row),
        end_row=normalize_range_value(end_row),
        rows=[item_to_plan_row(match) for match in matches],
        created_at=datetime.now().isoformat(timespec="seconds"),
    )
    path = plan_cache_path()
    path.write_text(
        json.dumps(asdict(cache), ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return path


def load_voucher_plan_cache() -> dict:
    path = plan_cache_path()
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise TableMatchError("未找到上次预检查队列计划，请先执行预检查。") from exc
    except json.JSONDecodeError as exc:
        raise TableMatchError("上次预检查队列计划已损坏，请重新执行预检查。") from exc


def validate_voucher_plan_cache(
    *,
    cache: dict,
    config: dict,
    limit,
    start_row,
    end_row,
    pending: list[ExcelVoucherItem],
) -> None:
    expected = {
        "excel_path": str(config.get("excel_path", "")),
        "sheet": str(config.get("sheet_my", "")),
        "has_header": bool(config.get("has_header", True)),
        "limit": normalize_range_value(limit),
        "start_row": normalize_range_value(start_row),
        "end_row": normalize_range_value(end_row),
    }
    actual = {key: cache.get(key) for key in expected}
    if actual != expected:
        raise TableMatchError(
            "上次预检查队列计划与当前参数不一致，请重新执行预检查。"
            f" 当前={expected} 计划={actual}"
        )

    plan_rows = cache.get("rows") or []
    pending_rows = [item.row for item in pending]
    cached_rows = [int(row.get("excel_row")) for row in plan_rows]
    if cached_rows != pending_rows:
        raise TableMatchError(
            "上次预检查队列计划与当前 Excel 队列不一致，请重新执行预检查。"
            f" 当前Excel行={pending_rows} 计划Excel行={cached_rows}"
        )


def matches_from_plan_cache(
    cache: dict, pending: list[ExcelVoucherItem]
) -> list[PendingMatch]:
    by_row = {item.row: item for item in pending}
    matches = []
    for row in cache.get("rows") or []:
        excel_row = int(row.get("excel_row"))
        item = by_row.get(excel_row)
        if item is None:
            raise TableMatchError(f"预检查队列计划包含当前队列外 Excel 行: {excel_row}")
        matches.append(
            PendingMatch(
                item=item,
                nc_row=int(row.get("nc_row")),
                row_data={
                    "row_index": int(row.get("nc_row")),
                    "amount": item.amount,
                    "partner": item.partner,
                },
            )
        )
    return matches
