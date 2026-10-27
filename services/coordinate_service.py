"""Coordinate validation and bulk application rules."""

from __future__ import annotations

from typing import Any, Iterable


def parse_coordinates(latitude: object, longitude: object) -> tuple[str, str]:
    try:
        lat = float(str(latitude).strip())
        lon = float(str(longitude).strip())
    except (TypeError, ValueError):
        raise ValueError("위도와 경도는 숫자로 입력해 주세요.") from None
    if not -90 <= lat <= 90:
        raise ValueError("위도는 -90~90 범위여야 합니다.")
    if not -180 <= lon <= 180:
        raise ValueError("경도는 -180~180 범위여야 합니다.")
    return format(lat, ".15g"), format(lon, ".15g")


def apply_coordinates(records: Iterable[dict[str, Any]], rows: Iterable[int], latitude: str,
                      longitude: str, *, only_empty: bool = True):
    result = [dict(record) for record in records]
    changed = []
    for row in sorted(set(rows)):
        if not 0 <= row < len(result):
            continue
        record = result[row]
        if only_empty and (str(record.get("위도", "")).strip() or str(record.get("경도", "")).strip()):
            continue
        record["위도"], record["경도"] = latitude, longitude
        changed.append(row)
    return result, changed
