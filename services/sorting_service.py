"""Stable record sorting without changing identifiers or project sequences."""

from __future__ import annotations

from typing import Any, Iterable


SORTABLE_FIELDS = {"국명", "학명", "미디어타입"}


def sort_records(records: Iterable[dict[str, Any]], field: str, descending: bool = False):
    if field not in SORTABLE_FIELDS:
        raise ValueError(f"지원하지 않는 정렬 기준입니다: {field}")

    def key(pair):
        _, record = pair
        text = str(record.get(field, "") or "").strip()
        return text.casefold()

    populated, empty = [], []
    for pair in enumerate(records):
        text = str(pair[1].get(field, "") or "").strip()
        (populated if text else empty).append(pair)
    populated.sort(key=key, reverse=descending)
    return [record for _, record in populated + empty]


def renumber_display_rows(records: Iterable[dict[str, Any]], start: int = 1):
    result = []
    for number, record in enumerate(records, start):
        updated = dict(record)
        updated["NO."] = number
        result.append(updated)
    return result
