"""Database identifier generation and validation rules.

This module intentionally contains no Qt types so the identifier contract can be
tested independently from the editor widgets.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Mapping


MEDIA_TYPE_CODES = {
    "정보": "OB",
    "이미지": "OI",
    "동영상": "OM",
    "사운드": "OS",
}


@dataclass(frozen=True)
class IdentifierResult:
    expected: str | None
    pattern: re.Pattern[str] | None
    error: str = ""


def media_code(media_type: object, stored_code: object = "") -> str:
    """Resolve a media code; missing type/code is the observation default OB."""
    media_type_text = str(media_type or "").strip()
    stored_code_text = str(stored_code or "").strip()
    if not media_type_text and not stored_code_text:
        return "OB"
    return MEDIA_TYPE_CODES.get(media_type_text, stored_code_text)


def build_identifier(
    organization: object,
    taxon: object,
    year: object,
    order: object,
    project_sequence: object,
    *,
    media: object = "",
    specimen: bool = False,
) -> str:
    values = [str(value or "").strip() for value in (organization, taxon, year, order, project_sequence)]
    organization_text, taxon_text, year_text, order_text, sequence_text = values
    if not all(values) or (not specimen and not str(media or "").strip()):
        return ""
    try:
        year_code = int(year_text) % 100
        order_code = int(order_text)
        sequence_code = int(sequence_text)
    except (TypeError, ValueError):
        return ""
    prefix = f"{organization_text}-{taxon_text}" if specimen else f"{organization_text}-{str(media).strip()}-{taxon_text}"
    return f"{prefix}-{year_code:02d}{order_code:02d}{sequence_code:05d}"


def expected_identifier(record: Mapping[str, object], *, specimen: bool = False) -> IdentifierResult:
    organization = str(record.get("기관코드", "") or "").strip()
    taxon = str(record.get("분류코드", "") or "").strip()
    year = str(record.get("사업년도", "") or "").strip()
    order = str(record.get("순번", "") or "").strip()
    sequence = str(record.get("사업별 순번", "") or "").strip()
    resolved_media = "" if specimen else media_code(record.get("미디어타입"), record.get("미디어코드"))

    required = [
        ("기관코드", organization),
        ("분류코드", taxon),
        *((() if specimen else (("미디어타입", resolved_media),))),
        ("사업년도", year),
        ("순번", order),
        ("사업별 순번", sequence),
    ]
    missing = [label for label, value in required if not value]
    if missing:
        return IdentifierResult(None, None, f"필수값 누락: {', '.join(missing)}")

    try:
        year_code = int(year) % 100
        order_code = int(order)
        sequence_code = int(sequence)
        if order_code < 0 or sequence_code < 1:
            raise ValueError
    except (TypeError, ValueError):
        return IdentifierResult(None, None, "사업년도·순번·사업별 순번이 올바른 숫자가 아님")

    code_prefix = f"{organization}-{taxon}" if specimen else f"{organization}-{resolved_media}-{taxon}"
    prefix = f"{code_prefix}-{year_code:02d}{order_code:02d}"
    return IdentifierResult(
        f"{prefix}{sequence_code:05d}",
        re.compile(rf"{re.escape(prefix)}\d{{5}}"),
    )
