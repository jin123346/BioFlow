"""Map structured VWorld results without altering provider administrative regions."""

from copy import deepcopy
from functools import lru_cache
import re

from services.coordinate_service import parse_coordinates

METADATA_KEY = "_vworld_address"


@lru_cache(maxsize=4096)
def exact_master_english(level, korean_text):
    """Use the standardized flat dictionary only for exact Korean matches."""
    from services.address_master_service import AddressMasterService
    if not hasattr(exact_master_english, "service"):
        exact_master_english.service = AddressMasterService()
    english, exact = exact_master_english.service.match_english(level, korean_text)
    return english if exact else ""


def address_detail_parts(address):
    """Return a Korean ri and a normalized parcel without cadastral land type."""
    s = address.structure
    parcel = s.get("level5", "").strip() or s.get("detail", "").strip()
    match = re.fullmatch(r"(산)?\s*(\d+(?:-\d+)?)(?:\s*[가-힣])?", parcel)
    if match:
        parcel = f"{'산 ' if match.group(1) else ''}{match.group(2)}"
    ri = ""
    town = s.get("level4L", "").strip()
    # VWorld does not expose Korean ri as a separate structure field. Its text
    # does preserve it between level4L and level5 for rural parcel addresses.
    tokens = address.text.split()
    if town in tokens:
        start = tokens.index(town) + 1
        for token in tokens[start:]:
            if token.endswith("리"):
                ri = token
                break
    return ri, parcel


def map_address(address, specimen=False):
    s = address.structure
    level2, level3 = s.get("level2", ""), s.get("level3", "")
    combined = re.fullmatch(r"(.+시)\s+([^\s]+구)", level2)
    if combined and (not level3 or level3 == combined.group(2)):
        level2, level3 = combined.groups()
    city = level2 if level3 and level2.endswith("시") else ""
    district = level3 if city else " ".join(x for x in (level2, level3) if x)
    ri, parcel = address_detail_parts(address)
    prefix = "채집위치" if specimen else "관찰위치"
    fields = {
        "국가국문명": s.get("level0", ""),
        f"{prefix} 시도 국문": s.get("level1", ""),
        f"{prefix} 시군구 국문": " ".join(x for x in (city, district) if x),
        f"{prefix} 읍면동 국문": s.get("level4L", ""),
        f"{prefix} 상세 국문": " ".join(x for x in (ri, parcel) if x),
    }
    metadata = {
        "provider": "VWorld", "latitude": address.latitude, "longitude": address.longitude,
        "text": address.text,
        "structure": dict(s), "city": city, "district": district,
        "legal_dong_code": s.get("level4LC", ""), "detail_aux": s.get("detail", ""),
        "ri_korean": ri, "parcel": parcel,
    }
    return fields, metadata


def apply_address(record, address, specimen=False, overwrite=False, reference=None, mode="all"):
    if mode not in {"all", "korean", "english"}:
        raise ValueError(f"Unsupported address mode: {mode}")
    fields, metadata = map_address(address, specimen)
    updated = deepcopy(record)
    if mode in {"all", "korean"}:
        for field, value in fields.items():
            if value and (overwrite or not str(record.get(field, "") or "").strip()):
                updated[field] = value
    metadata["applied_fields"] = {field: str(updated.get(field, "") or "") for field in fields}
    metadata["matches_displayed_address"] = all(
        not value or metadata["applied_fields"][field] == value for field, value in fields.items()
    )
    metadata["write_fields"] = list(fields) if mode in {"all", "korean"} else []
    if mode in {"all", "english"} and reference is not None:
        from services.english_address_service import english_fields
        prefix = "채집위치" if specimen else "관찰위치"
        aligned = all(
            updated.get(f"{prefix} {part} 국문", "") == fields[f"{prefix} {part} 국문"]
            for part in ("시도", "시군구", "읍면동")
        )
        translations = english_fields(address, reference, specimen)
        levels = {"시도": "sido", "시군구": "sigungu", "읍면동": "eupmyeondong", "상세": "detail"}
        for field, reference_value in translations.items():
            if mode == "all" and not aligned and field != "국가영문명":
                continue
            korean_field = field.replace("영문", "국문")
            part = next((name for name in levels if f" {name} " in field), None)
            standard_value = exact_master_english(levels[part], updated.get(korean_field, "")) if part else ""
            # Exact standard-name matches are safe even when VWorld uses a newer
            # province label. Otherwise retain the strict provider alignment.
            if standard_value:
                value = standard_value
            elif mode == "english":
                # English-only lookup is an explicit, preview-confirmed action.
                # Apply the retrieved structural/detail value even when a
                # free-text Korean detail does not equal the provider string.
                value = reference_value
            elif updated.get(korean_field, "") == fields.get(korean_field, ""):
                value = reference_value
            else:
                continue
            if value and (overwrite or not str(updated.get(field, "") or "").strip()):
                updated[field] = value
                metadata["write_fields"].append(field)
    updated[METADATA_KEY] = metadata
    return updated


def valid_metadata(record, metadata):
    """Discard stale codes after a manual coordinate/address edit."""
    if not isinstance(metadata, dict) or metadata.get("provider") != "VWorld":
        return None
    try:
        current = parse_coordinates(record.get("위도"), record.get("경도"))
        source = parse_coordinates(metadata.get("latitude"), metadata.get("longitude"))
    except ValueError:
        return None
    fields = metadata.get("applied_fields")
    if current != source or not isinstance(fields, dict):
        return None
    if any(str(record.get(k, "") or "") != v for k, v in fields.items()):
        return None
    return deepcopy(metadata)
