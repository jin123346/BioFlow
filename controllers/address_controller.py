"""Shared observation/specimen VWorld workflow; no credential diagnostics."""

import re

from PySide6.QtWidgets import QDialog, QMessageBox

from services.api_key_service import ApiKeyError, ApiKeyService
from services.vworld_geocoding_service import VWorldAddress, VWorldGeocodingService
from services.address_mapping_service import METADATA_KEY, exact_master_english, valid_metadata
from services.english_address_service import AddressDatabaseError, EnglishAddress, EnglishAddressService
from ui.dialogs.api_settings_dialog import ApiSettingsDialog
from ui.dialogs.address_lookup_dialog import AddressLookupDialog, AddressPreviewDialog

_SESSION_ADDRESS_CACHE = {}


def registered_english_address(record, specimen=False):
    """Build a local-only English reference from exact user dictionary matches."""
    prefix = "채집위치" if specimen else "관찰위치"
    korean_sido = str(record.get(f"{prefix} 시도 국문", "") or "").strip()
    korean_sigungu = str(record.get(f"{prefix} 시군구 국문", "") or "").strip()
    korean_town = str(record.get(f"{prefix} 읍면동 국문", "") or "").strip()
    korean_detail = str(record.get(f"{prefix} 상세 국문", "") or "").strip()
    sido = exact_master_english("sido", korean_sido)
    town = exact_master_english("eupmyeondong", korean_town) if korean_town else ""
    if not sido:
        return None

    city = district = ""
    whole_district = exact_master_english("sigungu", korean_sigungu)
    if not korean_sigungu:
        pass
    elif whole_district:
        district = whole_district
    else:
        translated = [(part, exact_master_english("sigungu", part)) for part in korean_sigungu.split()]
        if translated and all(english for _part, english in translated):
            for korean, english in translated:
                if korean.endswith("시") and len(translated) > 1:
                    city = english
                else:
                    district = english

    ri = ""
    for token in korean_detail.split():
        if token.endswith("리"):
            ri = exact_master_english("detail", token)
            break
    parcel_match = re.search(r"(?:^|\s)(산\s*)?(\d+(?:-\d+)?)(?:\s|$)", korean_detail)
    parcel = f"{'산 ' if parcel_match and parcel_match.group(1) else ''}{parcel_match.group(2)}" if parcel_match else ""
    structure = {
        "level0": str(record.get("국가국문명", "") or "대한민국"),
        "level1": korean_sido, "level2": korean_sigungu, "level3": "",
        "level4L": korean_town, "level4LC": "", "level5": parcel, "detail": "",
    }
    address = VWorldAddress(
        structure, " ".join(x for x in (korean_sido, korean_sigungu, korean_town, korean_detail) if x),
        str(record.get("위도", "") or ""), str(record.get("경도", "") or ""),
    )
    return address, EnglishAddress(sido, city, district, town, ri)


def lookup_addresses(parent, targets, specimen=False, differences_only=False, mode="all"):
    if not targets:
        return {}
    results, errors, references, reference_errors = {}, {}, {}, {}
    pending = []
    english_service = EnglishAddressService() if mode == "english" else None
    for identifier, record in targets:
        metadata = valid_metadata(record, record.get(METADATA_KEY)) if mode == "english" else None
        if metadata:
            address = VWorldAddress(
                dict(metadata.get("structure", {})), metadata.get("text", ""),
                str(metadata.get("latitude", "")), str(metadata.get("longitude", "")),
            )
            results[identifier] = address
            try:
                reference = english_service.lookup(metadata.get("legal_dong_code", ""))
                if reference:
                    references[identifier] = reference
                else:
                    registered = registered_english_address(record, specimen)
                    if registered:
                        results[identifier], references[identifier] = registered
                    else:
                        reference_errors[identifier] = "법정동코드가 영문 DB에 없습니다. 기존 영문주소를 유지합니다."
            except AddressDatabaseError as exc:
                reference_errors[identifier] = str(exc)
        else:
            registered = registered_english_address(record, specimen) if mode == "english" else None
            if registered:
                results[identifier], references[identifier] = registered
            else:
                pending.append((identifier, record))

    if pending:
        keys = ApiKeyService()
        try:
            key, expiry = keys.load()
            if not key:
                ApiSettingsDialog(parent, key_service=keys).exec()
                key, expiry = keys.load()
            if not key:
                return {}
        except ApiKeyError as exc:
            QMessageBox.warning(parent, "VWorld API 키", str(exc))
            return {}
        geocoder = VWorldGeocodingService(cache=_SESSION_ADDRESS_CACHE)
        dialog = AddressLookupDialog(geocoder, key, pending, parent, load_english=mode != "korean")
        key = ""
        if dialog.exec() != QDialog.Accepted:
            return {}
        results.update(dialog.worker.results)
        errors.update(dialog.worker.errors)
        references.update(dialog.worker.references)
        reference_errors.update(dialog.worker.reference_errors)

    preview = AddressPreviewDialog(targets, results, errors, specimen, parent,
                                   references, reference_errors, differences_only, mode)
    if preview.exec() != QDialog.Accepted:
        return {}
    return preview.planned_records()
