"""Indexed, read-only access to the external administrative reference database."""

from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
import re
import sqlite3


class AddressDatabaseError(Exception):
    pass


def database_path():
    from config.paths import BASE_DIR, EXECUTABLE_DIR, FROZEN
    return (EXECUTABLE_DIR if FROZEN else BASE_DIR) / "data" / "address" / "english_admin.sqlite"


@dataclass(frozen=True)
class EnglishAddress:
    sido: str
    city: str
    district: str
    town: str
    ri: str


class EnglishAddressService:
    def __init__(self, path=None):
        self.path = Path(path) if path is not None else database_path()

    def connect(self):
        if not self.path.is_file():
            raise AddressDatabaseError("영문주소 기준 DB가 없습니다. 배포 폴더의 data/address/english_admin.sqlite 파일을 확인해 주세요.")
        db = None
        try:
            db = sqlite3.connect(self.path.resolve().as_uri() + "?mode=ro", uri=True)
            db.execute("PRAGMA query_only=ON")
            db.execute("PRAGMA cache_size=-256")
            if db.execute("PRAGMA user_version").fetchone()[0] != 1:
                db.close()
                raise AddressDatabaseError("지원하지 않는 영문주소 DB 버전입니다.")
            return db
        except sqlite3.Error:
            if db is not None:
                db.close()
            raise AddressDatabaseError("영문주소 기준 DB를 열 수 없습니다. 배포 파일을 확인해 주세요.") from None

    def validate(self):
        try:
            with closing(self.connect()) as db:
                db.execute("SELECT code, sido, city, district, town, ri FROM admin_names LIMIT 0")
        except sqlite3.Error:
            raise AddressDatabaseError("영문주소 기준 DB 구조가 올바르지 않습니다.") from None

    def lookup(self, code):
        if not isinstance(code, str) or not re.fullmatch(r"[0-9]{10}", code):
            return None
        try:
            with closing(self.connect()) as db:
                rows = db.execute("SELECT sido,city,district,town,ri FROM admin_names WHERE code=? LIMIT 2", (code,)).fetchall()
                if len(rows) > 1:
                    raise AddressDatabaseError("법정동코드에 서로 다른 영문명이 있어 자동 적용하지 않습니다. 기존 영문주소를 유지합니다.")
                return EnglishAddress(*rows[0]) if rows else None
        except sqlite3.Error:
            raise AddressDatabaseError("영문주소 기준 DB 조회에 실패했습니다.") from None


def english_fields(address, reference, specimen=False):
    """Only structural names and provider parcel numbers, never free-text translation."""
    prefix = "채집위치" if specimen else "관찰위치"
    district = ", ".join(x for x in (reference.district, reference.city) if x)
    from services.address_mapping_service import address_detail_parts
    _, parcel = address_detail_parts(address)
    if parcel.startswith("산 "):
        parcel = "San " + parcel[2:]
    elif parcel and not re.fullmatch(r"\d+(?:-\d+)?", parcel):
        parcel = ""
    return {
        f"{prefix} 시도 영문": reference.sido,
        f"{prefix} 시군구 영문": district,
        f"{prefix} 읍면동 영문": reference.town,
        f"{prefix} 상세 영문": ", ".join(x for x in (parcel, reference.ri) if x),
    }
