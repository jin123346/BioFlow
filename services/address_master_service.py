import difflib
import json
import logging

import pandas as pd

from config.paths import CACHE_DIR, DATA_DIR
from utils.text_utils import normalize_text

ADDRESS_FILE = DATA_DIR / "address" / "영문주소_정리.xlsx"
ADDRESS_CACHE_FILE = CACHE_DIR / "address_lookup.json"

# level -> [(시트명, 국문 컬럼, 영문 컬럼), ...] — 여러 시트를 한 레벨로 풀링
LEVEL_SHEETS: dict[str, list[tuple[str, str, str]]] = {
    "sido": [("시도", "국문명", "영문명")],
    "sigungu": [("시군구", "국문명", "영문명")],
    "eupmyeondong": [("읍면동", "국문명", "영문명")],
    # "상세주소" 자동완성/자동매칭용: 리·도로명·자연지명·과거 상세위치 예시를 하나로 풀링
    "detail": [
        ("리", "국문명", "영문명"),
        ("도로명", "국문명", "영문명"),
        ("자연지명_기타", "국문명", "영문명"),
        ("상세위치_참고", "국문(상세위치 예시)", "영문(상세위치 예시)"),
    ],
}


class AddressMasterService:
    def __init__(self):
        self._lookup: dict[str, dict[str, str]] = {}
        self._names: dict[str, list[str]] = {}
        self._loaded = False

    def _ensure_loaded(self):
        if self._loaded:
            return
        if not ADDRESS_FILE.exists():
            self._loaded = True
            return
        stat = ADDRESS_FILE.stat()
        fingerprint = [stat.st_mtime_ns, stat.st_size]
        try:
            cached = json.loads(ADDRESS_CACHE_FILE.read_text(encoding="utf-8"))
            lookup = cached["lookup"]
            if cached.get("version") == 1 and cached.get("source") == fingerprint and all(
                isinstance(lookup.get(level), dict) and all(isinstance(k, str) and isinstance(v, str) for k, v in lookup[level].items())
                for level in LEVEL_SHEETS
            ):
                self._lookup = lookup
                self._names = {level: list(values) for level, values in lookup.items()}
                self._loaded = True
                return
        except (OSError, ValueError, KeyError, TypeError, AttributeError):
            pass
        try:
            excel_file = pd.ExcelFile(ADDRESS_FILE)
        except Exception:
            return

        for level, sheets in LEVEL_SHEETS.items():
            lookup: dict[str, str] = {}
            for sheet_name, kr_col, en_col in sheets:
                if sheet_name not in excel_file.sheet_names:
                    continue
                df = excel_file.parse(sheet_name)
                for row in df.to_dict("records"):
                    korean = normalize_text(row.get(kr_col))
                    english = normalize_text(row.get(en_col))
                    if korean and english and korean not in lookup:
                        lookup[korean] = english
            self._lookup[level] = lookup
            self._names[level] = list(lookup.keys())
        excel_file.close()
        self._loaded = True
        try:
            ADDRESS_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
            temporary = ADDRESS_CACHE_FILE.with_suffix(".tmp")
            temporary.write_text(json.dumps({"version": 1, "source": fingerprint, "lookup": self._lookup}, ensure_ascii=False), encoding="utf-8")
            temporary.replace(ADDRESS_CACHE_FILE)
        except OSError:
            logging.getLogger(__name__).warning("Could not cache address lookup", exc_info=True)

    def candidates(self, level: str) -> list[str]:
        """자동완성용 해당 레벨의 국문 후보 전체 목록."""
        self._ensure_loaded()
        return self._names.get(level, [])

    def match_english(self, level: str, korean_text: str) -> tuple[str, bool]:
        """국문 지명을 영문으로 매칭. (영문명, 완전일치여부)를 반환. 없으면 ("", False)."""
        self._ensure_loaded()
        korean_text = normalize_text(korean_text)
        lookup = self._lookup.get(level, {})
        if not korean_text or not lookup:
            return "", False

        if korean_text in lookup:
            return lookup[korean_text], True

        # ponytail: difflib 유사도 매칭, 계층(시도->시군구) 필터 없이 전체에서 최근접 1건만 사용
        candidates = difflib.get_close_matches(korean_text, self._names.get(level, []), n=1, cutoff=0.6)
        if candidates:
            return lookup[candidates[0]], False
        return "", False
