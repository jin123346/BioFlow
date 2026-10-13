"""Versioned classification reference and exact matching through family."""

import hashlib
import json
from pathlib import Path

import pandas as pd

from config.paths import DATA_DIR

CLASSIFICATION_DIR = DATA_DIR / "classification"
CLASSIFICATION_FILE = CLASSIFICATION_DIR / "master.json"
CLASSIFICATION_VERSION_FILE = CLASSIFICATION_DIR / "version.json"
REQUIRED_COLUMNS = ["clssc_sstem_info_id", "taxon_phlm_sn", "taxon_kingdom_sn", "taxonomy_name", "taxonomy_latin"]
RANK_FIELDS = [("계 국문", "계 영명"), ("문 국문", "문 영명"), ("강 국문", "강 영명"),
               ("목 국문", "목 영명"), ("과 국문", "과 영명")]


class ClassificationService:
    def __init__(self):
        self._fingerprint = None
        self._records = []

    @staticmethod
    def key(value):
        return " ".join(str(value or "").split()).casefold()

    def match(self, values):
        if not any(self.key(values.get(field)) for field in RANK_FIELDS[-1]):
            return None, "과 정보 없음"
        try:
            stat = CLASSIFICATION_FILE.stat()
            fingerprint = (stat.st_mtime_ns, stat.st_size)
            if fingerprint != self._fingerprint:
                records = validate_records(json.loads(CLASSIFICATION_FILE.read_text(encoding="utf-8")))
                self._records = [(record,
                                  [self.key(v) for v in record["taxonomy_name"].split(">")],
                                  [self.key(v) for v in record["taxonomy_latin"].split(">")]) for record in records]
                self._fingerprint = fingerprint
        except (OSError, ValueError, TypeError, KeyError):
            return None, "분류군 기준자료 확인 필요"
        matches = []
        for record, korean, latin in self._records:
            if len(korean) != 5 or len(latin) != 5:
                continue
            if all((not self.key(values.get(kr)) or self.key(values[kr]) == korean[index]) and
                   (not self.key(values.get(en)) or self.key(values[en]) == latin[index])
                   for index, (kr, en) in enumerate(RANK_FIELDS)):
                matches.append(record)
        if len(matches) == 1:
            return matches[0], "과·상위 분류 일치"
        return None, "후보가 여러 개: 상위 분류 확인" if matches else "일치하는 분류군 없음"


def validate_records(records):
    if not isinstance(records, list) or not records:
        raise ValueError("분류군 기준자료가 비어 있습니다.")
    ids = set()
    for record in records:
        if not isinstance(record, dict) or any(column not in record for column in REQUIRED_COLUMNS):
            raise ValueError("분류군 기준자료의 필수 컬럼이 없습니다.")
        identifier = str(record["clssc_sstem_info_id"])
        if not identifier.isdigit() or int(identifier) <= 0 or identifier in ids:
            raise ValueError("분류군 ID가 비어 있거나 중복되어 있습니다.")
        if not record["taxonomy_name"] or not record["taxonomy_latin"]:
            raise ValueError("분류군 이름이 비어 있습니다.")
        ids.add(identifier)
    return records


def read_version():
    try:
        data = json.loads(CLASSIFICATION_VERSION_FILE.read_text(encoding="utf-8"))
        return str(data.get("version", ""))
    except (OSError, ValueError, AttributeError):
        return ""


def install_reference(payload, version, expected_sha256=None):
    if not str(version).strip():
        raise ValueError("분류군 자료 버전을 입력해 주세요.")
    digest = hashlib.sha256(payload).hexdigest()
    if expected_sha256 and expected_sha256 != digest:
        raise ValueError("분류군 자료의 체크섬이 일치하지 않습니다.")
    records = validate_records(json.loads(payload.decode("utf-8")))
    CLASSIFICATION_DIR.mkdir(parents=True, exist_ok=True)
    temporary = CLASSIFICATION_FILE.with_suffix(".tmp")
    temporary.write_bytes(payload)
    temporary.replace(CLASSIFICATION_FILE)
    manifest = {"version": str(version).strip(), "sha256": digest, "rows": len(records), "matching_rank": "family", "genus_matching": False}
    temporary = CLASSIFICATION_VERSION_FILE.with_suffix(".tmp")
    temporary.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(CLASSIFICATION_VERSION_FILE)
    return manifest


def import_reference_excel(path, version):
    with pd.ExcelFile(path) as workbook:
        frame = workbook.parse(workbook.sheet_names[0])
    if any(column not in frame for column in REQUIRED_COLUMNS):
        raise ValueError("분류군정보 엑셀의 필수 컬럼이 없습니다.")
    def clean(value):
        if pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()
    records = [{column: clean(record[column]) for column in REQUIRED_COLUMNS} for record in frame.to_dict("records")]
    payload = (json.dumps(records, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
    return install_reference(payload, version)
