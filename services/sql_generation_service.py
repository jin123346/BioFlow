"""Static Oracle INSERT generation and validation; never connects to a database."""

from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation
import posixpath
import re
from string import Formatter

from config.sql_media_profiles import PROFILES
from services.sql_excel_service import MissingFormulaValue
from services.sql_file_lookup_service import MediaFileIndex
from utils.text_utils import normalize_text


FILENAME_COLUMNS = {"STORE_FILE_NM", "ORGIL_FILE_NM", "PH500_FILE_NM", "PH192_FILE_NM", "LVBNG_MVPCT_RPRSN_IMAGE"}


@dataclass
class SQLResult:
    excel_row: int
    status: str
    sql: str = ""
    error: str = ""
    key: str = ""
    filename: str = ""
    sequence: str = ""


def identifier(value):
    if not re.fullmatch(r"[A-Z][A-Z0-9_]*", str(value)):
        raise ValueError("올바르지 않은 DB 식별자입니다.")
    return value


def header_key(value):
    return re.sub(r"\s+", "", re.split(r"[(（]", str(value), maxsplit=1)[0]).casefold()


def apply_image_sequence_rule(profile):
    """Keep confirmed fixed rules independent of stale saved mappings."""
    if profile["kind"] in ("observation_image", "specimen_image"):
        expected = PROFILES[profile["kind"]]
        old_key = profile["key"]
        profile["key"] = expected["key"]
        if not profile.get("nas_template"):
            profile["nas_template"] = expected["nas_template"]
        profile.update(sequence_mode="fixed", sequence_value="1", sequence_confirmed=True)
        for column in profile["columns"]:
            if column["name"] == old_key:
                column.update(name=expected["key"], type="VARCHAR2", aliases=expected["columns"][0]["aliases"])
                if old_key != expected["key"]:
                    column.update(source=None, constant=None)
            if column["name"] == "IMAGE_SEQ":
                column.update(source=None, constant="1", type="CHAR", length=1, required=True,
                              computed="sequence", folder_lookup=False, filename_extension="")
    if profile["kind"] in ("observation_video", "observation_sound"):
        profile.update(use_codes=["Y"], use_codes_confirmed=True)
        use_column = "LVBNG_MVPCT_USE_YN" if profile["kind"] == "observation_video" else "LVBNG_SOUND_USE_YN"
        for column in profile["columns"]:
            if column["name"] == use_column:
                column.update(source=None, constant="Y", required=True, computed=None,
                              folder_lookup=False, filename_extension="")
    expected = PROFILES[profile["kind"]]
    # Add newly confirmed columns while retaining the user's source mappings.
    if "image" in profile["kind"] and not any(c["name"] == "FILE_DTL_CONT" for c in profile["columns"]):
        profile["columns"].append(deepcopy(expected["columns"][-1]))
    schema = {c["name"]: c for c in expected["columns"]}
    for col in profile["columns"]:
        if col["name"] in schema:
            original = schema[col["name"]]
            col.update(type=original["type"], length=original["length"], length_unit=original["length_unit"])
    profile["schema_confirmed"] = True
    return profile


def auto_map_profile(profile, headers):
    profile = apply_image_sequence_rule(profile)
    if not profile.get("nas_template"):
        profile["nas_template"] = PROFILES[profile["kind"]]["nas_template"]
    lookup = {header_key(header): header for header in headers}
    for column in profile["columns"]:
        # Migrate earlier saved video mappings whose English fields were fixed NULL.
        if profile["kind"] == "observation_video" and column["name"] in ("LVBNG_MVPCT_SUBJC_ENG", "LVBNG_MVPCT_CONT_ENG"):
            column["computed"] = None
            if not column.get("aliases"):
                column["aliases"] = next(item["aliases"] for item in PROFILES[profile["kind"]]["columns"] if item["name"] == column["name"])
        if column.get("source") not in headers:
            candidates = [column["name"], *column.get("aliases", [])]
            column["source"] = next((lookup[header_key(name)] for name in candidates if header_key(name) in lookup), None)
    for field, alias in (("type_column", "미디어타입"), ("code_column", "미디어코드"),
                         ("database_number_column", "데이터베이스번호"),
                         ("vernacular_name_column", "국명"), ("scientific_name_column", "학명"),
                         ("organization_code_column", "기관코드")):
        if profile.get(field) not in headers:
            profile[field] = lookup.get(header_key(alias))
    if "image" in profile["kind"]:
        profile["organization_code_column"] = lookup.get(header_key("기관코드"))
    if not profile.get("database_number_column"):
        profile["database_number_column"] = lookup.get(header_key("이전데이터베이스번호"))
    # The user has explicitly selected this SQL type. Completed NARIS sheets lack type columns.
    profile["all_rows_confirmed"] = not profile.get("type_column") and not profile.get("code_column")
    return apply_image_sequence_rule(profile)


def duration_hhmm(value):
    if value is None or normalize_text(value) == "":
        return None
    if isinstance(value, (datetime, time)):
        hours, minutes = value.hour, value.minute
    elif isinstance(value, timedelta):
        total = int(value.total_seconds())
        if total < 0:
            raise ValueError("음수 재생시간")
        hours, minutes = total // 3600, (total % 3600) // 60
    elif isinstance(value, (float, Decimal)) and 0 <= value < 1:
        total = int(Decimal(str(value)) * 86400 + Decimal("0.000001"))
        hours, minutes = total // 3600, (total % 3600) // 60
    else:
        match = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?", str(value).strip())
        if not match:
            raise ValueError("재생시간은 HH:MM 또는 HH:MM:SS 형식이어야 합니다.")
        hours, minutes, seconds = int(match[1]), int(match[2]), int(match[3] or 0)
        if seconds > 59:
            raise ValueError("재생시간 초는 00~59여야 합니다.")
    if hours > 99 or not 0 <= minutes <= 59:
        raise ValueError("재생시간은 두 자리 시와 00~59분 범위여야 합니다.")
    return f"{hours:02d}:{minutes:02d}"


def number(value, positive_integer=False):
    if isinstance(value, bool):
        raise ValueError("논리값은 숫자로 사용할 수 없습니다.")
    text = normalize_text(value)
    if not text:
        return None
    try:
        result = Decimal(text)
    except InvalidOperation:
        raise ValueError("숫자 형식이 아닙니다.") from None
    if not result.is_finite():
        raise ValueError("유한한 숫자가 아닙니다.")
    if positive_integer and (result <= 0 or result != result.to_integral_value()):
        raise ValueError("양의 정수여야 합니다.")
    if positive_integer and isinstance(value, float) and abs(value) >= 2 ** 53:
        raise ValueError("엑셀 숫자 정밀도를 확인해 주세요. 큰 ID는 텍스트로 입력해야 합니다.")
    precision = len("".join(str(digit) for digit in result.as_tuple().digits).lstrip("0").rstrip("0")) or 1
    if precision > 38 or (result and not -130 <= result.adjusted() <= 125):
        raise ValueError("Oracle NUMBER 범위를 확인해 주세요.")
    return format(result, "f") if result != result.to_integral_value() else str(int(result))


def string_literal(value):
    # Encode line breaks as expressions so SQL*Plus cannot split a literal across blank lines.
    parts = re.split(r"([\r\n])", value)
    return " || ".join("CHR(13)" if part == "\r" else "CHR(10)" if part == "\n"
                       else "'" + part.replace("'", "''") + "'" for part in parts)


def nas_path(template, values):
    template = str(template).strip()
    if not template:
        raise ValueError("NAS 저장 경로를 설정해 주세요.")
    try:
        if any(field is not None and (field not in values or spec or conversion)
               for _, field, spec, conversion in Formatter().parse(template)):
            raise ValueError("Unknown path placeholder")
        path = template.format(**values)
    except (KeyError, ValueError, AttributeError, IndexError):
        raise ValueError("NAS 경로 치환값을 확인해 주세요. {SPECS_ESSNT_ID}, {OBSRV_INFO_SEQ}, {기관코드}를 사용할 수 있습니다.") from None
    if not path.startswith("/") or "\\" in path or re.search(r"[\x00-\x1f]", path):
        raise ValueError("NAS 경로는 /로 시작하는 서버 폴더 경로여야 합니다. 로컬 PC 경로는 사용할 수 없습니다.")
    if ".." in path.split("/"):
        raise ValueError("NAS 경로에 상위 폴더 이동(..)을 사용할 수 없습니다.")
    return posixpath.normpath(re.sub(r"/+", "/", path)).rstrip("/") + "/"


def mapped_nas_path(profile, row, key):
    values = {profile["key"]: key}
    template = profile["nas_template"]
    if "{기관코드}" in template:
        source = profile.get("organization_code_column")
        if "image" in profile["kind"]:
            source = next((header for header in row if header_key(header) == header_key("기관코드")), None)
        value = row.get(source) if source else (None if "image" in profile["kind"] else profile.get("organization_code"))
        if isinstance(value, MissingFormulaValue):
            raise ValueError("기관코드 수식을 엑셀에서 재계산 후 저장해 주세요.")
        code = normalize_text(value)
        if not code:
            raise ValueError("엑셀의 기관코드 컬럼과 해당 행의 기관코드 값을 확인해 주세요.")
        if code in (".", "..") or re.search(r"[/\\\s\x00-\x1f{}]", code):
            raise ValueError("기관코드는 경로 구분자나 공백 없는 폴더명이어야 합니다.")
        values["기관코드"] = code
    return nas_path(template, values)


def mapped_value(column, row):
    source = column.get("source")
    if source:
        if source not in row:
            raise ValueError(f"{column['name']}: 매핑한 엑셀 컬럼이 없습니다 ({source}).")
        if isinstance(row[source], MissingFormulaValue):
            raise ValueError(f"{column['name']}: 수식 계산값이 없습니다. 엑셀에서 재계산 후 다시 저장해 주세요.")
        value = row[source]
    else:
        value = column.get("constant")
    extension = str(column.get("filename_extension") or "").strip()
    if extension:
        if column["name"] not in FILENAME_COLUMNS:
            raise ValueError(f"{column['name']}: 확장자는 파일명 컬럼에만 지정할 수 있습니다.")
        extension = extension if extension.startswith(".") else "." + extension
        if not re.fullmatch(r"\.[A-Za-z0-9]{1,12}", extension):
            raise ValueError(f"{column['name']}: 올바른 확장자를 입력해 주세요 (예: .jpg, .mp4, .wav).")
        if value is not None and normalize_text(value):
            filename = value.strip() if isinstance(value, str) else normalize_text(value)
            # Keep an existing extension, including its case; never turn a.png into a.png.jpg.
            existing = re.search(r"\.([A-Za-z0-9]{1,12})$", filename)
            has_extension = existing and (any(char.isalpha() for char in existing[1]) or filename.lower().endswith(extension.lower()))
            value = filename if has_extension else filename + extension
    return value


def filter_row(row, profile):
    type_field, code_field = profile.get("type_column"), profile.get("code_column")
    media = normalize_text(row.get(type_field)).casefold() if type_field else ""
    code = normalize_text(row.get(code_field)) if code_field else ""
    target_names = {value.casefold() for value in profile["media_names"]}
    other_names = {name.casefold() for kind, item in PROFILES.items() if kind != profile["kind"]
                   for name in item["media_names"]} - target_names
    type_match = True if media in target_names else False if media in other_names else None
    if media and type_match is None:
        return "ERROR", f"확인되지 않은 미디어타입: {media}"
    code_match = True if code and code in profile.get("target_codes", []) else (
        False if code and code in profile.get("other_codes", []) else None)
    if type_match is False and code_match is not True:
        return "SKIP", "선택한 미디어 유형에 해당하지 않는 행"
    if code and code_match is None:
        return "ERROR", f"미디어코드 '{code}'의 유형 매핑이 미확정입니다. OB 등을 임의 해석하지 않습니다."
    if type_match is not None and code_match is not None and type_match != code_match:
        return "ERROR", "미디어타입과 미디어코드가 일치하지 않습니다."
    match = type_match if type_match is not None else code_match
    if match is None:
        if not type_field and not code_field and profile.get("all_rows_confirmed"):
            return "SUCCESS", ""
        return "ERROR", "미디어 유형을 판별할 값이 없습니다. 필터 매핑 또는 전체 행 유형을 확인해 주세요."
    return ("SUCCESS", "") if match else ("SKIP", "선택한 미디어 유형에 해당하지 않는 행")


def _sql_value(column, value, profile):
    label = column["name"]
    empty = value is None or normalize_text(value) == ""
    if empty:
        if column.get("required"):
            raise ValueError(f"{label}: 필수값 누락")
        return "NULL"
    data_type = column["type"]
    if data_type == "NUMBER":
        result = number(value, label in (profile["key"], profile["sequence"]))
        if label == "FILE_MGNTD" and Decimal(result) < 0:
            raise ValueError("파일크기는 음수일 수 없습니다.")
        return result
    if data_type == "DATE":
        if not isinstance(value, (date, datetime)):
            try:
                value = datetime.fromisoformat(str(value))
            except ValueError:
                raise ValueError(f"{label}: 올바른 날짜가 아닙니다.") from None
        return f"TO_DATE('{value.strftime('%Y-%m-%d')}','YYYY-MM-DD')"
    if data_type not in ("VARCHAR2", "CHAR"):
        raise ValueError(f"{label}: 지원되지 않는 DB 타입")
    text = normalize_text(value) if not isinstance(value, str) else value.strip()
    if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", text):
        raise ValueError(f"{label}: 지원되지 않는 제어 문자")
    length = column.get("length")
    if not isinstance(length, int) or length <= 0:
        raise ValueError(f"{label}: DB 문자열 최대 길이가 미확정입니다. 매핑 화면의 ‘최대 길이’에 실제 DB 컬럼 길이를 입력해 주세요.")
    try:
        size = len(text) if column.get("length_unit") == "CHAR" else len(text.encode(profile.get("encoding", "utf-8")))
    except (UnicodeError, LookupError):
        raise ValueError(f"{label}: DB 문자셋 인코딩을 확인해 주세요.") from None
    if size > length:
        raise ValueError(f"{label}: 문자열 길이 초과 ({size}/{length} {column.get('length_unit', 'BYTE')})")
    if len(text.encode(profile.get("encoding", "utf-8"))) > 4000:
        raise ValueError(f"{label}: 기본 Oracle SQL 문자열 리터럴의 4000바이트 범위를 초과합니다.")
    if label in FILENAME_COLUMNS:
        if "/" in text or "\\" in text or text in (".", "..") or re.search(r"[\r\n]", text):
            raise ValueError(f"{label}: 경로가 아닌 파일명만 지정해 주세요.")
    if label == "IMAGE_TPCD":
        if not profile.get("use_codes_confirmed") or text not in profile.get("value_codes", ["03"]):
            raise ValueError("IMAGE_TPCD: 확인되지 않은 이미지구분코드")
    if label.endswith("_USE_YN"):
        if not profile.get("use_codes_confirmed") or text not in profile.get("use_codes", []):
            raise ValueError(f"{label}: 사용 여부 코드 설정·확인이 필요합니다.")
    return string_literal(text)


def generate_sql(sheet, profile):
    profile = apply_image_sequence_rule(deepcopy(profile))
    expected = PROFILES[profile["kind"]]
    if (profile["table"], profile["key"], profile["sequence"]) != (expected["table"], expected["key"], expected["sequence"]):
        raise ValueError("대상 테이블 또는 연결 키 설정이 올바르지 않습니다.")
    table, key, sequence = (identifier(profile[name]) for name in ("table", "key", "sequence"))
    if [c["name"] for c in profile["columns"]] != [c["name"] for c in expected["columns"]]:
        raise ValueError("DB 컬럼 설정이 테이블과 일치하지 않습니다.")
    for configured, original in zip(profile["columns"], expected["columns"]):
        if original["required"] and not configured.get("required"):
            raise ValueError(f"{original['name']}: 요구사항의 필수값을 선택 해제할 수 없습니다.")
        allowed = ("VARCHAR2", "CHAR") if original["type"] == "VARCHAR2" else (original["type"],)
        if configured["type"] not in allowed:
            raise ValueError(f"{original['name']}: 요구사항과 다른 DB 타입입니다.")
        if original["name"] == "LVBNG_MVPCT_RPRSN_IMAGE" and (configured.get("length") or 100) > 100:
            raise ValueError("대표 썸네일 파일명 최대 길이는 100입니다.")
    file_index = None
    if any(column.get("folder_lookup") for column in profile["columns"]):
        file_index = MediaFileIndex(profile.get("media_lookup_dir"), profile.get("lookup_recursive", True), profile.get("lookup_skip_previews", True))
    results, duplicates = [], defaultdict(list)
    for excel_row, row in sheet.rows:
        status, reason = filter_row(row, profile)
        result = SQLResult(excel_row, status, error=f"엑셀 {excel_row}행: {reason}" if reason else "")
        results.append(result)
        if status != "SUCCESS":
            continue
        errors, expressions, raw = [], {}, {}
        if not profile.get("schema_confirmed"):
            errors.append("매핑 화면 하단의 ‘실제 DB 스키마의 컬럼 타입·최대 길이·필수값·연결 키를 확인했습니다’를 확인해 주세요.")
        if not profile.get("sequence_confirmed"):
            errors.append("SEQ 채번 규칙 확인이 필요합니다.")
        mode = profile.get("sequence_mode")
        if mode not in ("fixed", "max_per_key", "source"):
            errors.append("지원되는 SEQ 규칙을 선택해 주세요.")
        for col in profile["columns"]:
            try:
                if col["name"] == sequence and mode in ("fixed", "max_per_key"):
                    continue
                if col.get("folder_lookup"):
                    if col["name"] not in FILENAME_COLUMNS:
                        raise ValueError("폴더 파일명 매칭은 파일명 컬럼에서만 사용할 수 있습니다.")
                    raw[col["name"]] = file_index.filename(row.get(profile.get("database_number_column")), profile["kind"], col)
                else:
                    raw[col["name"]] = mapped_value(col, row)
            except ValueError as exc:
                errors.append(str(exc))
        try:
            result.key = normalize_text(raw.get(key)) if "image" in profile["kind"] else number(raw.get(key), True) or ""
            if not result.key:
                raise ValueError(f"{key}: 필수 연결 키 누락. 다른 ID로 대신할 수 없습니다.")
        except ValueError as exc:
            errors.append(f"{key}: {exc}")
        try:
            raw["FILE_STORE_CRSE"] = mapped_nas_path(profile, row, result.key)
        except ValueError as exc:
            errors.append(f"FILE_STORE_CRSE: {exc}")
        if mode == "fixed":
            raw[sequence] = profile.get("sequence_value")
        elif mode == "max_per_key":
            expressions[sequence] = f"NVL(MAX({sequence}), 0) + 1"
        if profile["kind"] == "observation_video":
            for field, name_column in (
                ("LVBNG_MVPCT_SUBJC_KOR", "vernacular_name_column"),
                ("LVBNG_MVPCT_CONT_KOR", "vernacular_name_column"),
                ("LVBNG_MVPCT_SUBJC_ENG", "scientific_name_column"),
                ("LVBNG_MVPCT_CONT_ENG", "scientific_name_column"),
            ):
                if not normalize_text(raw.get(field)):
                    fallback = row.get(profile.get(name_column))
                    if isinstance(fallback, MissingFormulaValue):
                        errors.append(f"{field}: 대체할 국명/학명 수식을 엑셀에서 재계산 후 저장해 주세요.")
                    else:
                        raw[field] = fallback
            raw_database_number = row.get(profile.get("database_number_column"))
            if isinstance(raw_database_number, MissingFormulaValue):
                errors.append("데이터베이스번호 수식을 엑셀에서 재계산 후 저장해 주세요.")
            database_number = normalize_text(raw_database_number)
            if not database_number:
                errors.append("동영상 내용에 추가할 데이터베이스번호의 매핑 또는 값이 없습니다.")
            else:
                field = "LVBNG_MVPCT_CONT_KOR"
                description = raw.get(field)
                description = description.strip() if isinstance(description, str) else normalize_text(description)
                marker = f"<br> : {database_number}"
                raw[field] = description if description.endswith(marker) else description + marker
        for col in profile["columns"]:
            name = col["name"]
            if col.get("computed") == "null":
                if col.get("required"):
                    errors.append(f"{name}: NULL 고정 항목을 필수값으로 설정할 수 없습니다.")
                expressions[name] = "NULL"
                continue
            if col.get("computed") == "sysdate":
                expressions[name] = "SYSDATE"
                continue
            if name in expressions:
                continue
            try:
                if name == "FRST_REGISTER_ID":
                    raw[name] = "narisorg"
                if col.get("computed") == "duration":
                    raw[name] = duration_hhmm(raw.get(name))
                if name.endswith("_USE_YN") and not profile.get("use_codes_confirmed"):
                    raise ValueError(f"{name}: 실제 사용 여부 코드 확인이 필요합니다.")
                expressions[name] = _sql_value(col, raw.get(name), profile)
            except (ValueError, TypeError) as exc:
                detail = str(exc)
                message = detail if detail.startswith(f"{name}:") else f"{name}: {detail}"
                if col["type"] == "NUMBER" and raw.get(name) is not None:
                    value = repr(raw[name])[:100]
                    source = col.get("source") or "고정값"
                    message += f" (매핑: {source}, 값: {value}). 헤더 행과 컬럼 매핑을 확인해 주세요."
                errors.append(message)
        filename_field = "PH500_FILE_NM" if profile["kind"] == "observation_video" else (
            "PH192_FILE_NM" if profile["kind"] == "observation_sound" else "STORE_FILE_NM")
        result.filename = normalize_text(raw.get(filename_field))
        if result.key and result.filename:
            duplicates[("file", result.key, result.filename)].append(result)
        if mode in ("fixed", "source"):
            try:
                result.sequence = number(raw.get(sequence), True) or ""
                if result.key and result.sequence:
                    duplicates[("sequence", result.key, result.sequence)].append(result)
            except ValueError:
                pass
        if not errors:
            columns = ", ".join(identifier(col["name"]) for col in profile["columns"])
            values = ", ".join(expressions[col["name"]] for col in profile["columns"])
            if mode == "max_per_key":
                result.sql = f"INSERT INTO {table} ({columns}) SELECT {values} FROM {table} WHERE {key} = {result.key};"
            else:
                result.sql = f"INSERT INTO {table} ({columns}) VALUES ({values});"
            if len(result.sql) > 32767:
                errors.append("INSERT SQL이 XLSX 셀 최대 문자 수(32767자)를 초과합니다.")
        if errors:
            result.status, result.sql = "ERROR", ""
            result.error = f"엑셀 {excel_row}행: " + "; ".join(dict.fromkeys(errors))
    for group, members in duplicates.items():
        if len(members) > 1:
            rows = ", ".join(str(item.excel_row) for item in members)
            reason = "같은 연결 ID의 파일명 중복" if group[0] == "file" else "같은 연결 ID의 고정/입력 SEQ 중복"
            for result in members:
                result.status, result.sql = "ERROR", ""
                result.error = (result.error + "; " if result.error else f"엑셀 {result.excel_row}행: ") + f"{reason} (행 {rows})"
    return results
