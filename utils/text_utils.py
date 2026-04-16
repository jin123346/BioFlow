from utils.excel_utils import safe_int
import re
import pandas as pd

def normalize_text(value) -> str:
    if value is None or pd.isna(value):
        return ""

    text = str(value)

    # 엑셀 특수 공백 제거
    text = text.replace("\xa0", " ")

    # 다중 공백 정리
    text = re.sub(r"\s+", " ", text)

    text = text.strip()

    # 문자열 nan 제거
    if text.lower() == "nan":
        return ""

    return text

def normalize_scientific_name(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    # 저자명 제거용 단순 정규화: 앞의 속명/종소명/하위분류까지만 비교
    parts = text.split()
    if len(parts) >= 2:
        # 예: Genus species / Genus species var. xxx 까지 허용
        keep = []
        for part in parts:
            if re.search(r"[()0-9,]", part):
                break
            keep.append(part)
            if len(keep) >= 4:
                break
        if len(keep) >= 2:
            return " ".join(keep).lower()
    return text.lower()

def build_db_no(inst_code: str, media_code: str, taxon_code: str, year: object, seq: object, no_value: object) -> str:
    inst_code = normalize_text(inst_code)
    taxon_code = normalize_text(taxon_code)
    if not inst_code or not taxon_code:
        return ""

    year_num = normalize_text(year)
    year_2 = year_num[-2:] if len(year_num) >= 2 else year_num.zfill(2)
    seq_num = f"{safe_int(seq):02d}"
    no_num = f"{safe_int(no_value):05d}"
    return f"{inst_code}-{media_code}-{taxon_code}-{year_2}{seq_num}{no_num}"