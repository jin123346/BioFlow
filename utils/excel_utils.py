import pandas as pd
import re

def safe_int(value: object, default: int = 0) -> int:
    if value is None or str(value).strip() == "":
        return default
    try:
        return int(float(value))
    except Exception:
        return default
    
    
def normalize_sheet_name(sheet_name: str) -> str:
    if sheet_name is None:
        return ""
    text = str(sheet_name).strip()
    text = text.replace("\n", "")
    text = re.sub(r"\s+", "", text)
    return text


def find_target_sheet_name(file_path: str, target_sheet_name: str) -> str | None:
    excel_file = pd.ExcelFile(file_path)
    normalized_target = normalize_sheet_name(target_sheet_name)

    for sheet_name in excel_file.sheet_names:
        if normalize_sheet_name(sheet_name) == normalized_target:
            return sheet_name
    return None