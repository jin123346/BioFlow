"""UI-independent Excel loading, normalization, and base-file output."""

from __future__ import annotations

import io
import itertools
import re
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd


def open_excel_lenient(file_path):
    """Open workbooks from Excel, Hancom, and WPS with missing style names."""
    try:
        return pd.ExcelFile(file_path)
    except TypeError as exc:
        if "_NamedCellStyle" not in str(exc):
            raise
    counter = itertools.count()
    buffer = io.BytesIO()
    with zipfile.ZipFile(file_path) as source, zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item)
            if item.filename == "xl/styles.xml":
                data = re.sub(
                    rb"<((?:\w+:)?cellStyle)(?=[\s/>])(?![^>]*\bname=)",
                    lambda match: b'<%s name="Unnamed%d"' % (match.group(1), next(counter)),
                    data,
                )
            target.writestr(item, data)
    buffer.seek(0)
    return pd.ExcelFile(buffer)


def parse_excel_sheet(workbook, sheet_name):
    name = workbook.sheet_names[sheet_name] if isinstance(sheet_name, int) else sheet_name
    sheet = workbook.book[name]
    first = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    second = [cell.value for cell in next(sheet.iter_rows(min_row=2, max_row=2))]
    header_row = 1 if "종고유ID" not in first and "종고유ID" in second and "학명" in second else 0
    frame = workbook.parse(sheet_name=sheet_name, header=header_row)
    frame.attrs["excel_epoch"] = workbook.book.epoch
    return frame


def read_excel_lenient(file_path, sheet_name=0):
    with open_excel_lenient(file_path) as workbook:
        return parse_excel_sheet(workbook, sheet_name)


def normalize_excel_date(value: Any, epoch: datetime = datetime(1899, 12, 30)) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()
    if text.isdigit() and len(text) == 6:
        return f"20{text[0:2]}-{text[2:4]}-{text[4:6]}"
    if text.isdigit() and len(text) == 8:
        return f"{text[0:4]}-{text[4:6]}-{text[6:8]}"
    if re.fullmatch(r"\d{1,5}(?:\.\d+)?", text):
        serial = float(text)
        if 1 <= serial < 100000:
            if epoch == datetime(1899, 12, 30) and serial < 60:
                serial += 1
            return (epoch + timedelta(days=serial)).strftime("%Y-%m-%d")
    return text


def write_records(path: str | Path, records, columns):
    pd.DataFrame(records, columns=columns).to_excel(path, index=False)


def derived_excel_path(file_path: str | Path, suffix: str) -> Path:
    path = Path(file_path)
    return path.with_name(f"{path.stem}{suffix}{path.suffix}")
