from __future__ import annotations

import io
import itertools
import json
import logging
import re
import shutil
import subprocess
import time
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
import zipfile

import pandas as pd
from PySide6.QtCore import QRect, QSize, Qt, QTimer, QUrl
from PySide6.QtGui import QBrush, QColor, QCursor, QDesktopServices, QPixmap

logger = logging.getLogger("bioflow.upload_tab")
DB_NUMBER_MANUAL_ROLE = Qt.UserRole + 1
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QCheckBox,
    QCompleter,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTableWidgetSelectionRange,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from config.paths import DATA_DIR, DRAFT_DIR, SETTINGS_DIR
from services.address_master_service import AddressMasterService
from ui.widgets.recent_file_dialog import RecentFileDialog as QFileDialog
from services.species_master_service import SpeciesMasterService
from services.map_preview_service import publish_map_preview
from services.final_excel_service import write_final_template
from services.gbif_excel_service import write_gbif_template
from services.media_storage_service import copy_database_media, find_media_tool, media_destination
from ui.widgets.column_mapping_dialog import ColumnMappingDialog, apply_mapping, auto_map
from ui.widgets.batch_apply_dialog import BatchApplyDialog
from utils.text_utils import normalize_text

try:
    from PIL import ExifTags, Image
except ImportError:
    ExifTags = None
    Image = None


def open_excel_lenient(file_path):
    """한셀/WPS 등이 저장한 xlsx는 이름 없는 셀 스타일이 있어 openpyxl이 거부함 → styles.xml에 이름을 채워 재시도."""
    try:
        return pd.ExcelFile(file_path)
    except TypeError as exc:
        if "_NamedCellStyle" not in str(exc):
            raise
    counter = itertools.count()
    buf = io.BytesIO()
    with zipfile.ZipFile(file_path) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item)
            if item.filename == "xl/styles.xml":
                # <cellStyle> 또는 <x:cellStyle> 처럼 네임스페이스 접두사가 붙은 경우 모두 처리
                data = re.sub(
                    rb"<((?:\w+:)?cellStyle)(?=[\s/>])(?![^>]*\bname=)",
                    lambda m: b'<%s name="Unnamed%d"' % (m.group(1), next(counter)),
                    data,
                )
            dst.writestr(item, data)
    buf.seek(0)
    return pd.ExcelFile(buf)


def read_excel_lenient(file_path, sheet_name=0):
    with open_excel_lenient(file_path) as workbook:
        return parse_excel_sheet(workbook, sheet_name)


def parse_excel_sheet(workbook, sheet_name):
    name = workbook.sheet_names[sheet_name] if isinstance(sheet_name, int) else sheet_name
    sheet = workbook.book[name]
    first = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    second = [cell.value for cell in next(sheet.iter_rows(min_row=2, max_row=2))]
    header_row = 1 if "종고유ID" not in first and "종고유ID" in second and "학명" in second else 0
    frame = workbook.parse(sheet_name=sheet_name, header=header_row)
    frame.attrs["excel_epoch"] = workbook.book.epoch
    return frame


REFERENCE_DIR = DATA_DIR / "reference"
PROJECTS_FILE = REFERENCE_DIR / "upload_projects.json"
ORGANIZATIONS_FILE = REFERENCE_DIR / "organizations.json"
DRAFT_FILE = DRAFT_DIR / "new_data_draft.json"
COLUMN_MAPPINGS_FILE = SETTINGS_DIR / "excel_column_mappings.json"

ROW_COLUMNS = [
    "NO.",
    "종고유ID",
    "미디어타입",
    "미디어코드",
    "원본경로",
    "원본파일명",
    "파일경로",
    "파일명",
    "동영상 제목",
    "동영상 상세설명",
    "동영상 재생시간",
    "동영상 파일명",
    "사운드 제목",
    "사운드 상세설명",
    "사운드 재생시간",
    "사운드 파일명",
    "분류코드",
    "사업명",
    "사업년도",
    "순번",
    "사업별 순번",
    "기관명",
    "기관코드",
    "국명",
    "학명",
    "이전목록번호 국문명",
    "이전목록번호 영문명",
    "데이터베이스번호",
    "동정자 국문",
    "동정자 영문",
    "동정년월일",
    "성별",
    "연령",
    "관찰개체수",
    "원본여부",
    "관찰 공개여부",
    "대륙/해양 국문",
    "대륙/해양 영문",
    "위도",
    "경도",
    "국가국문명",
    "국가영문명",
    "관찰위치 시도 국문",
    "관찰위치 시도 영문",
    "관찰위치 시군구 국문",
    "관찰위치 시군구 영문",
    "관찰위치 읍면동 국문",
    "관찰위치 읍면동 영문",
    "관찰위치 상세 국문",
    "관찰위치 상세 영문",
    "관찰자 국문",
    "관찰자 영문",
    "관찰일",
    "원산지",
    "원산지 공개여부",
    "자원공개여부",
    "국외반출여부",
    "분양가능여부",
]
PROJECT_COLUMNS = ["사업년도", "순번", "사업명"]
ORGANIZATION_COLUMNS = ["기관명", "기관코드"]

# 완성본(재구성) 엑셀 컬럼 순서: (완성본 컬럼명, 작성 폼의 원본 컬럼명 또는 None=항상 빈칸)
FINAL_EXPORT_COLUMNS: list[tuple[str, str | None]] = [
    ("관찰정보순번", None),
    ("종고유ID", "종고유ID"),
    ("학명", "학명"),
    ("기관코드", "기관코드"),
    ("기관명", "기관명"),
    ("동정자 국문", "동정자 국문"),
    ("동정자 영문", "동정자 영문"),
    ("동정년월일", "동정년월일"),
    ("대륙해양국문명", "대륙/해양 국문"),
    ("대륙해양영문명", "대륙/해양 영문"),
    ("국가 국문명", "국가국문명"),
    ("국가 영문명", "국가영문명"),
    ("경도", "경도"),
    ("위도", "위도"),
    ("성별", "성별"),
    ("연령", "연령"),
    ("이전목록번호국문명", "이전목록번호 국문명"),
    ("이전목록번호영문명", "이전목록번호 영문명"),
    ("데이터베이스번호", "데이터베이스번호"),
    ("원본여부", "원본여부"),
    ("관찰정보 공개여부", "관찰 공개여부"),
    ("관찰위치 시도 국문", "관찰위치 시도 국문"),
    ("관찰위치 시도 영문", "관찰위치 시도 영문"),
    ("관찰위치 시군구 국문", "관찰위치 시군구 국문"),
    ("관찰위치 시군구 영문", "관찰위치 시군구 영문"),
    ("관찰위치읍면동국문명", "관찰위치 읍면동 국문"),
    ("관찰위치읍면동영문명", "관찰위치 읍면동 영문"),
    ("관찰위치 상세 국문", "관찰위치 상세 국문"),
    ("관찰위치 상세 영문", "관찰위치 상세 영문"),
    ("관찰자 국문", "관찰자 국문"),
    ("관찰자 영문", "관찰자 영문"),
    ("관찰년월일", "관찰일"),
    ("사업년도", "사업년도"),
    ("순번", "순번"),
    ("사업명", "사업명"),
    ("원산지", "원산지"),
    ("원산지 공개여부", "원산지 공개여부"),
    ("자원공개여부", "자원공개여부"),
    ("국외반출여부", "국외반출여부"),
    ("분양가능여부", "분양가능여부"),
    ("관리번호", None),
    ("개체수", "관찰개체수"),
    ("최소고도", None),
    ("최대고도", None),
    ("최소수심", None),
    ("최대수심", None),
]

DROPDOWN_OPTIONS = {
    "미디어타입": ["정보", "이미지", "동영상", "사운드"],
    "분류코드": ["IN", "PL", "BI", "PI", "MM", "SP", "MR", "AR", "AM", "SW", "AG"],
    "성별": ["암컷", "수컷", "양성", "확인불가"],
    "연령": ["성체", "어린개체", "중간개체", "확인불가"],
    "원본여부": ["원본", "사본"],
    "관찰 공개여부": ["공개", "비공개"],
    "대륙/해양 국문": ["아시아", "태평양"],
    "대륙/해양 영문": ["Asia", "Pacific Ocean"],
    "원산지 공개여부": ["공개", "비공개"],
    "자원공개여부": ["공개", "비공개"],
    "국외반출여부": ["가능", "미반출"],
    "분양가능여부": ["가능", "불가능"],
}

DROPDOWN_DEFAULTS = {
    "성별": "확인불가",
    "연령": "확인불가",
    "원본여부": "원본",
    "관찰 공개여부": "공개",
    "원산지 공개여부": "공개",
    "자원공개여부": "공개",
    "국외반출여부": "미반출",
}

# 대부분 국내 표본/관찰이라 기본값으로 채워두고, 해외 건이면 직접 고치는 방식
TEXT_DEFAULTS = {
    "국가국문명": "대한민국",
    "국가영문명": "Republic of Korea",
}

DATE_FIELDS = {"동정년월일", "관찰일", "채집년월일"}

# 위도/경도 오류 표시용 배경색, 대한민국 대략 범위(위도/경도 뒤바뀜, 국외 여부 판단용)
LOCATION_ERROR_COLOR = QColor("#ffcdd2")
KOREA_LAT_RANGE = (33.0, 39.0)
KOREA_LON_RANGE = (124.0, 132.0)

# 사업/기관/미디어타입/분류코드는 여러 행을 연달아 입력할 때 세션 내내 유지되는 값
STICKY_FIELDS = {"사업명", "사업년도", "순번", "기관명", "기관코드", "미디어타입", "분류코드"}

VIDEO_FIELDS = ["동영상 제목", "동영상 상세설명", "동영상 재생시간", "동영상 파일명"]
SOUND_FIELDS = ["사운드 제목", "사운드 상세설명", "사운드 재생시간", "사운드 파일명"]

ADDRESS_FIELDS = {
    "관찰위치 시도 국문": ("관찰위치 시도 영문", "sido"),
    "관찰위치 시군구 국문": ("관찰위치 시군구 영문", "sigungu"),
    "관찰위치 읍면동 국문": ("관찰위치 읍면동 영문", "eupmyeondong"),
    "관찰위치 상세 국문": ("관찰위치 상세 영문", "detail"),
}

MEDIA_TYPE_CODES = {
    "정보": "OB",
    "이미지": "OI",
    "동영상": "OM",
    "사운드": "OS",
}

MEDIA_EXTENSION_TYPES = {
    ".jpg": "이미지", ".jpeg": "이미지", ".tif": "이미지", ".tiff": "이미지", ".heic": "이미지", ".png": "이미지",
    ".mp4": "동영상", ".mov": "동영상", ".m4v": "동영상", ".avi": "동영상", ".mts": "동영상", ".3gp": "동영상",
    ".wav": "사운드", ".mp3": "사운드", ".m4a": "사운드", ".aac": "사운드", ".flac": "사운드",
}

AUTO_FIELDS = {"사업별 순번", "NO.", "미디어코드", "사업년도", "순번", "기관코드", "데이터베이스번호"}
DEFAULT_ORGANIZATION_NAME = "국립중앙과학관"

SPECIMEN_DRAFT_FILE = DRAFT_DIR / "new_specimen_draft.json"

# 표본은 관찰과 달리 미디어타입 구분 없이 이미지 등록만 지원하고,
# 데이터베이스번호도 미디어코드 없이 "SP" 고정값을 사용 (기관코드-SP-분류코드-연도순번No)
SPECIMEN_ROW_COLUMNS = [
    "NO.",
    "종고유ID",
    "원본경로",
    "원본파일명",
    "파일경로",
    "파일명",
    "분류코드",
    "사업명",
    "사업년도",
    "순번",
    "사업별 순번",
    "기관명",
    "기관코드",
    "국명",
    "학명",
    "과명",
    "속명",
    "종명",
    "이전목록번호 국문명",
    "이전목록번호 영문명",
    "데이터베이스번호",
    "이전데이터베이스번호",
    "동정자 국문",
    "동정자 영문",
    "동정년월일",
    "성별",
    "연령",
    "원본여부",
    "표본 공개여부",
    "대륙/해양 국문",
    "대륙/해양 영문",
    "위도",
    "경도",
    "국가국문명",
    "국가영문명",
    "채집위치 시도 국문",
    "채집위치 시도 영문",
    "채집위치 시군구 국문",
    "채집위치 시군구 영문",
    "채집위치 읍면동 국문",
    "채집위치 읍면동 영문",
    "채집위치 상세 국문",
    "채집위치 상세 영문",
    "채집자 국문",
    "채집자 영문",
    "채집년월일",
    "채집시간",
    "표본제작형태",
    "원산지",
    "원산지 공개여부",
    "자원공개여부",
    "국외반출여부",
    "분양가능여부",
    "표본활용현황",
]

# ponytail: 채집시간/표본제작형태 옵션 목록은 사용자가 준 이미지에서 "오전,오후,..."/
# "액침,건조,..."로 잘려있어 임시 목록으로 채움. 콤보박스를 editable로 열어뒀으니
# 실제 운영 목록이 정해지면 여기 리스트만 갱신하면 됨.
SPECIMEN_DROPDOWN_OPTIONS = {
    "분류코드": ["IN", "PL", "BI", "PI", "MM", "SP", "MR", "AR", "AM", "SW", "AG"],
    "성별": ["암컷", "수컷", "양성", "확인불가"],
    "연령": ["성체", "어린개체", "중간개체", "확인불가"],
    "원본여부": ["원본", "사본"],
    "표본 공개여부": ["공개", "비공개"],
    "대륙/해양 국문": ["아시아", "태평양"],
    "대륙/해양 영문": ["Asia", "Pacific Ocean"],
    "채집시간": ["오전", "오후", "새벽", "야간", "확인불가"],
    "표본제작형태": ["액침", "건조", "박제", "골격", "기타"],
    "원산지 공개여부": ["공개", "비공개"],
    "자원공개여부": ["공개", "비공개"],
    "국외반출여부": ["가능", "미반출"],
    "분양가능여부": ["가능", "불가능"],
}

SPECIMEN_DROPDOWN_DEFAULTS = {
    "성별": "확인불가",
    "연령": "확인불가",
    "원본여부": "원본",
    "표본 공개여부": "공개",
    "원산지 공개여부": "공개",
    "자원공개여부": "공개",
    "국외반출여부": "미반출",
}

SPECIMEN_STICKY_FIELDS = {"사업명", "사업년도", "순번", "기관명", "기관코드", "분류코드"}

SPECIMEN_ADDRESS_FIELDS = {
    "채집위치 시도 국문": ("채집위치 시도 영문", "sido"),
    "채집위치 시군구 국문": ("채집위치 시군구 영문", "sigungu"),
    "채집위치 읍면동 국문": ("채집위치 읍면동 영문", "eupmyeondong"),
    "채집위치 상세 국문": ("채집위치 상세 영문", "detail"),
}

SPECIMEN_AUTO_FIELDS = {"사업별 순번", "NO.", "사업년도", "순번", "기관코드", "데이터베이스번호"}

# 과명/속명/종명은 국명 조회 시 자동으로 채워지지만, 종정보 마스터에 과명 컬럼이
# 없어 비워둔 채로 시작함 -> 사용자가 직접 채우거나 자동값을 고치면 노란색으로 표시해서
# "마스터와 다르게 수동 보정된 값"임을 추후 업데이트 때 구분할 수 있게 함.
SPECIMEN_TAXON_AUTOFILL_FIELDS = ["과명", "속명", "종명"]
SPECIMEN_TAXON_EDITED_COLOR = QColor("#fff3b0")

SPECIMEN_FINAL_EXPORT_COLUMNS: list[tuple[str, str | None]] = [
    ("표본정보순번", None),
    ("종고유ID", "종고유ID"),
    ("학명", "학명"),
    ("과명", "과명"),
    ("속명", "속명"),
    ("종명", "종명"),
    ("기관코드", "기관코드"),
    ("기관명", "기관명"),
    ("동정자 국문", "동정자 국문"),
    ("동정자 영문", "동정자 영문"),
    ("동정년월일", "동정년월일"),
    ("대륙해양국문명", "대륙/해양 국문"),
    ("대륙해양영문명", "대륙/해양 영문"),
    ("국가 국문명", "국가국문명"),
    ("국가 영문명", "국가영문명"),
    ("경도", "경도"),
    ("위도", "위도"),
    ("성별", "성별"),
    ("연령", "연령"),
    ("이전목록번호국문명", "이전목록번호 국문명"),
    ("이전목록번호영문명", "이전목록번호 영문명"),
    ("데이터베이스번호", "데이터베이스번호"),
    ("이전데이터베이스번호", "이전데이터베이스번호"),
    ("원본여부", "원본여부"),
    ("표본정보 공개여부", "표본 공개여부"),
    ("채집위치 시도 국문", "채집위치 시도 국문"),
    ("채집위치 시도 영문", "채집위치 시도 영문"),
    ("채집위치 시군구 국문", "채집위치 시군구 국문"),
    ("채집위치 시군구 영문", "채집위치 시군구 영문"),
    ("채집위치읍면동국문명", "채집위치 읍면동 국문"),
    ("채집위치읍면동영문명", "채집위치 읍면동 영문"),
    ("채집위치 상세 국문", "채집위치 상세 국문"),
    ("채집위치 상세 영문", "채집위치 상세 영문"),
    ("채집자 국문", "채집자 국문"),
    ("채집자 영문", "채집자 영문"),
    ("채집년월일", "채집년월일"),
    ("채집시간", "채집시간"),
    ("사업년도", "사업년도"),
    ("순번", "순번"),
    ("사업명", "사업명"),
    ("원산지", "원산지"),
    ("원산지 공개여부", "원산지 공개여부"),
    ("자원공개여부", "자원공개여부"),
    ("국외반출여부", "국외반출여부"),
    ("분양가능여부", "분양가능여부"),
    ("표본제작형태", "표본제작형태"),
    ("표본활용현황", "표본활용현황"),
    ("관리번호", None),
]

DEFAULT_PROJECTS = [
    {"사업년도": 2025, "순번": 1, "사업명": "국가생물다양성기관연합공동학술조사"},
    {"사업년도": 2025, "순번": 2, "사업명": "생물다양성 소재 확보ㆍ보급을 위한 국내외 네트워크 운영"},
    {"사업년도": 2025, "순번": 3, "사업명": "기본연구과제 국내"},
    {"사업년도": 2025, "순번": 4, "사업명": "기본연구과제 국외"},
    {"사업년도": 2025, "순번": 5, "사업명": "2단계 2차 시민참여 생물다양성정보 확보"},
]

DEFAULT_ORGANIZATIONS = [
    {"기관명": "DMZ생태연구소", "기관코드": "DMZC"},
    {"기관명": "경기도수산해양자원연구소", "기관코드": "GFRI"},
    {"기관명": "경남산림환경연구원", "기관코드": "GFEI"},
    {"기관명": "경북대학교자연사박물관", "기관코드": "KNUN"},
    {"기관명": "경상북도환경연수원", "기관코드": "GDET"},
    {"기관명": "경희대학교자연사박물관", "기관코드": "NHMK"},
    {"기관명": "계룡산자연사박물관", "기관코드": "GNHM"},
    {"기관명": "국립과천과학관", "기관코드": "GCSM"},
    {"기관명": "국립대구과학관", "기관코드": "DNSM"},
    {"기관명": "국립문화재연구소 천연기념물센터", "기관코드": "NHCA"},
    {"기관명": "국립산림과학원난대ㆍ아열대산림연구소", "기관코드": "NIFS"},
    {"기관명": "국립생태원", "기관코드": "NIE"},
    {"기관명": "국립수목원", "기관코드": "KNAM"},
    {"기관명": "국립수산과학원중앙내수면연구소", "기관코드": "NFRD"},
    {"기관명": "국립일제특작과학원", "기관코드": "NIHH"},
    {"기관명": "국립중앙과학관", "기관코드": "NSMK"},
    {"기관명": "국립창원대학교", "기관코드": "CWNU"},
    {"기관명": "국립해양생물자원관", "기관코드": "NMBM"},
    {"기관명": "군산철새연구소", "기관코드": "GMBO"},
    {"기관명": "금강철새조망대(환경정책과)", "기관코드": "GBWT"},
    {"기관명": "농림축산검역검사본부(식물검역부)", "기관코드": "FLQI"},
    {"기관명": "다살이생물자원연구소", "기관코드": "DRIB"},
    {"기관명": "대구광역시 수목원관리사무소", "기관코드": "DAMO"},
    {"기관명": "대전광역시 한밭수목원", "기관코드": "HBBG"},
    {"기관명": "목포자연사박물관", "기관코드": "MNHM"},
    {"기관명": "몽골자연사박물관", "기관코드": "MMNH"},
    {"기관명": "무주곤충박물관", "기관코드": "MJIN"},
    {"기관명": "무지개세상", "기관코드": "ECRB"},
    {"기관명": "문경자연생태박물관", "기관코드": "MNEM"},
    {"기관명": "별새꽃돌자연탐사과학관", "기관코드": "NMNO"},
    {"기관명": "부산해양자연사박물관", "기관코드": "BMNM"},
    {"기관명": "서대문자연사박물관", "기관코드": "SMNH"},
    {"기관명": "서울시립과학관", "기관코드": "SCSM"},
    {"기관명": "성신여자대학교자연사박물관", "기관코드": "SSNH"},
    {"기관명": "안면도쥬라기박물관", "기관코드": "ANJM"},
    {"기관명": "양평곤충박물관", "기관코드": "YPIM"},
    {"기관명": "우석헌자연사박물관", "기관코드": "WSHN"},
    {"기관명": "우포늪생태관", "기관코드": "UWEP"},
    {"기관명": "이화여자대학교자연사박물관", "기관코드": "ENHM"},
    {"기관명": "인제곤충바이오센터", "기관코드": "IJIB"},
    {"기관명": "전남해양수산과학관", "기관코드": "ASJK"},
    {"기관명": "전라남도완도수목원", "기관코드": "GNWP"},
    {"기관명": "전북산림환경연구소대아수목원", "기관코드": "JBFE"},
    {"기관명": "제주자치시도세계유산본부(한라수목원)", "기관코드": "JJHA"},
    {"기관명": "제주테크노파크 생물종다양성연구소", "기관코드": "JBRI"},
    {"기관명": "제주특별자치도 민속자연사박물관", "기관코드": "JNHM"},
    {"기관명": "(주)청록환경생태연구소", "기관코드": "CEER"},
    {"기관명": "창녕군우포늪따오기사업소", "기관코드": "CNWD"},
    {"기관명": "충남대학교", "기관코드": "CNU"},
    {"기관명": "충남대학교자연사박물관", "기관코드": "NHMC"},
    {"기관명": "충주자연생태체험관", "기관코드": "CNEC"},
    {"기관명": "충청남도 산림자원연구소", "기관코드": "CNFR"},
    {"기관명": "태백고생대자연사박물관", "기관코드": "TPNM"},
    {"기관명": "한국과학기술정보연구원", "기관코드": "KIDR"},
    {"기관명": "한국동굴생물연구소", "기관코드": "KIBS"},
    {"기관명": "한국물새네트워크", "기관코드": "KRWN"},
    {"기관명": "한국생명공학연구원 국가생명연구자원정보센터", "기관코드": "KRIB"},
    {"기관명": "한국생명공학연구원 생물자원센터", "기관코드": "KRBC"},
    {"기관명": "한국수목원정원관리원", "기관코드": "KAGI"},
    {"기관명": "한국수자원공사K-water연구원", "기관코드": "KIWE"},
    {"기관명": "한국자생식물원", "기관코드": "KOBG"},
    {"기관명": "한국환경생태연구소", "기관코드": "KIEE"},
    {"기관명": "한남대학교 자연사박물관", "기관코드": "HUNM"},
    {"기관명": "한화아쿠아리움", "기관코드": "HHNR"},
    {"기관명": "해외생물소재센터", "기관코드": "IBMRC"},
    {"기관명": "홍성조류탐사과학관", "기관코드": "HSBS"},
    {"기관명": "한국하의생연구원", "기관코드": "KIOM"},
    {"기관명": "국립호남권생물자원관", "기관코드": "HIBR"},
    {"기관명": "국립생물자원관", "기관코드": "NIBR"},
    {"기관명": "국립낙동강생물자원관", "기관코드": "NDBR"},
]


FILL_HANDLE_SIZE = 6


class ThumbnailLabel(QLabel):
    """미리보기 썸네일. 이미지가 표시돼 있을 때 클릭하면 팝업으로 확대해서 보여줌."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._source_path = ""

    def set_source_path(self, path: str):
        self._source_path = path
        self.setCursor(Qt.PointingHandCursor if path else Qt.ArrowCursor)
        self.setToolTip("클릭하면 크게 보기" if path else "")

    def mousePressEvent(self, event):
        if self._source_path:
            self._show_popup()
            return
        super().mousePressEvent(event)

    def _show_popup(self):
        pixmap = QPixmap(self._source_path)
        if pixmap.isNull():
            return
        ImagePreviewDialog(pixmap, Path(self._source_path).name, self).exec()


class ImagePreviewDialog(QDialog):
    """확대/축소(휠, 버튼)와 원본 100% 보기, 화면에 맞춤을 지원하는 이미지 미리보기 팝업."""

    ZOOM_MIN = 0.1
    ZOOM_MAX = 5.0
    ZOOM_STEP = 1.25

    def __init__(self, pixmap: QPixmap, title: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self._pixmap = pixmap
        self._zoom = 1.0

        layout = QVBoxLayout(self)

        controls = QHBoxLayout()
        zoom_out_button = QPushButton("축소(-)")
        zoom_out_button.clicked.connect(lambda: self.set_zoom(self._zoom / self.ZOOM_STEP))
        zoom_in_button = QPushButton("확대(+)")
        zoom_in_button.clicked.connect(lambda: self.set_zoom(self._zoom * self.ZOOM_STEP))
        actual_size_button = QPushButton("원본 크기(100%)")
        actual_size_button.clicked.connect(lambda: self.set_zoom(1.0))
        fit_button = QPushButton("화면에 맞춤")
        fit_button.clicked.connect(self.fit_to_view)
        self.zoom_status_label = QLabel("")
        controls.addWidget(zoom_out_button)
        controls.addWidget(zoom_in_button)
        controls.addWidget(actual_size_button)
        controls.addWidget(fit_button)
        controls.addWidget(self.zoom_status_label)
        controls.addStretch()
        layout.addLayout(controls)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(False)
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.scroll_area.setWidget(self.image_label)
        layout.addWidget(self.scroll_area, 1)

        screen = QApplication.primaryScreen()
        window_size = (screen.availableSize() * 0.8) if screen else QSize(1000, 800)
        self.resize(window_size)

        self.fit_to_view()

    def set_zoom(self, zoom: float):
        self._zoom = max(self.ZOOM_MIN, min(zoom, self.ZOOM_MAX))
        scaled = self._pixmap.scaled(
            self._pixmap.size() * self._zoom, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        self.image_label.setPixmap(scaled)
        self.image_label.resize(scaled.size())
        self.zoom_status_label.setText(f"{round(self._zoom * 100)}%")

    def fit_to_view(self):
        viewport = self.scroll_area.viewport().size()
        if self._pixmap.width() <= 0 or self._pixmap.height() <= 0:
            return
        zoom = min(viewport.width() / self._pixmap.width(), viewport.height() / self._pixmap.height(), 1.0)
        self.set_zoom(zoom)

    def wheelEvent(self, event):
        self.set_zoom(self._zoom * (self.ZOOM_STEP if event.angleDelta().y() > 0 else 1 / self.ZOOM_STEP))


class DataTableWidget(QTableWidget):
    """Ctrl+C/Ctrl+V로 엑셀과 여러 셀을 탭/줄바꿈 텍스트로 주고받거나, 선택 영역 우측 하단
    모서리를 드래그해서 값을 위/아래로 복사(엑셀 채우기 핸들과 비슷하되 단순 복사만 지원)."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.paste_requested = None
        self._fill_dragging = False
        self._fill_source_rows: list[int] = []
        self._fill_columns: list[int] = []
        self._fill_target_row: int | None = None

    def keyPressEvent(self, event):
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_V and self.paste_requested:
            self.paste_requested()
            return
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_C:
            self._copy_selection_to_clipboard()
            return
        super().keyPressEvent(event)

    def _copy_selection_to_clipboard(self):
        selected = self.selectedIndexes()
        if not selected:
            return
        rows = sorted({index.row() for index in selected})
        cols = sorted({index.column() for index in selected})
        lines = []
        for row in rows:
            cells = []
            for col in cols:
                item = self.item(row, col)
                cells.append(item.text() if item else "")
            lines.append("\t".join(cells))
        # 끝에 줄바꿈을 하나 더 붙여둠 (엑셀 복사 형식과 동일) -> 안 붙이면 빈 셀 1개만 복사했을 때
        # 클립보드 텍스트가 완전히 빈 문자열이 돼서, 붙여넣기 쪽의 "빈 줄 제거" 로직이 이걸 통째로
        # "붙여넣을 내용 없음"으로 오인해 Ctrl+V가 조용히 아무 동작도 안 하던 문제가 있었음
        QApplication.clipboard().setText("\n".join(lines) + "\n")

    def _fill_handle_rect(self) -> QRect | None:
        selected = self.selectedIndexes()
        if not selected:
            return None
        bottom_right = max(selected, key=lambda idx: (idx.row(), idx.column()))
        cell_rect = self.visualRect(bottom_right)
        return QRect(
            cell_rect.right() - FILL_HANDLE_SIZE,
            cell_rect.bottom() - FILL_HANDLE_SIZE,
            FILL_HANDLE_SIZE * 2,
            FILL_HANDLE_SIZE * 2,
        )

    def mousePressEvent(self, event):
        handle_rect = self._fill_handle_rect()
        if handle_rect and handle_rect.contains(event.position().toPoint()):
            selected = self.selectedIndexes()
            self._fill_source_rows = sorted({idx.row() for idx in selected})
            self._fill_columns = sorted({idx.column() for idx in selected})
            self._fill_target_row = self._fill_source_rows[-1] if self._fill_source_rows else None
            self._fill_dragging = bool(self._fill_source_rows)
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._fill_dragging:
            row = self.rowAt(event.position().toPoint().y())
            if row == -1:
                row = self.rowCount() - 1
            if row != self._fill_target_row:
                self._fill_target_row = row
                self._preview_fill_range(row)
            return

        handle_rect = self._fill_handle_rect()
        is_over_handle = bool(handle_rect and handle_rect.contains(event.position().toPoint()))
        self.viewport().setCursor(QCursor(Qt.CrossCursor) if is_over_handle else QCursor(Qt.ArrowCursor))
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._fill_dragging:
            self._commit_fill()
            self._fill_dragging = False
            self._fill_source_rows = []
            self._fill_columns = []
            self._fill_target_row = None
            return
        super().mouseReleaseEvent(event)

    def _preview_fill_range(self, target_row: int):
        if not self._fill_source_rows or not self._fill_columns:
            return
        top = min(self._fill_source_rows[0], target_row)
        bottom = max(self._fill_source_rows[-1], target_row)
        # 드래그 중 미리보기는 선택영역만 넓혀서 보여줌(값은 놓을 때 한번에 적용) ->
        # itemSelectionChanged를 계속 쏟아내며 폼을 흔들지 않도록 신호는 막아둠
        was_blocked = self.blockSignals(True)
        try:
            self.clearSelection()
            self.setRangeSelected(
                QTableWidgetSelectionRange(top, self._fill_columns[0], bottom, self._fill_columns[-1]), True
            )
        finally:
            self.blockSignals(was_blocked)

    def _commit_fill(self):
        if self._fill_target_row is None or not self._fill_source_rows or not self._fill_columns:
            return
        source_min, source_max = self._fill_source_rows[0], self._fill_source_rows[-1]
        target_row = self._fill_target_row
        if target_row > source_max:
            fill_rows = range(source_max + 1, target_row + 1)
        elif target_row < source_min:
            fill_rows = range(target_row, source_min)
        else:
            return

        # 값 쓰기는 신호를 막지 않음 -> 국명 자동조회/날짜 형식변환 등 기존 셀 편집 로직이
        # itemChanged를 통해 채워진 행에도 그대로 적용됨
        for offset, row in enumerate(fill_rows):
            source_row = self._fill_source_rows[offset % len(self._fill_source_rows)]
            for col in self._fill_columns:
                source_item = self.item(source_row, col)
                text = source_item.text() if source_item else ""
                target_item = self.item(row, col)
                if target_item is None:
                    self.setItem(row, col, QTableWidgetItem(text))
                elif target_item.text() != text:
                    target_item.setText(text)


class UploadTab(QWidget):
    def __init__(self):
        super().__init__()
        self.projects: list[dict[str, Any]] = []
        self.organizations: list[dict[str, Any]] = []
        self.project_lookup: dict[str, dict[str, Any]] = {}
        self.organization_lookup: dict[str, dict[str, Any]] = {}
        self.form_widgets: dict[str, QLineEdit | QComboBox | QSpinBox] = {}
        self.specimen_form_widgets: dict[str, QLineEdit | QComboBox | QSpinBox] = {}
        self.species_master_df = pd.DataFrame()
        self.species_name_lookup: dict[str, list[dict[str, Any]]] = {}
        self.master_service = SpeciesMasterService()
        self.address_master = AddressMasterService()
        self._last_selected_row: int | None = None
        self._specimen_last_selected_row: int | None = None
        self._specimen_taxon_edited_fields: set[str] = set()
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setSingleShot(True)
        self.autosave_timer.timeout.connect(self.autosave_draft)
        self.autosave_suspended = False
        self.specimen_autosave_timer = QTimer(self)
        self.specimen_autosave_timer.setSingleShot(True)
        self.specimen_autosave_timer.timeout.connect(self.autosave_specimen_draft)
        self.specimen_autosave_suspended = False
        self.media_storage_dir = ""
        self.specimen_media_storage_dir = ""

        self._ensure_reference_files()
        self.load_reference_data()
        self.load_species_master_lookup()
        self.init_ui()
        self.load_draft()
        self.load_specimen_draft()

    def init_ui(self):
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        title_label = QLabel("신규 데이터 직접 작성")
        main_layout.addWidget(title_label)
        self.preserve_database_numbers = QCheckBox("기존 데이터베이스번호 유지 (해제하면 새 번호 생성)")
        self.preserve_database_numbers.setChecked(True)
        main_layout.addWidget(self.preserve_database_numbers)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.tabs.addTab(self._build_row_editor_tab(), "관찰 정보작성")
        self.tabs.addTab(self._build_specimen_editor_tab(), "표본정보 작성")
        self.tabs.addTab(self._build_project_tab(), "사업 관리")
        self.tabs.addTab(self._build_organization_tab(), "기관 관리")

    def _build_row_editor_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        control_layout = QHBoxLayout()

        new_form_button = QPushButton("새 입력")
        new_form_button.clicked.connect(self.clear_form)
        import_folder_button = QPushButton("미디어 폴더 불러오기")
        import_folder_button.clicked.connect(self.import_media_folder)
        update_row_button = QPushButton("선택 행 수정")
        update_row_button.clicked.connect(self.update_selected_data_row)
        delete_row_button = QPushButton("선택 행 삭제")
        delete_row_button.clicked.connect(self.delete_selected_data_rows)
        lookup_species_button = QPushButton("국명·학명 조회")
        lookup_species_button.clicked.connect(self.fill_species_from_korean_name)
        lookup_address_button = QPushButton("위도/경도로 주소 조회")
        lookup_address_button.clicked.connect(self.fill_address_from_latlon)
        map_button = QPushButton("이 행 위치 확인")
        map_button.clicked.connect(self.open_location_on_map)
        self.species_status_label = QLabel("")
        self.configure_status_label(self.species_status_label)
        self.species_status_label.hide()

        control_layout.addWidget(new_form_button)
        control_layout.addWidget(import_folder_button)
        control_layout.addWidget(update_row_button)
        control_layout.addWidget(delete_row_button)
        control_layout.addWidget(lookup_species_button)
        control_layout.addWidget(lookup_address_button)
        control_layout.addWidget(map_button)
        control_layout.addStretch()
        layout.addLayout(control_layout)
        layout.addWidget(self.species_status_label)

        media_storage_layout = QHBoxLayout()
        self.media_storage_path_input = QLineEdit()
        self.media_storage_path_input.setReadOnly(True)
        self.media_storage_path_input.setPlaceholderText("미디어 파일을 DB번호 이름으로 저장할 폴더")
        media_storage_button = QPushButton("이미지·동영상·사운드 저장 위치 선택")
        media_storage_button.clicked.connect(self.select_media_storage_dir)
        media_storage_layout.addWidget(QLabel("미디어 저장 위치"))
        media_storage_layout.addWidget(self.media_storage_path_input, 1)
        media_storage_layout.addWidget(media_storage_button)
        layout.addLayout(media_storage_layout)
        layout.addWidget(QLabel("엑셀 저장 시 미디어가 있으면 저장 폴더를 다시 선택합니다. 동영상·캡처만 종고유ID 하위 폴더에 저장합니다."))

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        form_container = QWidget()
        self.row_form = QFormLayout(form_container)
        self.row_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self.build_row_form()
        scroll_area.setWidget(form_container)

        next_row_layout = QHBoxLayout()
        next_row_layout.addStretch()
        add_row_button = QPushButton("입력")
        add_row_button.clicked.connect(self.add_data_row_from_form)
        next_row_button = QPushButton("다음 행")
        next_row_button.clicked.connect(self.go_to_next_row)
        next_row_layout.addWidget(add_row_button)
        next_row_layout.addWidget(next_row_button)

        # 입력창(위)/표(아래) 경계를 사용자가 드래그해서 조절할 수 있게 QSplitter로 묶음
        # (기존엔 2:3 고정 비율이라 입력창을 더 길게 볼 수 없었음)
        top_pane = QWidget()
        top_pane_layout = QVBoxLayout(top_pane)
        top_pane_layout.setContentsMargins(0, 0, 0, 0)
        top_pane_layout.addWidget(scroll_area, 1)
        top_pane_layout.addLayout(next_row_layout)

        self.data_table = DataTableWidget(0, len(ROW_COLUMNS))
        self.data_table.setHorizontalHeaderLabels(ROW_COLUMNS)
        # SelectRows였으면 클릭 시 항상 행 전체가 선택돼 세로(한 컬럼) 범위를 고를 수 없었음 ->
        # Ctrl+C 복사/채우기 핸들이 가로(행)든 세로(컬럼)든 선택한 그대로 동작하도록 셀 단위로 변경
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.data_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed
            | QAbstractItemView.SelectedClicked
        )
        self.data_table.itemSelectionChanged.connect(self.load_selected_row_to_form)
        self.data_table.itemChanged.connect(self.on_data_table_item_changed)
        self.data_table.paste_requested = self.paste_clipboard_into_table
        # ponytail: ResizeToContents는 셀 값이 바뀔 때마다 해당 컬럼 전체(수천 행)의
        # 텍스트 폭을 다시 재는 탓에 행이 많아지면 한 번 쓰기에 10초 이상 걸림.
        # Interactive로 바꾸고 최초 1회만 폭을 맞춤 (이후 사용자가 드래그로 조절 가능).
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.data_table.resizeColumnsToContents()

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(top_pane)
        splitter.addWidget(self.data_table)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)
        layout.addWidget(splitter, 1)

        save_layout = QHBoxLayout()
        save_draft_button = QPushButton("작성 내용 저장")
        save_draft_button.clicked.connect(self.save_draft)
        load_draft_button = QPushButton("임시저장 불러오기")
        load_draft_button.clicked.connect(self.reload_draft)
        export_button = QPushButton("엑셀로 저장")
        export_button.clicked.connect(self.export_rows_to_excel)
        import_button = QPushButton("엑셀 불러오기")
        import_button.clicked.connect(self.load_rows_from_excel)
        thumbnail_button = QPushButton("동영상 미리보기 생성")
        thumbnail_button.clicked.connect(self.generate_video_thumbnails)
        address_all_button = QPushButton("위치 전체 주소 조회")
        address_all_button.clicked.connect(self.fill_all_addresses_from_latlon)
        species_project_all_button = QPushButton("종ID·사업·미디어 일괄 적용")
        species_project_all_button.clicked.connect(self.fill_all_species_and_project)
        map_all_button = QPushButton("전체 위치 지도 확인")
        map_all_button.clicked.connect(self.open_all_locations_on_map)
        validate_button = self.create_validation_button(self.data_table, ROW_COLUMNS)
        clear_button = QPushButton("작성 내용 비우기")
        clear_button.clicked.connect(self.clear_data_rows)
        self.save_status_label = QLabel("")
        self.configure_status_label(self.save_status_label)

        save_layout.addStretch()
        save_layout.addWidget(save_draft_button)
        save_layout.addWidget(load_draft_button)
        save_layout.addWidget(export_button)
        save_layout.addWidget(import_button)
        save_layout.addWidget(thumbnail_button)
        save_layout.addWidget(species_project_all_button)
        save_layout.addWidget(address_all_button)
        save_layout.addWidget(validate_button)
        save_layout.addWidget(map_all_button)
        save_layout.addWidget(clear_button)
        layout.addLayout(save_layout)
        layout.addWidget(self.save_status_label)
        self.refresh_combos()
        return tab

    @staticmethod
    def configure_status_label(label: QLabel):
        # Long paths/status messages must not increase the main window's minimum width.
        label.setWordWrap(True)
        label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        label.setFixedHeight(40)
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

    def build_row_form(self):
        for column_name in ROW_COLUMNS:
            label = QLabel(column_name)
            if column_name in {"사업명", "기관명"}:
                widget = QComboBox()
                widget.currentTextChanged.connect(self.refresh_generated_fields)
                if column_name == "사업명":
                    widget.currentTextChanged.connect(self.apply_project_to_form)
                else:
                    widget.currentTextChanged.connect(self.apply_organization_to_form)
            elif column_name == "사업년도":
                widget = QSpinBox()
                widget.setRange(2000, 2100)
                widget.setValue(datetime.now().year)
                widget.valueChanged.connect(self.refresh_generated_fields)
            elif column_name in {"NO.", "순번", "관찰개체수"}:
                widget = QSpinBox()
                widget.setRange(0, 99999)
                widget.setValue(1 if column_name in {"NO.", "관찰개체수"} else 0)
                widget.valueChanged.connect(self.refresh_generated_fields)
            elif column_name in DROPDOWN_OPTIONS:
                widget = QComboBox()
                widget.addItems(DROPDOWN_OPTIONS[column_name])
                if column_name in DROPDOWN_DEFAULTS:
                    widget.setCurrentText(DROPDOWN_DEFAULTS[column_name])
                if column_name in {"대륙/해양 국문", "대륙/해양 영문", "분류코드"}:
                    widget.setEditable(True)
                widget.currentTextChanged.connect(self.refresh_generated_fields)
                if column_name == "미디어타입":
                    widget.currentTextChanged.connect(self.update_media_type_field_visibility)
            else:
                widget = QLineEdit()
                if column_name in DATE_FIELDS:
                    widget.setPlaceholderText("2016-01-01 (260508, 20260508도 자동 변환)")
                    widget.editingFinished.connect(lambda field=column_name: self.normalize_date_field(field))
                elif column_name == "위도":
                    widget.setPlaceholderText("36.123456")
                elif column_name == "경도":
                    widget.setPlaceholderText("126.123456")
                elif column_name == "종고유ID":
                    widget.setPlaceholderText("종정보 마스터의 종고유ID")
                elif column_name in {"원본경로", "원본파일명", "파일경로", "파일명"}:
                    widget.setReadOnly(True)
                    if column_name == "파일경로":
                        widget.setPlaceholderText("원본 미디어 선택 버튼으로 선택")
                    elif column_name == "원본경로":
                        widget.setPlaceholderText("저장 전 원본 파일 경로(자동 기록, 이후 저장해도 유지)")
                elif column_name in {"동영상 재생시간", "사운드 재생시간"}:
                    widget.setPlaceholderText("예: 0:45")
                elif column_name in {"동영상 파일명", "사운드 파일명"}:
                    widget.setPlaceholderText("미디어 저장 시 자동 입력, 직접 수정 가능")
                elif column_name == "국명":
                    widget.setPlaceholderText("입력 후 Enter 또는 국명 조회")
                    widget.editingFinished.connect(self.fill_species_from_korean_name)
                elif column_name == "학명":
                    widget.setPlaceholderText("학명 입력 후 Enter 또는 국명·학명 조회")
                    widget.editingFinished.connect(lambda: self.fill_species_from_korean_name(prefer_scientific=True))
                elif column_name in ADDRESS_FIELDS:
                    if column_name == "관찰위치 상세 국문":
                        widget.setPlaceholderText("직접 입력 (자동완성 후보 제공, 영문 자동 매칭)")
                        # ponytail: 12,000여개 후보라 정렬+시작문자 매칭(이진탐색)으로 붙여야
                        # 타이핑할 때마다 전체를 훑지 않음. 부분일치가 필요해지면 완성도 낮추고
                        # 후보군 축소(예: 지역 선필터) 후 MatchContains로 되돌리는 걸 고려.
                        completer = QCompleter(sorted(self.address_master.candidates("detail")), widget)
                        completer.setCaseSensitivity(Qt.CaseInsensitive)
                        completer.setModelSorting(QCompleter.CaseInsensitivelySortedModel)
                        completer.setFilterMode(Qt.MatchStartsWith)
                        widget.setCompleter(completer)
                    else:
                        widget.setPlaceholderText("국문 입력 후 Enter (영문 자동 매칭)")
                    widget.editingFinished.connect(
                        lambda field=column_name: self.fill_address_english(field)
                    )
            if column_name in TEXT_DEFAULTS and isinstance(widget, QLineEdit):
                widget.setText(TEXT_DEFAULTS[column_name])
            if column_name == "데이터베이스번호":
                widget.editingFinished.connect(lambda: self.edit_database_number(False))
            if column_name == "사업별 순번":
                widget.textEdited.connect(lambda: self.refresh_generated_fields(force_number=True))
                widget.editingFinished.connect(lambda: self.commit_project_sequence(False))
            if column_name in AUTO_FIELDS and column_name not in {"데이터베이스번호", "사업별 순번"}:
                widget.setEnabled(False)
                if isinstance(widget, QLineEdit):
                    widget.setReadOnly(True)
                if isinstance(widget, QSpinBox):
                    widget.setButtonSymbols(QSpinBox.NoButtons)
            self.form_widgets[column_name] = widget
            self.connect_autosave_signal(widget)
            if column_name == "파일경로":
                # "파일명"은 별도 컬럼(read-only)이지만 "파일경로"가 바뀔 때마다
                # 자동으로 따라가도록 여기서 연결(sync_media_filename_field)
                widget.textChanged.connect(self.sync_media_filename_field)
                widget.textChanged.connect(self.update_media_thumbnail)
                media_layout = QHBoxLayout()
                media_layout.addWidget(widget, 1)
                select_media_button = QPushButton("원본 선택")
                select_media_button.clicked.connect(self.select_media_source_file)
                media_layout.addWidget(select_media_button)
                self.row_form.addRow(label, media_layout)

                self.media_thumbnail_label = ThumbnailLabel()
                self.media_thumbnail_label.setFixedSize(180, 180)
                self.media_thumbnail_label.setAlignment(Qt.AlignCenter)
                self.media_thumbnail_label.setStyleSheet("border: 1px solid #ccc; background: #fafafa;")
                self.row_form.addRow("미리보기", self.media_thumbnail_label)
            else:
                self.row_form.addRow(label, widget)
        self.update_media_type_field_visibility()

    def sync_media_filename_field(self, *_):
        value = self.get_form_value("파일경로")
        self.set_form_value("파일명", Path(value).name if value else "")

    def update_media_thumbnail(self, *_):
        self.render_media_thumbnail(self.media_thumbnail_label, self.get_form_value("파일경로"))

    @staticmethod
    def render_media_thumbnail(label: ThumbnailLabel, file_path_text: str):
        if not file_path_text:
            label.clear()
            label.set_source_path("")
            return

        path = Path(file_path_text)
        media_type = MEDIA_EXTENSION_TYPES.get(path.suffix.lower())
        image_path = (
            path if media_type == "이미지" else UploadTab.video_thumbnail_path(path) if media_type == "동영상" else None
        )

        if media_type == "동영상":
            preview_path = UploadTab.video_preview_thumbnail_path(path)
            if preview_path.is_file():
                image_path = preview_path
        pixmap = QPixmap(str(image_path)) if image_path and image_path.exists() else None
        if pixmap and not pixmap.isNull():
            label.setPixmap(pixmap.scaled(180, 180, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            label.set_source_path(str(image_path))
        elif media_type == "동영상":
            label.setText("🎬 동영상\n(미리보기 없음 -\n'동영상 미리보기 생성' 이용)")
            label.set_source_path("")
        elif media_type == "사운드":
            label.setText("🔊 사운드 파일")
            label.set_source_path("")
        else:
            label.setText("미리보기 불가")
            label.set_source_path("")

    @staticmethod
    def normalize_date_text(text: str) -> str:
        # 순수 숫자 6자리(YYMMDD)/8자리(YYYYMMDD)만 변환. 이미 구분자가 있는 값은
        # (형식이 맞든 틀리든) 잘못 재해석하지 않도록 그대로 둠.
        digits = text.strip()
        if not digits.isdigit():
            return text
        if len(digits) == 6:
            return f"20{digits[0:2]}-{digits[2:4]}-{digits[4:6]}"
        if len(digits) == 8:
            return f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]}"
        return text

    def normalize_date_field(self, field: str):
        self.set_form_value(field, self.normalize_date_text(self.get_form_value(field)))

    def update_media_type_field_visibility(self, *_):
        media_type = self.get_form_value("미디어타입")
        for column_name in VIDEO_FIELDS:
            self.set_field_row_visible(column_name, media_type == "동영상")
        for column_name in SOUND_FIELDS:
            self.set_field_row_visible(column_name, media_type == "사운드")

    def set_field_row_visible(self, column_name: str, visible: bool):
        widget = self.form_widgets.get(column_name)
        if widget is None:
            return
        self.row_form.setRowVisible(widget, visible)

    # ------------------------------------------------------------------
    # 표본정보 작성 (관찰 정보작성과 같은 구조를 재사용: 사업/기관 참조데이터, 종정보 마스터,
    # 주소 마스터, EXIF/ffprobe 미디어 추출 등은 self.<staticmethod>로 그대로 재사용하고,
    # 폼/표 상태만 별도(specimen_ 접두사)로 둠. ponytail: 관찰 쪽 CRUD와 구조가 거의
    # 동일해 다소 중복이지만, 컬럼/DB번호 규칙/이미지 단일등록 등 차이가 커 공통 베이스로
    # 추출하려면 관찰 코드 전체를 건드려야 해서 보류함 -> 다음에 손댈 일 있으면 공통 베이스
    # 클래스로 뽑아내는 걸 고려.
    # ------------------------------------------------------------------

    def _build_specimen_editor_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        control_layout = QHBoxLayout()

        new_form_button = QPushButton("새 입력")
        new_form_button.clicked.connect(self.clear_specimen_form)
        import_folder_button = QPushButton("이미지 폴더 불러오기")
        import_folder_button.clicked.connect(self.import_specimen_media_folder)
        update_row_button = QPushButton("선택 행 수정")
        update_row_button.clicked.connect(self.update_selected_specimen_data_row)
        delete_row_button = QPushButton("선택 행 삭제")
        delete_row_button.clicked.connect(self.delete_selected_specimen_data_rows)
        lookup_species_button = QPushButton("국명·학명 조회")
        lookup_species_button.clicked.connect(self.fill_specimen_taxon_from_korean_name)
        lookup_address_button = QPushButton("위도/경도로 주소 조회")
        lookup_address_button.clicked.connect(self.fill_specimen_address_from_latlon)
        map_button = QPushButton("이 행 위치 확인")
        map_button.clicked.connect(self.open_specimen_location_on_map)
        self.specimen_species_status_label = QLabel("")
        self.configure_status_label(self.specimen_species_status_label)
        self.specimen_species_status_label.hide()

        control_layout.addWidget(new_form_button)
        control_layout.addWidget(import_folder_button)
        control_layout.addWidget(update_row_button)
        control_layout.addWidget(delete_row_button)
        control_layout.addWidget(lookup_species_button)
        control_layout.addWidget(lookup_address_button)
        control_layout.addWidget(map_button)
        control_layout.addStretch()
        layout.addLayout(control_layout)
        layout.addWidget(self.specimen_species_status_label)

        media_storage_layout = QHBoxLayout()
        self.specimen_media_storage_path_input = QLineEdit()
        self.specimen_media_storage_path_input.setReadOnly(True)
        self.specimen_media_storage_path_input.setPlaceholderText("표본 이미지를 DB번호 이름으로 저장할 폴더")
        media_storage_button = QPushButton("이미지 저장 위치 선택")
        media_storage_button.clicked.connect(self.select_specimen_media_storage_dir)
        media_storage_layout.addWidget(QLabel("이미지 저장 위치"))
        media_storage_layout.addWidget(self.specimen_media_storage_path_input, 1)
        media_storage_layout.addWidget(media_storage_button)
        layout.addLayout(media_storage_layout)
        layout.addWidget(QLabel("엑셀 저장 시 이미지가 있으면 저장 폴더를 다시 선택합니다. 이미지는 DB번호 이름으로 복사하며 원본은 유지합니다."))

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        form_container = QWidget()
        self.specimen_row_form = QFormLayout(form_container)
        self.specimen_row_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self.build_specimen_row_form()
        scroll_area.setWidget(form_container)

        next_row_layout = QHBoxLayout()
        next_row_layout.addStretch()
        add_row_button = QPushButton("입력")
        add_row_button.clicked.connect(self.add_specimen_data_row_from_form)
        next_row_button = QPushButton("다음 행")
        next_row_button.clicked.connect(self.go_to_next_specimen_row)
        next_row_layout.addWidget(add_row_button)
        next_row_layout.addWidget(next_row_button)

        top_pane = QWidget()
        top_pane_layout = QVBoxLayout(top_pane)
        top_pane_layout.setContentsMargins(0, 0, 0, 0)
        top_pane_layout.addWidget(scroll_area, 1)
        top_pane_layout.addLayout(next_row_layout)

        self.specimen_data_table = DataTableWidget(0, len(SPECIMEN_ROW_COLUMNS))
        self.specimen_data_table.setHorizontalHeaderLabels(SPECIMEN_ROW_COLUMNS)
        self.specimen_data_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.specimen_data_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed
            | QAbstractItemView.SelectedClicked
        )
        self.specimen_data_table.itemSelectionChanged.connect(self.load_selected_specimen_row_to_form)
        self.specimen_data_table.itemChanged.connect(self.on_specimen_data_table_item_changed)
        self.specimen_data_table.paste_requested = self.paste_clipboard_into_specimen_table
        self.specimen_data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.specimen_data_table.resizeColumnsToContents()

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(top_pane)
        splitter.addWidget(self.specimen_data_table)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)
        layout.addWidget(splitter, 1)

        save_layout = QHBoxLayout()
        save_draft_button = QPushButton("작성 내용 저장")
        save_draft_button.clicked.connect(self.save_specimen_draft)
        load_draft_button = QPushButton("임시저장 불러오기")
        load_draft_button.clicked.connect(self.reload_specimen_draft)
        export_button = QPushButton("엑셀로 저장")
        export_button.clicked.connect(self.export_specimen_rows_to_excel)
        import_button = QPushButton("엑셀 불러오기")
        import_button.clicked.connect(self.load_specimen_rows_from_excel)
        address_all_button = QPushButton("위치 전체 주소 조회")
        address_all_button.clicked.connect(self.fill_all_specimen_addresses_from_latlon)
        species_project_all_button = QPushButton("종ID·사업 일괄 적용")
        species_project_all_button.clicked.connect(self.fill_all_specimen_species_and_project)
        map_all_button = QPushButton("전체 위치 지도 확인")
        map_all_button.clicked.connect(self.open_all_specimen_locations_on_map)
        validate_button = self.create_validation_button(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
        clear_button = QPushButton("작성 내용 비우기")
        clear_button.clicked.connect(self.clear_specimen_data_rows)
        self.specimen_save_status_label = QLabel("")
        self.configure_status_label(self.specimen_save_status_label)

        save_layout.addStretch()
        save_layout.addWidget(save_draft_button)
        save_layout.addWidget(load_draft_button)
        save_layout.addWidget(export_button)
        save_layout.addWidget(import_button)
        save_layout.addWidget(species_project_all_button)
        save_layout.addWidget(address_all_button)
        save_layout.addWidget(validate_button)
        save_layout.addWidget(map_all_button)
        save_layout.addWidget(clear_button)
        layout.addLayout(save_layout)
        layout.addWidget(self.specimen_save_status_label)
        self.refresh_combos()
        return tab

    def build_specimen_row_form(self):
        for column_name in SPECIMEN_ROW_COLUMNS:
            label = QLabel(column_name)
            if column_name in {"사업명", "기관명"}:
                widget = QComboBox()
                widget.currentTextChanged.connect(self.refresh_specimen_generated_fields)
                if column_name == "사업명":
                    widget.currentTextChanged.connect(self.apply_specimen_project_to_form)
                else:
                    widget.currentTextChanged.connect(self.apply_specimen_organization_to_form)
            elif column_name in {"사업년도", "NO."}:
                widget = QSpinBox()
                widget.setRange(2000, 2100) if column_name == "사업년도" else widget.setRange(0, 99999)
                widget.setValue(datetime.now().year if column_name == "사업년도" else 1)
                widget.valueChanged.connect(self.refresh_specimen_generated_fields)
            elif column_name in SPECIMEN_DROPDOWN_OPTIONS:
                widget = QComboBox()
                widget.addItems(SPECIMEN_DROPDOWN_OPTIONS[column_name])
                if column_name in SPECIMEN_DROPDOWN_DEFAULTS:
                    widget.setCurrentText(SPECIMEN_DROPDOWN_DEFAULTS[column_name])
                if column_name in {"대륙/해양 국문", "대륙/해양 영문", "채집시간", "표본제작형태", "분류코드"}:
                    widget.setEditable(True)
                widget.currentTextChanged.connect(self.refresh_specimen_generated_fields)
            else:
                widget = QLineEdit()
                if column_name in DATE_FIELDS:
                    widget.setPlaceholderText("2016-01-01 (260508, 20260508도 자동 변환)")
                    widget.editingFinished.connect(lambda field=column_name: self.normalize_specimen_date_field(field))
                elif column_name == "위도":
                    widget.setPlaceholderText("36.123456")
                elif column_name == "경도":
                    widget.setPlaceholderText("126.123456")
                elif column_name == "종고유ID":
                    widget.setPlaceholderText("종정보 마스터의 종고유ID")
                elif column_name in {"원본경로", "원본파일명", "파일경로", "파일명"}:
                    widget.setReadOnly(True)
                    if column_name == "파일경로":
                        widget.setPlaceholderText("원본 이미지 선택 버튼으로 선택")
                    elif column_name == "원본경로":
                        widget.setPlaceholderText("저장 전 원본 파일 경로(자동 기록, 이후 저장해도 유지)")
                elif column_name == "국명":
                    widget.setPlaceholderText("입력 후 Enter 또는 국명 조회")
                    widget.editingFinished.connect(self.fill_specimen_taxon_from_korean_name)
                elif column_name == "학명":
                    widget.editingFinished.connect(lambda: self.fill_specimen_taxon_from_korean_name(prefer_scientific=True))
                elif column_name in SPECIMEN_TAXON_AUTOFILL_FIELDS:
                    widget.setPlaceholderText("국명 조회 시 자동 입력 (직접 수정 시 노란색 표시)")
                    widget.textEdited.connect(lambda _text, field=column_name: self.mark_specimen_taxon_edited(field))
                elif column_name in SPECIMEN_ADDRESS_FIELDS:
                    if column_name == "채집위치 상세 국문":
                        widget.setPlaceholderText("직접 입력 (자동완성 후보 제공, 영문 자동 매칭)")
                        completer = QCompleter(sorted(self.address_master.candidates("detail")), widget)
                        completer.setCaseSensitivity(Qt.CaseInsensitive)
                        completer.setModelSorting(QCompleter.CaseInsensitivelySortedModel)
                        completer.setFilterMode(Qt.MatchStartsWith)
                        widget.setCompleter(completer)
                    else:
                        widget.setPlaceholderText("국문 입력 후 Enter (영문 자동 매칭)")
                    widget.editingFinished.connect(
                        lambda field=column_name: self.fill_specimen_address_english(field)
                    )
            if column_name in TEXT_DEFAULTS and isinstance(widget, QLineEdit):
                widget.setText(TEXT_DEFAULTS[column_name])
            if column_name == "데이터베이스번호":
                widget.editingFinished.connect(lambda: self.edit_database_number(True))
            if column_name == "사업별 순번":
                widget.textEdited.connect(lambda: self.refresh_specimen_generated_fields(force_number=True))
                widget.editingFinished.connect(lambda: self.commit_project_sequence(True))
            if column_name in SPECIMEN_AUTO_FIELDS and column_name not in {"데이터베이스번호", "사업별 순번"}:
                widget.setEnabled(False)
                if isinstance(widget, QLineEdit):
                    widget.setReadOnly(True)
                if isinstance(widget, QSpinBox):
                    widget.setButtonSymbols(QSpinBox.NoButtons)
            self.specimen_form_widgets[column_name] = widget
            self.connect_specimen_autosave_signal(widget)
            if column_name == "파일경로":
                widget.textChanged.connect(self.sync_specimen_media_filename_field)
                widget.textChanged.connect(self.update_specimen_media_thumbnail)
                media_layout = QHBoxLayout()
                media_layout.addWidget(widget, 1)
                select_media_button = QPushButton("원본 선택")
                select_media_button.clicked.connect(self.select_specimen_image_source_file)
                media_layout.addWidget(select_media_button)
                self.specimen_row_form.addRow(label, media_layout)

                self.specimen_media_thumbnail_label = ThumbnailLabel()
                self.specimen_media_thumbnail_label.setFixedSize(180, 180)
                self.specimen_media_thumbnail_label.setAlignment(Qt.AlignCenter)
                self.specimen_media_thumbnail_label.setStyleSheet("border: 1px solid #ccc; background: #fafafa;")
                self.specimen_row_form.addRow("미리보기", self.specimen_media_thumbnail_label)
            else:
                self.specimen_row_form.addRow(label, widget)

    def apply_specimen_project_to_form(self, label: str = ""):
        self.apply_project_to_form(label, self.specimen_form_widgets, self.refresh_specimen_generated_fields)

    def apply_specimen_organization_to_form(self, label: str = ""):
        self.apply_organization_to_form(label, self.specimen_form_widgets, self.refresh_specimen_generated_fields)

    def refresh_specimen_generated_fields(self, *, force_number=False):
        if self.specimen_autosave_suspended:
            return
        if not self.get_form_value("사업별 순번", self.specimen_form_widgets):
            self.set_form_value("사업별 순번", self.project_sequence_for_form(True), self.specimen_form_widgets)
        if not force_number and self.restore_existing_database_number(True):
            return
        organization_code = self.get_form_value("기관코드", self.specimen_form_widgets)
        taxon_code = self.get_form_value("분류코드", self.specimen_form_widgets)
        year = self.get_form_value("사업년도", self.specimen_form_widgets)
        order = self.get_form_value("순번", self.specimen_form_widgets)
        no = self.get_form_value("사업별 순번", self.specimen_form_widgets)
        if organization_code and taxon_code and year and order and str(no).isdigit():
            year_text = str(year)[-2:].zfill(2)
            self.set_form_value(
                "데이터베이스번호",
                f"{organization_code}-SP-{taxon_code}-{year_text}{int(order):02d}{int(no):05d}",
                self.specimen_form_widgets,
            )
        else:
            self.set_form_value("데이터베이스번호", "", self.specimen_form_widgets)

    def mark_specimen_taxon_edited(self, field: str):
        self._specimen_taxon_edited_fields.add(field)
        widget = self.specimen_form_widgets.get(field)
        if isinstance(widget, QLineEdit):
            widget.setStyleSheet(f"background-color: {SPECIMEN_TAXON_EDITED_COLOR.name()};")

    def reset_specimen_taxon_edit_tracking(self):
        self._specimen_taxon_edited_fields = set()
        for field in SPECIMEN_TAXON_AUTOFILL_FIELDS:
            widget = self.specimen_form_widgets.get(field)
            if isinstance(widget, QLineEdit):
                widget.setStyleSheet("")

    def fill_specimen_taxon_from_korean_name(self, prefer_scientific=False):
        self.ensure_species_master_current()
        korean_name = normalize_text(self.get_form_value("국명", self.specimen_form_widgets))
        scientific_input = normalize_text(self.get_form_value("학명", self.specimen_form_widgets))
        if not korean_name and not scientific_input:
            self.set_specimen_species_status("")
            return

        if self.species_master_df.empty:
            self.set_specimen_species_status("종정보 마스터 없음")
            return

        matches = self.find_species_matches(korean_name, scientific_input, prefer_scientific)
        if not matches:
            self.set_specimen_species_status("마스터 매칭 없음")
            return

        match = matches[0]
        self.set_form_value("국명", normalize_text(match.get("국명")), self.specimen_form_widgets)
        scientific_name = normalize_text(match.get("학명"))
        species_id = self.clean_master_value(match.get("종고유ID"))
        if species_id in {"", "0"}:
            species_id = self.clean_master_value(match.get("종학명정보ID"))
        genus_name = normalize_text(match.get("속명국문명"))
        species_epithet = normalize_text(match.get("종소명국문명"))

        self.set_form_value("학명", scientific_name, self.specimen_form_widgets)
        self.set_form_value("종고유ID", species_id, self.specimen_form_widgets)
        # 과명은 종정보 마스터에 컬럼이 없어 자동으로 채우지 못함 -> 그대로 두고 직접 입력받음
        # (직접 입력 즉시 mark_specimen_taxon_edited로 노란색 표시됨)
        for field, value in (("속명", genus_name), ("종명", species_epithet)):
            self._specimen_taxon_edited_fields.discard(field)
            widget = self.specimen_form_widgets.get(field)
            if isinstance(widget, QLineEdit):
                widget.setStyleSheet("")
            self.set_form_value(field, value, self.specimen_form_widgets)

        if len(matches) > 1:
            self.set_specimen_species_status(f"중복 {len(matches)}건, 첫 항목 적용")
        else:
            self.set_specimen_species_status("마스터 적용 완료")

    def fill_specimen_address_english(self, korean_field: str):
        english_field, level = SPECIMEN_ADDRESS_FIELDS[korean_field]
        korean_text = normalize_text(self.get_form_value(korean_field, self.specimen_form_widgets))
        if not korean_text:
            return

        english_text, exact = self.address_master.match_english(level, korean_text)
        if not english_text:
            self.set_specimen_save_status(f"{korean_field}: 주소 마스터에 없음 (직접 입력 필요)")
            return

        self.set_form_value(english_field, english_text, self.specimen_form_widgets)
        if exact:
            self.set_specimen_save_status(f"{korean_field} → {english_text} 자동 적용")
        else:
            self.set_specimen_save_status(f"{korean_field} → {english_text} (유사 매칭, 확인 필요)")

    def open_specimen_location_on_map(self):
        self.open_location_on_map_for(
            self.get_form_value("위도", self.specimen_form_widgets),
            self.get_form_value("경도", self.specimen_form_widgets),
        )

    def open_location_on_map_for(self, latitude_text: str, longitude_text: str):
        latitude = self.to_float(latitude_text)
        longitude = self.to_float(longitude_text)
        if latitude is None or longitude is None:
            QMessageBox.warning(self, "지도 확인", "위도/경도를 먼저 입력해 주세요.")
            return
        # 지도 라이브러리를 앱에 내장하면 빌드 용량이 크게 늘어나서(QtWebEngine 150MB+),
        # 기본 브라우저로 OpenStreetMap을 열어 좌표 위치에 핀을 찍어 보여주는 방식으로 처리
        url = f"https://www.openstreetmap.org/?mlat={latitude}&mlon={longitude}#map=17/{latitude}/{longitude}"
        QDesktopServices.openUrl(QUrl(url))

    def open_all_specimen_locations_on_map(self):
        self.open_all_locations_on_map_for(
            self.specimen_data_table, SPECIMEN_ROW_COLUMNS, SPECIMEN_ADDRESS_FIELDS, "specimen_map_preview.html"
        )

    def open_all_locations_on_map_for(
        self, table: QTableWidget, row_columns: list[str], address_fields: dict, file_name: str
    ):
        points = self.collect_map_points(table, row_columns, address_fields)
        if not points:
            QMessageBox.information(self, "지도 확인", "위도/경도가 입력된 행이 없습니다.")
            return

        try:
            url = publish_map_preview(self.build_locations_map_html(points))
            if not QDesktopServices.openUrl(QUrl(url)):
                raise RuntimeError("기본 브라우저를 열 수 없습니다.")
        except Exception as exc:
            QMessageBox.warning(self, "지도 확인 실패", str(exc))
            return
        # 몇 개 행이 지도에 찍혔는지 바로 알 수 있게 표시 (위도/경도가 비어있는 행은
        # 애초에 지도에 안 찍히므로, 핀이 하나만 보이면 그만큼만 좌표가 채워져 있다는 뜻)
        QMessageBox.information(self, "지도 확인", f"위치가 입력된 {len(points)}개 행을 지도에 표시했습니다.")

    @staticmethod
    def collect_map_points(table: QTableWidget, row_columns: list[str], address_fields: dict) -> list[dict]:
        lat_col = row_columns.index("위도")
        lon_col = row_columns.index("경도")
        no_col = row_columns.index("데이터베이스번호")
        # 상세(detail)는 역지오코딩 결과에 없는 필드라 주소 표시에서 제외
        address_cols = [
            row_columns.index(korean_field) for korean_field, (_, key) in address_fields.items() if key != "detail"
        ]

        points = []
        for row in range(table.rowCount()):
            lat_item = table.item(row, lat_col)
            lon_item = table.item(row, lon_col)
            latitude = UploadTab.to_float(lat_item.text() if lat_item else "")
            longitude = UploadTab.to_float(lon_item.text() if lon_item else "")
            if latitude is None or longitude is None:
                continue
            if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
                continue

            no_item = table.item(row, no_col)
            db_no = no_item.text() if no_item and no_item.text() else f"{row + 1}행"
            address_parts = [table.item(row, col).text() for col in address_cols if table.item(row, col) and table.item(row, col).text()]
            points.append(
                {
                    "no": db_no,
                    "lat": latitude,
                    "lon": longitude,
                    "address": " ".join(address_parts) or "주소 없음",
                }
            )
        return points

    @staticmethod
    def build_locations_map_html(points: list[dict]) -> str:
        # </script> 시퀀스가 주소 텍스트에 섞여 들어와도 스크립트 태그가 깨지지 않게 이스케이프
        points_json = json.dumps(points, ensure_ascii=False).replace("</", "<\\/")
        return f"""<!doctype html>
<html><head><meta charset="utf-8" />
<title>BioFlow 위치 확인</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
html, body, #map {{ height: 100%; margin: 0; }}
#status {{ position: absolute; top: 12px; left: 60px; right: 12px;
  z-index: 1000; padding: 10px; background: white; border-radius: 4px;
  font: 14px sans-serif; }}
</style>
</head>
<body>
<div id="map"></div>
<div id="status" role="status">지도를 불러오는 중입니다. 인터넷 연결이 필요합니다.</div>
<script>
  var status = document.getElementById('status');
  if (typeof L === 'undefined') {{
    status.textContent = '지도 프로그램을 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침해 주세요.';
  }} else {{
  var points = {points_json};
  var map = L.map('map', {{ maxZoom: 19 }});
  var tiles = L.tileLayer('https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
  }});
  var failedTiles = false;
  tiles.on('loading', function () {{
    failedTiles = false;
    status.hidden = false;
    status.textContent = '지도 배경을 불러오는 중입니다.';
  }});
  tiles.on('tileerror', function () {{
    failedTiles = true;
    status.hidden = false;
    status.textContent = '일부 지도 배경을 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침해 주세요.';
  }});
  tiles.on('load', function () {{ status.hidden = !failedTiles; }});
  tiles.addTo(map);
  function escapeHtml(value) {{
    var element = document.createElement('span');
    element.textContent = String(value);
    return element.innerHTML;
  }}
  var markers = points.map(function (p) {{
    return L.marker([p.lat, p.lon])
      .bindPopup('<b>' + escapeHtml(p.no) + '</b><br>' + p.lat + ', ' + p.lon + '<br>' + escapeHtml(p.address))
      .addTo(map);
  }});
  if (markers.length) {{
    map.fitBounds(L.featureGroup(markers).getBounds().pad(0.2), {{ maxZoom: 17 }});
  }} else {{
    map.setView([36.5, 127.8], 7);
  }}
  }}
</script>
</body>
</html>
"""

    def fill_all_specimen_addresses_from_latlon(self):
        self.batch_fill_addresses(
            self.specimen_data_table, SPECIMEN_ROW_COLUMNS, SPECIMEN_ADDRESS_FIELDS, self.set_specimen_save_status
        )

    def fill_all_specimen_species_and_project(self):
        self.batch_fill_species_and_project(
            self.specimen_data_table,
            SPECIMEN_ROW_COLUMNS,
            self.specimen_form_widgets,
            self.set_specimen_save_status,
            self.schedule_specimen_autosave,
        )

    def batch_fill_addresses(self, table: QTableWidget, row_columns: list[str], address_fields: dict, status_setter):
        lat_col = row_columns.index("위도")
        lon_col = row_columns.index("경도")
        # 시도가 이미 채워진 행은 재조회하지 않음 (Nominatim 초당 1건 제한이라 전체를 다시
        # 돌리면 오래 걸림 -> 비어있는 행만 대상으로)
        sido_field = next(k for k, (_, key) in address_fields.items() if key == "sido")
        sido_col = row_columns.index(sido_field)

        targets = []
        for row in range(table.rowCount()):
            lat_item = table.item(row, lat_col)
            lon_item = table.item(row, lon_col)
            latitude = self.to_float(lat_item.text() if lat_item else "")
            longitude = self.to_float(lon_item.text() if lon_item else "")
            if latitude is None or longitude is None:
                continue
            sido_item = table.item(row, sido_col)
            if sido_item and sido_item.text():
                continue
            targets.append((row, latitude, longitude))

        if not targets:
            QMessageBox.information(
                self, "주소 일괄 조회", "주소를 조회할 행이 없습니다 (위도/경도가 있고 주소가 비어있는 행만 대상)."
            )
            return

        progress = QProgressDialog(
            "주소를 조회하는 중... (조회 서비스 제한으로 행마다 시간이 걸립니다)", "취소", 0, len(targets), self
        )
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(300)

        field_triples = [
            (korean_field, english_field, key)
            for korean_field, (english_field, key) in address_fields.items()
            if key != "detail"
        ]

        success = 0
        failed = 0
        last_error = ""
        was_blocked = table.blockSignals(True)
        try:
            for index, (row, latitude, longitude) in enumerate(targets):
                progress.setLabelText(f"{index + 1}/{len(targets)}행 주소 조회 중...")
                progress.setValue(index)
                QApplication.processEvents()
                if progress.wasCanceled():
                    break

                try:
                    address_ko = self.reverse_geocode(latitude, longitude, "ko")
                    # Nominatim 무료 API는 초당 1건이 상한선이라, 요청 자체에 걸리는 시간까지
                    # 감안해서 여유를 두지 않으면(정확히 1초만 쉬면) 한도를 넘겨 조용히
                    # 차단(빈 결과/오류)당하는 경우가 있었음 -> 여유를 더 둠
                    time.sleep(1.5)
                    address_en = self.reverse_geocode(latitude, longitude, "en")
                    time.sleep(1.5)
                except Exception as exc:
                    failed += 1
                    last_error = str(exc)
                    continue
                if not address_ko:
                    failed += 1
                    last_error = "조회 결과 없음"
                    continue

                for korean_field, english_field, key in field_triples:
                    korean_value = address_ko.get(key, "")
                    if not korean_value:
                        continue
                    table.setItem(row, row_columns.index(korean_field), QTableWidgetItem(korean_value))

                    english_value = address_en.get(key, "") if address_en else ""
                    if english_value and english_value != korean_value:
                        table.setItem(row, row_columns.index(english_field), QTableWidgetItem(english_value))
                    else:
                        english_text, _ = self.address_master.match_english(key, korean_value)
                        if english_text:
                            table.setItem(row, row_columns.index(english_field), QTableWidgetItem(english_text))
                success += 1
        finally:
            table.blockSignals(was_blocked)
            progress.setValue(len(targets))

        failure_note = f", 조회 실패 {failed}건 (마지막 사유: {last_error})" if failed else ""
        status_setter(f"주소 일괄 조회 완료: 성공 {success}/{len(targets)}행{failure_note}")

    def fill_specimen_address_from_latlon(self):
        latitude = self.to_float(self.get_form_value("위도", self.specimen_form_widgets))
        longitude = self.to_float(self.get_form_value("경도", self.specimen_form_widgets))
        if latitude is None or longitude is None:
            QMessageBox.warning(self, "주소 조회", "위도/경도를 먼저 입력해 주세요.")
            return

        try:
            address_ko = self.reverse_geocode(latitude, longitude, "ko")
            address_en = self.reverse_geocode(latitude, longitude, "en")
        except Exception as exc:
            self.set_specimen_save_status(f"주소 조회 실패: {exc}")
            return

        if not address_ko:
            self.set_specimen_save_status("주소 조회 결과가 없습니다.")
            return

        for korean_field, english_field, key in (
            ("채집위치 시도 국문", "채집위치 시도 영문", "sido"),
            ("채집위치 시군구 국문", "채집위치 시군구 영문", "sigungu"),
            ("채집위치 읍면동 국문", "채집위치 읍면동 영문", "eupmyeondong"),
        ):
            korean_value = address_ko.get(key, "")
            if not korean_value:
                continue
            self.set_form_value(korean_field, korean_value, self.specimen_form_widgets)

            english_value = address_en.get(key, "")
            if english_value and english_value != korean_value:
                self.set_form_value(english_field, english_value, self.specimen_form_widgets)
            else:
                self.fill_specimen_address_english(korean_field)

        self.set_specimen_save_status("위도/경도 기반 주소를 조회했습니다. (OpenStreetMap Nominatim, 내용 확인 필요)")

    def import_specimen_media_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "이미지 폴더 선택", "")
        if not folder:
            return

        # 표본은 이미지 단일 등록만 지원 -> 폴더에 동영상/사운드가 섞여 있어도 이미지만 골라서 씀
        image_files = [path for path, media_type in self.scan_media_folder(Path(folder)) if media_type == "이미지"]
        if not image_files:
            QMessageBox.information(self, "폴더 불러오기", "지원하는 이미지 파일을 찾지 못했습니다.")
            return

        progress = QProgressDialog("이미지 파일 정보를 읽는 중...", "취소", 0, len(image_files), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(300)

        added_count = 0
        self.specimen_autosave_suspended = True
        try:
            # 폼에 이전에 작성하던 국명/학명 등 내용이 남아있으면 새로 만드는 모든 행에
            # 그대로 찍혀버리므로, 사업명/기관/분류코드만 남기고 비운 뒤 시작
            self.clear_specimen_form()
            for index, file_path in enumerate(image_files):
                progress.setLabelText(f"{file_path.name} 처리 중... ({index + 1}/{len(image_files)})")
                progress.setValue(index)
                QApplication.processEvents()
                if progress.wasCanceled():
                    break

                self.set_form_value("NO.", self.next_specimen_no_value(), self.specimen_form_widgets)
                self.set_form_value("파일경로", str(file_path), self.specimen_form_widgets)
                self.set_form_value("원본경로", str(file_path), self.specimen_form_widgets)
                self.set_form_value("원본파일명", file_path.stem, self.specimen_form_widgets)
                self.apply_location_from_specimen_media(file_path)
                self.apply_date_from_specimen_media(file_path)
                self.add_specimen_data_row(self.collect_specimen_form_values())
                added_count += 1
        finally:
            self.specimen_autosave_suspended = False
            progress.setValue(len(image_files))

        self.refresh_location_validation(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
        self.autosave_specimen_draft()
        cancel_note = "\n(취소되어 일부만 불러왔습니다.)" if added_count < len(image_files) else ""
        QMessageBox.information(
            self,
            "폴더 불러오기",
            f"{added_count}건의 이미지를 행으로 추가했습니다.{cancel_note}\n"
            "각 행을 선택해 종명 등 나머지 정보를 입력한 뒤 '선택 행 수정'으로 저장해 주세요.",
        )
        self.set_specimen_save_status(f"폴더 불러오기 완료: {added_count}건 (나머지 정보 입력 필요)")

    def select_specimen_image_source_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "원본 이미지 선택", "", self.media_file_filter("이미지"))
        if not file_path:
            return

        self.set_form_value("파일경로", file_path, self.specimen_form_widgets)
        # 새로 고른 원본 파일 정보를 바로 기록 (저장 시 파일경로/파일명만 새 저장위치로 바뀜)
        self.set_form_value("원본경로", file_path, self.specimen_form_widgets)
        self.set_form_value("원본파일명", Path(file_path).stem, self.specimen_form_widgets)
        self.apply_location_from_specimen_media(Path(file_path))
        self.apply_date_from_specimen_media(Path(file_path))
        self.schedule_specimen_autosave()

    def apply_location_from_specimen_media(self, source_path: Path):
        latitude, longitude, message = self.extract_location_from_media(source_path, "이미지")
        if latitude is None or longitude is None:
            self.set_specimen_save_status(message)
            return

        self.set_form_value("위도", f"{latitude:.7f}", self.specimen_form_widgets)
        self.set_form_value("경도", f"{longitude:.7f}", self.specimen_form_widgets)
        self.set_specimen_save_status(f"이미지 위치정보 적용: 위도 {latitude:.7f}, 경도 {longitude:.7f}")

    def apply_date_from_specimen_media(self, source_path: Path):
        date_text = self.extract_media_date(source_path, "이미지")
        if date_text:
            self.set_form_value("채집년월일", date_text, self.specimen_form_widgets)

    def prepare_image_for_specimen_record(self, values: dict[str, Any]) -> dict[str, Any]:
        if getattr(self, "_defer_media_copy", False):
            return values
        media_path_text = normalize_text(values.get("파일경로"))
        if not media_path_text:
            return values

        source_path = Path(media_path_text)
        if not source_path.exists():
            return values

        saved_path = self.save_specimen_media_file_to_storage(source_path, values)
        if saved_path:
            values = values.copy()
            if not normalize_text(values.get("원본경로")):
                # 최초 저장 시 1번만 기록 -> 재저장해도 진짜 원본 경로가 덮어써지지 않음
                values["원본경로"] = str(source_path)
                values["원본파일명"] = source_path.stem
                self.set_form_value("원본경로", str(source_path), self.specimen_form_widgets)
                self.set_form_value("원본파일명", source_path.stem, self.specimen_form_widgets)
            values["파일경로"] = str(saved_path)
            values["파일명"] = saved_path.name
            self.set_form_value("파일경로", str(saved_path), self.specimen_form_widgets)
            self.apply_location_from_specimen_media(saved_path)
            values["위도"] = self.get_form_value("위도", self.specimen_form_widgets)
            values["경도"] = self.get_form_value("경도", self.specimen_form_widgets)
        return values

    def apply_media_copy_to_specimen_row_values(self, values: dict[str, Any]) -> dict[str, Any]:
        media_path_text = normalize_text(values.get("파일경로"))
        if not media_path_text:
            return values

        source_path = Path(media_path_text)
        if not source_path.is_file():
            source_path = Path(normalize_text(values.get("원본경로")))
            if not source_path.is_file():
                return values

        saved_path = self.save_specimen_media_file_to_storage(source_path, values)
        if not saved_path:
            return values

        values = values.copy()
        if not normalize_text(values.get("원본경로")):
            values["원본경로"] = str(source_path)
            values["원본파일명"] = source_path.stem
        values["파일경로"] = str(saved_path)
        values["파일명"] = saved_path.name
        latitude, longitude, _ = self.extract_location_from_media(saved_path, "이미지")
        if latitude is not None and longitude is not None:
            values["위도"] = f"{latitude:.7f}"
            values["경도"] = f"{longitude:.7f}"
        return values

    def copy_pending_media_for_all_specimen_rows(self, strict=False):
        for row in range(self.specimen_data_table.rowCount()):
            values = {
                column: (self.specimen_data_table.item(row, col).text() if self.specimen_data_table.item(row, col) else "")
                for col, column in enumerate(SPECIMEN_ROW_COLUMNS)
            }
            updated = self.apply_media_copy_to_specimen_row_values(values)
            if strict and normalize_text(values.get("파일경로")) and updated is values:
                return False
            if updated is not values:
                self.write_specimen_data_row(row, updated)

        return True

    def select_specimen_media_storage_dir(self):
        selected_dir = QFileDialog.getExistingDirectory(
            self, "이미지 저장 위치 선택", self.specimen_media_storage_dir or ""
        )
        if not selected_dir:
            return
        self.specimen_media_storage_dir = selected_dir
        self.specimen_media_storage_path_input.setText(selected_dir)
        self.autosave_specimen_draft()

    def ensure_specimen_media_storage_dir(self) -> Path | None:
        if self.specimen_media_storage_dir:
            return Path(self.specimen_media_storage_dir)

        selected_dir = QFileDialog.getExistingDirectory(self, "이미지 저장 위치 선택", "")
        if not selected_dir:
            return None
        self.specimen_media_storage_dir = selected_dir
        self.specimen_media_storage_path_input.setText(selected_dir)
        return Path(selected_dir)

    def save_specimen_media_file_to_storage(self, source_path: Path, values: dict[str, Any] | None = None) -> Path | None:
        values = values or self.collect_specimen_form_values()
        db_number = normalize_text(values.get("데이터베이스번호"))
        if not db_number:
            QMessageBox.warning(self, "이미지 저장", "DB번호를 먼저 만들 수 있도록 분류코드, 사업, 기관 값이 필요합니다.")
            return None

        target_dir = self.ensure_specimen_media_storage_dir()
        if target_dir is None:
            self.set_specimen_save_status("이미지 저장 위치 선택이 취소되었습니다.")
            return None

        expected_stem = self.safe_filename(db_number)
        same_folder = source_path.parent.resolve() == target_dir.resolve()
        if same_folder and source_path.stem == expected_stem:
            # 이미 이 폴더에 DB번호 이름으로 저장돼 있는 파일 -> 복사할 필요 없음
            return source_path
        original_path = normalize_text(values.get("원본경로"))
        already_copied = bool(original_path and Path(original_path).resolve() != source_path.resolve())
        if same_folder and not already_copied:
            QMessageBox.warning(
                self,
                "이미지 저장",
                "이미지 저장 위치가 원본 파일이 있는 폴더와 같습니다. "
                "같은 폴더로 복사하면 원본과 뒤섞이니, 원본과 다른 폴더를 저장 위치로 선택해 주세요.",
            )
            self.set_specimen_save_status("이미지 저장 실패: 저장 위치가 원본 폴더와 동일합니다.")
            return None

        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            target_path = copy_database_media(source_path, target_dir, expected_stem)
            self.set_specimen_save_status(f"이미지 저장 완료: {target_path}")
            return target_path
        except Exception as exc:
            QMessageBox.critical(self, "이미지 저장 실패", str(exc))
            self.set_specimen_save_status(f"이미지 저장 실패: {exc}")
            return None

    def add_specimen_data_row_from_form(self):
        self.set_form_value("NO.", self.next_specimen_no_value(), self.specimen_form_widgets)
        values = self.collect_specimen_form_values()
        values = self.prepare_new_project_number(values, True)
        if not self.validate_specimen_form_values(values):
            return
        values = self.prepare_image_for_specimen_record(values)
        self.add_specimen_data_row(values)
        self.prepare_next_specimen_input(values)
        self.autosave_specimen_draft()

    def update_selected_specimen_data_row(self):
        row = self.current_specimen_data_row()
        if row is None:
            QMessageBox.warning(self, "선택 행 수정", "수정할 행을 먼저 선택해 주세요.")
            return
        values = self.collect_specimen_form_values()
        if not self.validate_specimen_form_values(values):
            return
        values = self.prepare_image_for_specimen_record(values)
        self.write_specimen_data_row(row, values)
        self.autosave_specimen_draft()
        self.set_specimen_save_status("선택 행을 수정했습니다.")

    def load_selected_specimen_row_to_form(self):
        previous_row = self._specimen_last_selected_row
        row = self.current_specimen_data_row()

        if previous_row is not None and previous_row != row:
            values = self.collect_specimen_form_values()
            if self.has_specimen_form_content(values):
                self.write_specimen_data_row(previous_row, values)

        self._specimen_last_selected_row = row
        if row is None:
            return
        self.specimen_autosave_suspended = True
        self._specimen_taxon_edited_fields = set()
        for col, column_name in enumerate(SPECIMEN_ROW_COLUMNS):
            item = self.specimen_data_table.item(row, col)
            text = item.text() if item else ""
            if column_name in SPECIMEN_STICKY_FIELDS and not text:
                continue
            self.set_form_value(column_name, text, self.specimen_form_widgets)
            if column_name in SPECIMEN_TAXON_AUTOFILL_FIELDS:
                is_edited = bool(item and item.background().color() == SPECIMEN_TAXON_EDITED_COLOR)
                widget = self.specimen_form_widgets.get(column_name)
                if is_edited:
                    self._specimen_taxon_edited_fields.add(column_name)
                if isinstance(widget, QLineEdit):
                    widget.setStyleSheet(
                        f"background-color: {SPECIMEN_TAXON_EDITED_COLOR.name()};" if is_edited else ""
                    )
        self.specimen_autosave_suspended = False

    def on_specimen_data_table_item_changed(self, item: QTableWidgetItem):
        if self.specimen_autosave_suspended:
            return
        column_name = SPECIMEN_ROW_COLUMNS[item.column()]
        if column_name == "사업별 순번":
            self.renumber_project_sequences_below(self.specimen_data_table, SPECIMEN_ROW_COLUMNS, item.row(), True)
        if column_name in {"사업년도", "순번"}:
            self.recalculate_project_sequences(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
        if column_name == "데이터베이스번호":
            self.mark_database_number_manual(self.specimen_data_table, item)
        if column_name in SPECIMEN_TAXON_AUTOFILL_FIELDS:
            item.setBackground(SPECIMEN_TAXON_EDITED_COLOR)
        if item.column() == SPECIMEN_ROW_COLUMNS.index("NO."):
            self.renumber_specimen_rows_below(item.row())
        if column_name in {"위도", "경도"}:
            self.revalidate_row_location(self.specimen_data_table, SPECIMEN_ROW_COLUMNS, item.row())
        if item.row() == self.current_specimen_data_row():
            self.load_selected_specimen_row_to_form()
            if column_name in {"국명", "학명"}:
                self.fill_specimen_taxon_from_korean_name(prefer_scientific=column_name == "학명")
                self.write_specimen_data_row(item.row(), self.collect_specimen_form_values())
            elif column_name in DATE_FIELDS:
                self.normalize_specimen_date_field(column_name)
                self.write_specimen_data_row(item.row(), self.collect_specimen_form_values())
        self.schedule_specimen_autosave()

    def renumber_specimen_rows_below(self, changed_row: int):
        no_column = SPECIMEN_ROW_COLUMNS.index("NO.")
        item = self.specimen_data_table.item(changed_row, no_column)
        try:
            next_no = int(item.text() if item else "") + 1
        except ValueError:
            return

        was_blocked = self.specimen_data_table.blockSignals(True)
        try:
            for row in range(changed_row + 1, self.specimen_data_table.rowCount()):
                self.specimen_data_table.setItem(row, no_column, QTableWidgetItem(str(next_no)))
                next_no += 1
        finally:
            self.specimen_data_table.blockSignals(was_blocked)

        if not self.specimen_autosave_suspended:
            self.recalculate_project_sequences(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)

    @staticmethod
    def default_specimen_field_value(column_name: str) -> Any:
        if column_name in SPECIMEN_DROPDOWN_DEFAULTS:
            return SPECIMEN_DROPDOWN_DEFAULTS[column_name]
        if column_name in TEXT_DEFAULTS:
            return TEXT_DEFAULTS[column_name]
        if column_name in SPECIMEN_DROPDOWN_OPTIONS and SPECIMEN_DROPDOWN_OPTIONS[column_name]:
            return SPECIMEN_DROPDOWN_OPTIONS[column_name][0]
        return ""

    def clear_specimen_form(self):
        self._specimen_last_selected_row = None
        blocked = self.specimen_data_table.blockSignals(True)
        try:
            self.specimen_data_table.clearSelection()
        finally:
            self.specimen_data_table.blockSignals(blocked)
        for column_name in SPECIMEN_ROW_COLUMNS:
            if column_name in SPECIMEN_STICKY_FIELDS:
                continue
            if column_name == "NO.":
                self.set_form_value(column_name, self.next_specimen_no_value(), self.specimen_form_widgets)
            else:
                self.set_form_value(
                    column_name, self.default_specimen_field_value(column_name), self.specimen_form_widgets
                )
        self.reset_specimen_taxon_edit_tracking()
        self.apply_specimen_project_to_form()
        self.set_default_organization(self.specimen_form_widgets)
        self.apply_specimen_organization_to_form()
        self.set_form_value("사업별 순번", self.project_sequence_for_form(True, new=True), self.specimen_form_widgets)
        self.refresh_specimen_generated_fields(force_number=True)
        self.schedule_specimen_autosave()

    def collect_specimen_form_values(self) -> dict[str, Any]:
        self.refresh_specimen_generated_fields()
        return {
            column_name: self.get_form_value(column_name, self.specimen_form_widgets)
            for column_name in SPECIMEN_ROW_COLUMNS
        }

    def validate_specimen_form_values(self, values: dict[str, Any]) -> bool:
        required_fields = [
            "NO.", "분류코드", "사업명", "사업년도", "순번", "기관명", "기관코드", "국명", "학명",
        ]
        missing = [field for field in required_fields if values.get(field) in ("", 0)]
        if missing:
            QMessageBox.warning(self, "입력 확인", f"필수값을 입력해 주세요: {', '.join(missing)}")
            return False
        return True

    def prepare_next_specimen_input(self, values: dict[str, Any]):
        # 관찰 탭과 동일한 이유로 "새 입력"과 동일하게 폼을 완전히 리셋
        self.clear_specimen_form()

    def next_specimen_no_value(self) -> int:
        max_no = 0
        no_column = SPECIMEN_ROW_COLUMNS.index("NO.")
        for row in range(self.specimen_data_table.rowCount()):
            item = self.specimen_data_table.item(row, no_column)
            max_no = max(max_no, self.safe_sort_int(item.text() if item else 0))
        return max_no + 1

    def current_specimen_data_row(self) -> int | None:
        indexes = self.specimen_data_table.selectedIndexes()
        if not indexes:
            return None
        # selectedIndexes()는 정렬 보장이 없어서(드래그 방향에 따라 순서가 달라짐)
        # indexes[0]을 쓰면 셀 범위를 드래그로 선택했을 때 엉뚱한 행이 잡힘 -> 항상 맨 위 행으로
        return min(index.row() for index in indexes)

    def add_specimen_data_row(self, values: dict[str, Any]):
        row = self.specimen_data_table.rowCount()
        self.specimen_data_table.insertRow(row)
        self.write_specimen_data_row(row, values)

    def write_specimen_data_row(self, row: int, values: dict[str, Any]):
        old_sequence = self.specimen_data_table.item(row, SPECIMEN_ROW_COLUMNS.index("사업별 순번"))
        sequence_changed = old_sequence is not None and old_sequence.text().strip() != str(values.get("사업별 순번", "")).strip()
        values = self.preserve_row_database_number(self.specimen_data_table, SPECIMEN_ROW_COLUMNS, row, values)
        was_blocked = self.specimen_data_table.blockSignals(True)
        try:
            for col, column_name in enumerate(SPECIMEN_ROW_COLUMNS):
                value = values.get(column_name, "")
                item = QTableWidgetItem("" if value is None else str(value))
                if column_name == "데이터베이스번호":
                    item.setData(DB_NUMBER_MANUAL_ROLE, bool(values.get("_database_number_manual")))
                if column_name in SPECIMEN_TAXON_AUTOFILL_FIELDS and column_name in self._specimen_taxon_edited_fields:
                    item.setBackground(SPECIMEN_TAXON_EDITED_COLOR)
                self.specimen_data_table.setItem(row, col, item)
        finally:
            self.specimen_data_table.blockSignals(was_blocked)

        if not self.specimen_autosave_suspended:
            self.recalculate_project_sequences(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)

        if sequence_changed:
            self.renumber_project_sequences_below(self.specimen_data_table, SPECIMEN_ROW_COLUMNS, row, True)

    def paste_clipboard_into_specimen_table(self):
        text = QApplication.clipboard().text()
        if not text:
            return

        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        if not lines:
            return
        rows = [line.split("\t") for line in lines]

        # currentColumn()은 드래그로 범위를 선택했을 때 "마지막으로 지나간 셀"의 열이라
        # 목적지 왼쪽 위가 아닐 수 있음 -> 선택 범위 중 가장 왼쪽 위 셀로 고정.
        # 선택된 셀이 하나도 없으면 (0,0)으로 잘못 덮어쓰지 않도록 그냥 붙여넣기를 취소함
        indexes = self.specimen_data_table.selectedIndexes()
        if not indexes:
            self.set_specimen_save_status("붙여넣기 취소: 붙여넣을 셀을 먼저 선택해 주세요.")
            return
        start_row = min(index.row() for index in indexes)
        start_col = min(index.column() for index in indexes)

        was_blocked = self.specimen_data_table.blockSignals(True)
        try:
            for row_offset, row_values in enumerate(rows):
                target_row = start_row + row_offset
                while target_row >= self.specimen_data_table.rowCount():
                    self.specimen_data_table.insertRow(self.specimen_data_table.rowCount())
                for col_offset, value in enumerate(row_values):
                    target_col = start_col + col_offset
                    if target_col >= len(SPECIMEN_ROW_COLUMNS):
                        break
                    self.specimen_data_table.setItem(target_row, target_col, QTableWidgetItem(value))
        finally:
            self.specimen_data_table.blockSignals(was_blocked)

        # 붙여넣기 중엔 신호를 막아뒀으니, 현재 폼에 로드된 행이 이번에 붙여넣은 범위에
        # 포함됐다면 국명->학명 자동조회/날짜 형식변환을 셀 직접 수정 때와 동일하게 적용
        current_row = self.current_specimen_data_row()
        if current_row is not None and start_row <= current_row < start_row + len(rows):
            row_values = rows[current_row - start_row]
            touched_columns = {
                SPECIMEN_ROW_COLUMNS[start_col + offset]
                for offset in range(len(row_values))
                if start_col + offset < len(SPECIMEN_ROW_COLUMNS)
            }
            self.load_selected_specimen_row_to_form()
            if "국명" in touched_columns:
                self.fill_specimen_taxon_from_korean_name()
            for date_field in DATE_FIELDS & touched_columns:
                self.normalize_specimen_date_field(date_field)
            self.write_specimen_data_row(current_row, self.collect_specimen_form_values())

        self.refresh_location_validation(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
        self.autosave_specimen_draft()
        self.set_specimen_save_status(f"붙여넣기 완료: {len(rows)}행")

    def go_to_next_specimen_row(self):
        total = self.specimen_data_table.rowCount()
        if total == 0:
            self.set_specimen_save_status("작성된 행이 없습니다.")
            return

        row = self.current_specimen_data_row()
        if row is None and self._specimen_last_selected_row is None:
            values = self.collect_specimen_form_values()
            if self.has_specimen_form_content(values):
                reply = QMessageBox.question(
                    self,
                    "다음 행",
                    "'입력'으로 저장하지 않은 작성 중인 내용이 있습니다. "
                    "행을 선택하면 지금 작성 중인 내용은 사라집니다. 계속할까요?",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if reply != QMessageBox.Yes:
                    return

        next_row = 0 if row is None else row + 1
        if next_row >= total:
            self.set_specimen_save_status("마지막 행입니다.")
            return

        self.specimen_data_table.selectRow(next_row)

    def delete_selected_specimen_data_rows(self):
        rows = sorted({index.row() for index in self.specimen_data_table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.warning(self, "선택 행 삭제", "삭제할 행을 먼저 선택해 주세요.")
            return
        self.delete_selected_rows(self.specimen_data_table)
        self.autosave_specimen_draft()
        self.set_specimen_save_status(f"{len(rows)}개 행을 삭제했습니다.")

    def clear_specimen_data_rows(self):
        if self.specimen_data_table.rowCount() == 0:
            return
        reply = QMessageBox.question(self, "작성 내용 비우기", "작성 중인 행을 모두 비울까요?")
        if reply == QMessageBox.Yes:
            self.specimen_data_table.setRowCount(0)
            self.autosave_specimen_draft()

    def has_specimen_form_content(self, values: dict[str, Any]) -> bool:
        # has_form_content와 동일한 이유로 default_specimen_field_value와 직접 비교
        ignored_fields = SPECIMEN_STICKY_FIELDS | {"NO.", "데이터베이스번호"}
        for key, value in values.items():
            if key in ignored_fields:
                continue
            normalized = value if value is not None else ""
            if str(normalized) != str(self.default_specimen_field_value(key)):
                return True
        return False

    def sync_current_specimen_form_to_table(self):
        # 선택된 행이 있을 때만 그 행에 지금 폼 내용을 반영. 새 행 추가는 "입력" 버튼으로만
        # 하도록 함 - 예전엔 선택 없을 때 폼 내용으로 새 행을 추측해서 만들었는데, 그게
        # 저장(임시저장/엑셀로 저장) 시 의도치 않은 중복 행이 생기는 원인이었음.
        row = self.current_specimen_data_row()
        if row is None:
            return
        values = self.collect_specimen_form_values()
        if not self.has_specimen_form_content(values):
            return
        values = self.prepare_image_for_specimen_record(values)
        self.write_specimen_data_row(row, values)

    def connect_specimen_autosave_signal(self, widget: QLineEdit | QComboBox | QSpinBox):
        if isinstance(widget, QLineEdit):
            widget.textChanged.connect(self.schedule_specimen_autosave)
        elif isinstance(widget, QComboBox):
            widget.currentTextChanged.connect(self.schedule_specimen_autosave)
        elif isinstance(widget, QSpinBox):
            widget.valueChanged.connect(self.schedule_specimen_autosave)

    def schedule_specimen_autosave(self, *_):
        if self.specimen_autosave_suspended:
            return
        self.specimen_autosave_timer.start(700)

    def autosave_specimen_draft(self):
        if self.specimen_autosave_suspended or not hasattr(self, "specimen_data_table") or not self.specimen_form_widgets:
            return
        self.specimen_autosave_suspended = True
        try:
            records = self.table_to_records(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
            self.write_specimen_draft(records, self.collect_specimen_form_values())
            self.set_specimen_save_status("자동 임시저장 완료")
        finally:
            self.specimen_autosave_suspended = False

    def write_specimen_draft(self, rows: list[dict[str, Any]], current_form: dict[str, Any]):
        payload = {
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "rows": rows,
            "current_form": current_form,
            "media_storage_dir": self.specimen_media_storage_dir,
        }
        self.write_json(SPECIMEN_DRAFT_FILE, payload)

    def read_specimen_draft(self) -> dict[str, Any]:
        data = self.read_json(SPECIMEN_DRAFT_FILE, {})
        if not isinstance(data, dict):
            return {"rows": [], "current_form": {}, "media_storage_dir": ""}
        rows = data.get("rows", [])
        current_form = data.get("current_form", {})
        return {
            "rows": rows if isinstance(rows, list) else [],
            "current_form": current_form if isinstance(current_form, dict) else {},
            "media_storage_dir": str(data.get("media_storage_dir", "")),
        }

    def save_specimen_draft(self):
        self.sync_current_specimen_form_to_table()
        self.copy_pending_media_for_all_specimen_rows()
        records = self.table_to_records(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
        self.write_specimen_draft(records, self.collect_specimen_form_values())
        self.set_specimen_save_status(f"임시저장 완료: {len(records)}건")
        QMessageBox.information(
            self, "작성 내용 저장", f"작성 중인 표본 데이터 {len(records)}건을 저장했습니다.\n{SPECIMEN_DRAFT_FILE}"
        )

    def load_specimen_draft(self):
        if not SPECIMEN_DRAFT_FILE.exists():
            return
        draft = self.read_specimen_draft()
        self.specimen_autosave_suspended = True
        try:
            for record in draft["rows"]:
                self.add_specimen_data_row(record)
            current_form = draft.get("current_form", {})
            if current_form:
                for column_name in SPECIMEN_ROW_COLUMNS:
                    self.set_form_value(column_name, current_form.get(column_name, ""), self.specimen_form_widgets)
            self.specimen_media_storage_dir = draft.get("media_storage_dir", "")
            if self.specimen_media_storage_dir and hasattr(self, "specimen_media_storage_path_input"):
                self.specimen_media_storage_path_input.setText(self.specimen_media_storage_dir)
        finally:
            self.specimen_autosave_suspended = False
        self.refresh_location_validation(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)

    def reload_specimen_draft(self):
        self.specimen_autosave_suspended = True
        try:
            self.specimen_data_table.setRowCount(0)
        finally:
            self.specimen_autosave_suspended = False
        self.load_specimen_draft()
        self.set_specimen_save_status("임시저장을 다시 불러왔습니다.")

    def export_specimen_rows_to_excel(self):
        deferred = getattr(self, "_defer_media_copy", False)
        self._defer_media_copy = True
        try:
            self.sync_current_specimen_form_to_table()
        finally:
            self._defer_media_copy = deferred
        records = self.table_to_records(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
        if not records:
            QMessageBox.warning(self, "엑셀 저장", "저장할 행이 없습니다.")
            return
        if not self.confirm_location_issues_before_save(self.specimen_data_table, SPECIMEN_ROW_COLUMNS):
            return

        default_name = f"new_specimen_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "표본 데이터 엑셀 저장", default_name, "Excel Files (*.xlsx)")
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            if not self.prepare_media_for_excel_export(True, choose_destination=True):
                return
            records = self.table_to_records(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
            self.write_specimen_excel_file(file_path, records)
            final_path = self.final_excel_path(file_path)
            self.write_specimen_final_excel_file(str(final_path), records)
            gbif_path = self.gbif_excel_path(file_path)
            write_gbif_template(str(gbif_path), records, "specimen")
            self.reset_after_specimen_export()
            QMessageBox.information(
                self,
                "엑셀 저장",
                f"표본 데이터 엑셀 3개를 저장했습니다.\n{file_path}\n{final_path}\n{gbif_path}\n\n"
                "작성 중이던 내용은 초기화되어 NO.가 1부터 다시 시작됩니다.",
            )
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 저장 실패", str(exc))

    def reset_after_specimen_export(self):
        self.specimen_autosave_suspended = True
        try:
            self.specimen_data_table.setRowCount(0)
        finally:
            self.specimen_autosave_suspended = False
        self.clear_specimen_form()
        self.write_specimen_draft([], self.collect_specimen_form_values())
        self.set_specimen_save_status("엑셀 저장 완료. 새 작성을 시작합니다 (NO. 1부터).")

    def load_specimen_rows_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "표본 데이터 엑셀 불러오기", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            df = self.select_excel_sheet(file_path)
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 불러오기 실패", str(exc))
            return

        if df is None:
            return

        df = self.map_excel_columns(df, SPECIMEN_ROW_COLUMNS, SPECIMEN_FINAL_EXPORT_COLUMNS, "specimen")
        if df is None:
            return

        if self.specimen_data_table.rowCount() > 0:
            reply = QMessageBox.question(self, "엑셀 불러오기", "현재 작성 중인 내용을 지우고 불러올까요?")
            if reply != QMessageBox.Yes:
                return

        self.specimen_autosave_suspended = True
        try:
            self.specimen_data_table.setRowCount(0)
            for _, row in df.iterrows():
                record = {column: self.clean_master_value(row[column]) for column in SPECIMEN_ROW_COLUMNS}
                if not record["데이터베이스번호"] and record.get("이전데이터베이스번호"):
                    record["데이터베이스번호"] = record["이전데이터베이스번호"]
                self.add_specimen_data_row(record)
        finally:
            self.specimen_autosave_suspended = False

        self.refresh_location_validation(self.specimen_data_table, SPECIMEN_ROW_COLUMNS)
        self.autosave_specimen_draft()
        self.set_specimen_save_status(f"엑셀 불러오기 완료: {file_path} ({len(df)}건)")


    @staticmethod
    def write_specimen_excel_file(file_path: str, records: list[dict[str, Any]]):
        pd.DataFrame(records, columns=SPECIMEN_ROW_COLUMNS).to_excel(file_path, index=False)

    @staticmethod
    def write_specimen_final_excel_file(file_path: str, records: list[dict[str, Any]]):
        write_final_template(file_path, records, SPECIMEN_FINAL_EXPORT_COLUMNS, "specimen")

    def sync_specimen_media_filename_field(self, *_):
        value = self.get_form_value("파일경로", self.specimen_form_widgets)
        self.set_form_value("파일명", Path(value).name if value else "", self.specimen_form_widgets)

    def update_specimen_media_thumbnail(self, *_):
        self.render_media_thumbnail(
            self.specimen_media_thumbnail_label, self.get_form_value("파일경로", self.specimen_form_widgets)
        )

    def normalize_specimen_date_field(self, field: str):
        current = self.get_form_value(field, self.specimen_form_widgets)
        self.set_form_value(field, self.normalize_date_text(current), self.specimen_form_widgets)

    def set_specimen_species_status(self, message: str):
        if hasattr(self, "specimen_species_status_label"):
            self.specimen_species_status_label.setText(message)
            self.specimen_species_status_label.setVisible(bool(message.strip()))

    def set_specimen_save_status(self, message: str):
        if hasattr(self, "specimen_save_status_label"):
            self.specimen_save_status_label.setText(message)
            self.specimen_save_status_label.setToolTip(message)

    def _build_project_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        form_layout = QHBoxLayout()
        self.project_year_input = QSpinBox()
        self.project_year_input.setRange(2000, 2100)
        self.project_year_input.setValue(datetime.now().year)
        self.project_order_input = QSpinBox()
        self.project_order_input.setRange(1, 999)
        self.project_name_input = QLineEdit()
        self.project_name_input.setPlaceholderText("사업명")

        add_button = QPushButton("사업 추가")
        add_button.clicked.connect(self.add_project)
        delete_button = QPushButton("선택 사업 삭제")
        delete_button.clicked.connect(self.delete_selected_projects)
        save_button = QPushButton("사업 저장")
        save_button.clicked.connect(self.save_projects_from_table)

        form_layout.addWidget(QLabel("사업년도"))
        form_layout.addWidget(self.project_year_input)
        form_layout.addWidget(QLabel("순번"))
        form_layout.addWidget(self.project_order_input)
        form_layout.addWidget(self.project_name_input, 2)
        form_layout.addWidget(add_button)
        form_layout.addWidget(delete_button)
        form_layout.addWidget(save_button)
        layout.addLayout(form_layout)

        self.project_table = QTableWidget(0, len(PROJECT_COLUMNS))
        self.project_table.setHorizontalHeaderLabels(PROJECT_COLUMNS)
        self.project_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.project_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.project_table)
        self.populate_reference_table(self.project_table, self.projects, PROJECT_COLUMNS)
        return tab

    def _build_organization_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        form_layout = QHBoxLayout()
        self.organization_name_input = QLineEdit()
        self.organization_name_input.setPlaceholderText("기관명")
        self.organization_code_input = QLineEdit()
        self.organization_code_input.setPlaceholderText("기관코드")

        add_button = QPushButton("기관 추가")
        add_button.clicked.connect(self.add_organization)
        delete_button = QPushButton("선택 기관 삭제")
        delete_button.clicked.connect(self.delete_selected_organizations)
        save_button = QPushButton("기관 저장")
        save_button.clicked.connect(self.save_organizations_from_table)

        form_layout.addWidget(self.organization_name_input, 2)
        form_layout.addWidget(self.organization_code_input)
        form_layout.addWidget(add_button)
        form_layout.addWidget(delete_button)
        form_layout.addWidget(save_button)
        layout.addLayout(form_layout)

        self.organization_table = QTableWidget(0, len(ORGANIZATION_COLUMNS))
        self.organization_table.setHorizontalHeaderLabels(ORGANIZATION_COLUMNS)
        self.organization_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.organization_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.organization_table)
        self.populate_reference_table(self.organization_table, self.organizations, ORGANIZATION_COLUMNS)
        return tab

    def _ensure_reference_files(self):
        REFERENCE_DIR.mkdir(parents=True, exist_ok=True)
        if not PROJECTS_FILE.exists():
            self.write_json(PROJECTS_FILE, DEFAULT_PROJECTS)
        if not ORGANIZATIONS_FILE.exists():
            self.write_json(ORGANIZATIONS_FILE, DEFAULT_ORGANIZATIONS)

    def load_reference_data(self):
        self.projects = self.read_json(PROJECTS_FILE, DEFAULT_PROJECTS)
        self.organizations = self.read_json(ORGANIZATIONS_FILE, DEFAULT_ORGANIZATIONS)
        self.rebuild_lookups()

    def load_species_master_lookup(self, dataframe=None):
        self.species_scientific_lookup = {}
        self._species_master_fingerprint = SpeciesMasterService.master_fingerprint()
        try:
            df = self.master_service.load_last_master() if dataframe is None else dataframe
        except Exception:
            self.species_master_df = pd.DataFrame()
            self.species_name_lookup = {}
            return

        required_columns = {"국명", "학명"}
        if not required_columns.issubset(df.columns):
            self.species_master_df = pd.DataFrame()
            self.species_name_lookup = {}
            return

        self.species_master_df = df
        self.species_scientific_lookup = {}
        lookup: dict[str, list[dict[str, Any]]] = {}
        for record in df.to_dict("records"):
            scientific_key = self.scientific_name_key(record.get("학명"))
            if scientific_key:
                self.species_scientific_lookup.setdefault(scientific_key, []).append(record)
            korean_name = normalize_text(record.get("국명"))
            if not korean_name:
                continue
            lookup.setdefault(korean_name, []).append(record)
        self.species_name_lookup = lookup
        self._species_master_fingerprint = SpeciesMasterService.master_fingerprint()

    def ensure_species_master_current(self):
        if getattr(self, "_species_master_fingerprint", None) != SpeciesMasterService.master_fingerprint():
            self.load_species_master_lookup()

    def find_species_matches(self, korean, scientific, prefer_scientific=False):
        latin_matches = getattr(self, "species_scientific_lookup", {}).get(self.scientific_name_key(scientific), [])
        if prefer_scientific:
            return latin_matches
        return self.species_name_lookup.get(korean, []) if korean else latin_matches

    def rebuild_lookups(self):
        self.project_lookup = {self.project_label(project): project for project in self.projects}
        self.organization_lookup = {self.organization_label(org): org for org in self.organizations}

    def refresh_combos(self):
        self.rebuild_lookups()

        if "사업명" in self.form_widgets and "기관명" in self.form_widgets:
            self.refresh_combo_items("사업명", list(self.project_lookup.keys()))
            self.refresh_combo_items("기관명", list(self.organization_lookup.keys()))
            self.set_default_organization()
            self.apply_project_to_form()
            self.apply_organization_to_form()

        if "사업명" in self.specimen_form_widgets and "기관명" in self.specimen_form_widgets:
            self.refresh_combo_items("사업명", list(self.project_lookup.keys()), self.specimen_form_widgets)
            self.refresh_combo_items("기관명", list(self.organization_lookup.keys()), self.specimen_form_widgets)
            self.set_default_organization(self.specimen_form_widgets)
            self.apply_specimen_project_to_form()
            self.apply_specimen_organization_to_form()

    def refresh_combo_items(self, column_name: str, labels: list[str], widgets: dict[str, Any] | None = None):
        widgets = widgets if widgets is not None else self.form_widgets
        widget = widgets.get(column_name)
        if not isinstance(widget, QComboBox):
            return

        selected_value = self.get_form_value(column_name, widgets) if widget.count() else ""
        widget.blockSignals(True)
        widget.clear()
        widget.addItems(labels)
        selected_index = self.find_combo_index(widget, selected_value, column_name)
        if selected_index >= 0:
            widget.setCurrentIndex(selected_index)
        widget.blockSignals(False)

    def apply_project_to_form(self, label: str = "", widgets: dict[str, Any] | None = None, on_change=None):
        widgets = widgets if widgets is not None else self.form_widgets
        if not label and isinstance(widgets.get("사업명"), QComboBox):
            label = widgets["사업명"].currentText()
        project = self.project_lookup.get(label, {})
        if not project:
            return
        self.set_form_value("사업년도", project.get("사업년도", datetime.now().year), widgets)
        self.set_form_value("순번", project.get("순번", 0), widgets)
        (on_change or self.refresh_generated_fields)()

    def apply_organization_to_form(self, label: str = "", widgets: dict[str, Any] | None = None, on_change=None):
        widgets = widgets if widgets is not None else self.form_widgets
        if not label and isinstance(widgets.get("기관명"), QComboBox):
            label = widgets["기관명"].currentText()
        organization = self.organization_lookup.get(label, {})
        if not organization:
            return
        self.set_form_value("기관코드", organization.get("기관코드", ""), widgets)
        (on_change or self.refresh_generated_fields)()

    def set_default_organization(self, widgets: dict[str, Any] | None = None):
        widget = (widgets if widgets is not None else self.form_widgets).get("기관명")
        if not isinstance(widget, QComboBox):
            return
        for row in range(widget.count()):
            if widget.itemText(row).startswith(f"{DEFAULT_ORGANIZATION_NAME} ("):
                widget.setCurrentIndex(row)
                return

    def refresh_generated_fields(self, *, force_number=False):
        if self.autosave_suspended:
            return
        if not self.get_form_value("사업별 순번", self.form_widgets):
            self.set_form_value("사업별 순번", self.project_sequence_for_form(False), self.form_widgets)
        media_type = self.get_form_value("미디어타입")
        media_code = MEDIA_TYPE_CODES.get(media_type, "")
        self.set_form_value("미디어코드", media_code)
        if not force_number and self.restore_existing_database_number(False):
            return

        organization_code = self.get_form_value("기관코드")
        media_code = self.get_form_value("미디어코드")
        taxon_code = self.get_form_value("분류코드")
        year = self.get_form_value("사업년도")
        order = self.get_form_value("순번")
        no = self.get_form_value("사업별 순번")
        if organization_code and media_code and taxon_code and year and order and str(no).isdigit():
            year_text = str(year)[-2:].zfill(2)
            self.set_form_value(
                "데이터베이스번호",
                f"{organization_code}-{media_code}-{taxon_code}-{year_text}{int(order):02d}{int(no):05d}",
            )
        else:
            self.set_form_value("데이터베이스번호", "")

    def add_data_row_from_form(self):
        self.set_form_value("NO.", self.next_no_value())
        values = self.collect_form_values()
        values = self.prepare_new_project_number(values, False)
        if not self.validate_form_values(values):
            return
        values = self.prepare_media_for_record(values)
        self.add_data_row(values)
        self.prepare_next_input(values)
        self.autosave_draft()

    def update_selected_data_row(self):
        row = self.current_data_row()
        if row is None:
            QMessageBox.warning(self, "선택 행 수정", "수정할 행을 먼저 선택해 주세요.")
            return
        values = self.collect_form_values()
        if not self.validate_form_values(values):
            return
        values = self.prepare_media_for_record(values)
        self.write_data_row(row, values)
        self.autosave_draft()
        self.set_save_status("선택 행을 수정했습니다.")

    def load_selected_row_to_form(self):
        previous_row = self._last_selected_row
        row = self.current_data_row()

        # 선택 행이 바뀌는 시점에만 이전에 편집하던 폼 내용을 표에 반영
        # (타이핑 중에는 반영하지 않음 -> 큰 표에서도 입력이 즉시 반응)
        if previous_row is not None and previous_row != row:
            values = self.collect_form_values()
            if self.has_form_content(values):
                self.write_data_row(previous_row, values)

        self._last_selected_row = row
        if row is None:
            return
        self.autosave_suspended = True
        for col, column_name in enumerate(ROW_COLUMNS):
            item = self.data_table.item(row, col)
            text = item.text() if item else ""
            # 사업명 등 세션 유지 값은 그 행에 값이 없으면 지금 폼에 있는 값을 그대로 둠
            # (예: 첫 행에서 고른 사업명이 다음 빈 행에도 이어짐)
            if column_name in STICKY_FIELDS and not text:
                continue
            self.set_form_value(column_name, text)
        self.autosave_suspended = False

    def on_data_table_item_changed(self, item: QTableWidgetItem):
        if self.autosave_suspended:
            return
        column_name = ROW_COLUMNS[item.column()]
        if column_name == "사업별 순번":
            self.renumber_project_sequences_below(self.data_table, ROW_COLUMNS, item.row(), False)
        if column_name in {"사업년도", "순번"}:
            self.recalculate_project_sequences(self.data_table, ROW_COLUMNS)
        if column_name == "데이터베이스번호":
            self.mark_database_number_manual(self.data_table, item)
        if item.column() == ROW_COLUMNS.index("NO."):
            self.renumber_rows_below(item.row())
        if ROW_COLUMNS[item.column()] in {"위도", "경도"}:
            self.revalidate_row_location(self.data_table, ROW_COLUMNS, item.row())
        if item.row() == self.current_data_row():
            self.load_selected_row_to_form()
            if column_name in {"국명", "학명"}:
                # 표에서 국명 셀을 직접 고쳤을 때도 폼에서 편집할 때와 동일하게
                # 학명/종고유ID를 조회해서 채우고, 그 결과를 다시 표에 반영
                self.fill_species_from_korean_name(prefer_scientific=column_name == "학명")
                self.write_data_row(item.row(), self.collect_form_values())
            elif column_name in DATE_FIELDS:
                self.normalize_date_field(column_name)
                self.write_data_row(item.row(), self.collect_form_values())
        self.schedule_autosave()

    def renumber_rows_below(self, changed_row: int):
        no_column = ROW_COLUMNS.index("NO.")
        item = self.data_table.item(changed_row, no_column)
        try:
            next_no = int(item.text() if item else "") + 1
        except ValueError:
            return

        was_blocked = self.data_table.blockSignals(True)
        try:
            for row in range(changed_row + 1, self.data_table.rowCount()):
                self.data_table.setItem(row, no_column, QTableWidgetItem(str(next_no)))
                next_no += 1
        finally:
            self.data_table.blockSignals(was_blocked)

    @staticmethod
    def default_field_value(column_name: str) -> Any:
        if column_name == "관찰개체수":
            return 1
        if column_name in DROPDOWN_DEFAULTS:
            return DROPDOWN_DEFAULTS[column_name]
        if column_name in TEXT_DEFAULTS:
            return TEXT_DEFAULTS[column_name]
        if column_name in DROPDOWN_OPTIONS and DROPDOWN_OPTIONS[column_name]:
            # set_form_value가 콤보에서 ""와 일치하는 항목을 못 찾으면 첫 항목으로
            # 떨어지는(findText 실패 -> index 0) 기존 동작과 맞춰야 함
            return DROPDOWN_OPTIONS[column_name][0]
        return ""

    def clear_form(self):
        self._last_selected_row = None
        blocked = self.data_table.blockSignals(True)
        try:
            self.data_table.clearSelection()
        finally:
            self.data_table.blockSignals(blocked)
        for column_name in ROW_COLUMNS:
            if column_name in STICKY_FIELDS:
                continue
            if column_name == "NO.":
                self.set_form_value(column_name, self.next_no_value())
            else:
                self.set_form_value(column_name, self.default_field_value(column_name))
        self.apply_project_to_form()
        self.set_default_organization()
        self.apply_organization_to_form()
        self.set_form_value("사업별 순번", self.project_sequence_for_form(False, new=True))
        self.refresh_generated_fields(force_number=True)
        self.schedule_autosave()

    def collect_form_values(self) -> dict[str, Any]:
        self.refresh_generated_fields()
        return {column_name: self.get_form_value(column_name) for column_name in ROW_COLUMNS}

    def validate_form_values(self, values: dict[str, Any]) -> bool:
        required_fields = [
            "NO.",
            "미디어타입",
            "미디어코드",
            "분류코드",
            "사업명",
            "사업년도",
            "순번",
            "사업별 순번",
            "기관명",
            "기관코드",
            "국명",
            "학명",
        ]
        if values.get("미디어타입") == "동영상":
            required_fields += VIDEO_FIELDS
        elif values.get("미디어타입") == "사운드":
            required_fields += SOUND_FIELDS
        missing = [field for field in required_fields if values.get(field) in ("", 0)]
        if missing:
            QMessageBox.warning(self, "입력 확인", f"필수값을 입력해 주세요: {', '.join(missing)}")
            return False
        return True

    def prepare_next_input(self, values: dict[str, Any]):
        # NO.만 올리고 나머지(국명/학명/파일경로 등)를 그대로 두면, "입력" 이후 폼에 남아있는
        # 직전 값이 다음 저장(임시저장/엑셀저장) 때 새 행으로 중복 추가되는 문제가 있었음
        # -> "새 입력"과 동일하게 폼을 완전히 리셋
        self.clear_form()

    def next_no_value(self) -> int:
        max_no = 0
        no_column = ROW_COLUMNS.index("NO.")
        for row in range(self.data_table.rowCount()):
            item = self.data_table.item(row, no_column)
            max_no = max(max_no, self.safe_sort_int(item.text() if item else 0))
        return max_no + 1

    def fill_species_from_korean_name(self, prefer_scientific=False):
        self.ensure_species_master_current()
        korean_name = normalize_text(self.get_form_value("국명"))
        scientific_input = normalize_text(self.get_form_value("학명"))
        if not korean_name and not scientific_input:
            self.set_species_status("")
            return

        if self.species_master_df.empty:
            self.set_species_status("종정보 마스터 없음")
            return

        matches = self.find_species_matches(korean_name, scientific_input, prefer_scientific)
        if not matches:
            self.set_species_status("마스터 매칭 없음")
            return

        match = matches[0]
        self.set_form_value("국명", normalize_text(match.get("국명")))
        scientific_name = normalize_text(match.get("학명"))
        species_id = self.clean_master_value(match.get("종고유ID"))
        if species_id in {"", "0"}:
            species_id = self.clean_master_value(match.get("종학명정보ID"))

        self.set_form_value("학명", scientific_name)
        self.set_form_value("종고유ID", species_id)

        if len(matches) > 1:
            self.set_species_status(f"중복 {len(matches)}건, 첫 항목 적용")
        else:
            self.set_species_status("마스터 적용 완료")

    def fill_address_english(self, korean_field: str):
        english_field, level = ADDRESS_FIELDS[korean_field]
        korean_text = normalize_text(self.get_form_value(korean_field))
        if not korean_text:
            return

        english_text, exact = self.address_master.match_english(level, korean_text)
        if not english_text:
            self.set_save_status(f"{korean_field}: 주소 마스터에 없음 (직접 입력 필요)")
            return

        self.set_form_value(english_field, english_text)
        if exact:
            self.set_save_status(f"{korean_field} → {english_text} 자동 적용")
        else:
            self.set_save_status(f"{korean_field} → {english_text} (유사 매칭, 확인 필요)")

    def open_location_on_map(self):
        self.open_location_on_map_for(self.get_form_value("위도"), self.get_form_value("경도"))

    def open_all_locations_on_map(self):
        self.open_all_locations_on_map_for(self.data_table, ROW_COLUMNS, ADDRESS_FIELDS, "map_preview.html")

    def fill_all_addresses_from_latlon(self):
        self.batch_fill_addresses(self.data_table, ROW_COLUMNS, ADDRESS_FIELDS, self.set_save_status)

    def fill_all_species_and_project(self):
        self.batch_fill_species_and_project(
            self.data_table, ROW_COLUMNS, self.form_widgets, self.set_save_status, self.schedule_autosave
        )

    @staticmethod
    def scientific_name_key(name: Any) -> str:
        # "Aegypius monachus (Linnaeus, 1766)" → "aegypius monachus" (명명자/연도 차이 무시)
        return " ".join(normalize_text(name).split()[:2]).lower()

    def batch_fill_species_and_project(
        self, table: QTableWidget, row_columns: list[str], widgets: dict[str, Any], status_setter, after_change
    ):
        """설정 창에서 선택한 종정보/사업/미디어를 전체 행에 적용."""
        self.ensure_species_master_current()
        row_count = table.rowCount()
        if row_count == 0:
            QMessageBox.information(self, "종ID·사업 일괄 적용", "적용할 행이 없습니다.")
            return
        dialog = BatchApplyDialog(
            self.projects, MEDIA_TYPE_CODES if "미디어타입" in row_columns else {}, row_count, self
        )
        if dialog.exec() != QDialog.Accepted:
            return
        options = dialog.values()
        start_number = options.get("database_start_number", 1)
        if start_number + row_count - 1 > 99999:
            QMessageBox.warning(self, "DB 번호 생성", "시작 순번과 전체 행 수의 합이 5자리 범위(99999)를 초과합니다. 시작 순번을 낮춰 주세요.")
            return
        project = options["project"]
        media_type = options["media_type"]
        taxon_code = options.get("taxon_code", "")

        scientific_lookup: dict[str, dict[str, Any]] = {}
        for record in self.species_master_df.to_dict("records") if not self.species_master_df.empty else []:
            key = self.scientific_name_key(record.get("학명"))
            if key:
                scientific_lookup.setdefault(key, record)

        col = {name: index for index, name in enumerate(row_columns)}

        def cell(row: int, name: str) -> str:
            item = table.item(row, col[name])
            return item.text().strip() if item else ""

        def put(row: int, name: str, value: Any):
            table.setItem(row, col[name], QTableWidgetItem("" if value is None else str(value)))

        matched = unmatched = 0
        project_counts = {}
        was_blocked = table.blockSignals(True)
        table.setUpdatesEnabled(False)
        try:
            for row in range(row_count):
                if project:
                    put(row, "사업명", project.get("사업명", ""))
                    put(row, "사업년도", project.get("사업년도", ""))
                    put(row, "순번", project.get("순번", ""))

                if media_type:
                    put(row, "미디어타입", media_type)
                    put(row, "미디어코드", MEDIA_TYPE_CODES[media_type])
                if taxon_code:
                    put(row, "분류코드", taxon_code)
                group = (cell(row, "사업년도"), cell(row, "순번"))
                project_counts[group] = project_counts.get(group, 0) + 1
                put(row, "사업별 순번", start_number + project_counts[group] - 1)

                if (project or media_type or taxon_code or start_number != 1 or not options.get("preserve_database_numbers", True)) and not (options.get("preserve_database_numbers", True) and cell(row, "데이터베이스번호")):
                    parts = [cell(row, name) for name in ("기관코드", "분류코드", "사업년도", "순번", "사업별 순번")]
                    code = cell(row, "미디어코드") if "미디어코드" in col else "SP"
                    number = ""
                    if all(parts) and code:
                        try:
                            org, taxon, year, order, no = parts
                            number = f"{org}-{code}-{taxon}-{year[-2:].zfill(2)}{int(order):02d}{int(no):05d}"
                        except ValueError:
                            pass
                    put(row, "데이터베이스번호", number)

                if not options["species"]:
                    continue

                korean_name = normalize_text(cell(row, "국명"))
                scientific_name = cell(row, "학명")
                matches = self.species_name_lookup.get(korean_name, []) if korean_name else []
                match = matches[0] if matches else scientific_lookup.get(self.scientific_name_key(scientific_name) or None)
                if not match:
                    if korean_name or scientific_name:
                        unmatched += 1
                    continue
                species_id = self.clean_master_value(match.get("종고유ID"))
                if species_id in {"", "0"}:
                    species_id = self.clean_master_value(match.get("종학명정보ID"))
                put(row, "종고유ID", species_id)
                put(row, "학명", normalize_text(match.get("학명")))
                if not korean_name:
                    put(row, "국명", normalize_text(match.get("국명")))
                if "속명" in col:  # 표본: 속명/종명도 마스터 값으로
                    put(row, "속명", normalize_text(match.get("속명국문명")))
                    put(row, "종명", normalize_text(match.get("종소명국문명")))
                matched += 1
        finally:
            table.setUpdatesEnabled(True)
            table.blockSignals(was_blocked)

        if table is self.data_table:
            self.load_selected_row_to_form()
        else:
            self.load_selected_specimen_row_to_form()
        after_change()
        unmatched_note = f", 마스터에 없음 {unmatched}행" if unmatched else ""
        status_setter(f"일괄 적용 완료: 전체 {row_count}행, 종 매칭 {matched}행{unmatched_note}")

    def fill_address_from_latlon(self):
        latitude = self.to_float(self.get_form_value("위도"))
        longitude = self.to_float(self.get_form_value("경도"))
        if latitude is None or longitude is None:
            QMessageBox.warning(self, "주소 조회", "위도/경도를 먼저 입력해 주세요.")
            return

        try:
            address_ko = self.reverse_geocode(latitude, longitude, "ko")
            address_en = self.reverse_geocode(latitude, longitude, "en")
        except Exception as exc:
            self.set_save_status(f"주소 조회 실패: {exc}")
            return

        if not address_ko:
            self.set_save_status("주소 조회 결과가 없습니다.")
            return

        for korean_field, english_field, key in (
            ("관찰위치 시도 국문", "관찰위치 시도 영문", "sido"),
            ("관찰위치 시군구 국문", "관찰위치 시군구 영문", "sigungu"),
            ("관찰위치 읍면동 국문", "관찰위치 읍면동 영문", "eupmyeondong"),
        ):
            korean_value = address_ko.get(key, "")
            if not korean_value:
                continue
            self.set_form_value(korean_field, korean_value)

            # API가 영문명을 따로 주면 그걸 바로 쓰고, 못 주면(=한글을 그대로 돌려주면)
            # 기존 주소 마스터 국->영 매칭으로 대체 채움
            english_value = address_en.get(key, "")
            if english_value and english_value != korean_value:
                self.set_form_value(english_field, english_value)
            else:
                self.fill_address_english(korean_field)

        self.set_save_status("위도/경도 기반 주소를 조회했습니다. (OpenStreetMap Nominatim, 내용 확인 필요)")

    @staticmethod
    def reverse_geocode(latitude: float, longitude: float, language: str = "ko") -> dict[str, str]:
        # ponytail: Nominatim 무료 API(키 불필요, 초당 1건 제한). 정확도가 중요하면
        # 공공데이터포털/브이월드 좌표->주소 API(키 발급 필요)로 교체.
        params = urllib.parse.urlencode(
            {
                "format": "jsonv2",
                "lat": latitude,
                "lon": longitude,
                "zoom": 14,
                "addressdetails": 1,
                "accept-language": language,
            }
        )
        request = urllib.request.Request(
            f"https://nominatim.openstreetmap.org/reverse?{params}",
            headers={"User-Agent": "BioFlow-species-upload/1.0"},
        )
        with urllib.request.urlopen(request, timeout=10) as response:
            data = json.loads(response.read().decode("utf-8"))

        address = data.get("address", {})
        return {
            "sido": normalize_text(address.get("state") or address.get("province")),
            "sigungu": normalize_text(
                address.get("city") or address.get("county") or address.get("borough")
            ),
            "eupmyeondong": normalize_text(
                address.get("town") or address.get("suburb") or address.get("village") or address.get("quarter")
            ),
        }

    def import_media_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "미디어 폴더 선택", "")
        if not folder:
            return

        media_files = self.scan_media_folder(Path(folder))
        if not media_files:
            QMessageBox.information(self, "폴더 불러오기", "지원하는 이미지/동영상/사운드 파일을 찾지 못했습니다.")
            return

        progress = QProgressDialog("미디어 파일 정보를 읽는 중...", "취소", 0, len(media_files), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(300)

        added_count = 0
        self.autosave_suspended = True
        try:
            # 폼에 이전에 작성하던 국명/학명 등 내용이 남아있으면 새로 만드는 모든 행에
            # 그대로 찍혀버리므로, 사업명/기관/분류코드/미디어타입만 남기고 비운 뒤 시작
            self.clear_form()
            for index, (file_path, media_type) in enumerate(media_files):
                progress.setLabelText(f"{file_path.name} 처리 중... ({index + 1}/{len(media_files)})")
                progress.setValue(index)
                # subprocess 호출로 화면이 멈춰 보이지 않도록 파일 사이사이 이벤트를 처리
                QApplication.processEvents()
                if progress.wasCanceled():
                    break

                self.set_form_value("미디어타입", media_type)
                self.set_form_value("NO.", self.next_no_value())
                self.set_form_value("파일경로", str(file_path))
                self.set_form_value("원본경로", str(file_path))
                self.set_form_value("원본파일명", file_path.stem)
                # 동영상/사운드는 위치/재생시간/촬영일자가 모두 같은 ffprobe 출력에서 나오므로
                # 파일당 세 번 부르는 대신 한 번만 불러 공유 (exiftool 기동 비용이 커서 체감 효과 큼)
                ffprobe_format = (
                    self.probe_ffprobe_format(file_path) if media_type in {"동영상", "사운드"} else None
                )
                self.apply_location_from_media(file_path, media_type, ffprobe_format)
                self.apply_duration_from_media(file_path, media_type, ffprobe_format)
                self.apply_date_from_media(file_path, media_type, ffprobe_format)
                self.add_data_row(self.collect_form_values())
                added_count += 1
        finally:
            self.autosave_suspended = False
            progress.setValue(len(media_files))

        self.refresh_location_validation(self.data_table, ROW_COLUMNS)
        self.autosave_draft()
        cancel_note = "\n(취소되어 일부만 불러왔습니다.)" if added_count < len(media_files) else ""
        QMessageBox.information(
            self,
            "폴더 불러오기",
            f"{added_count}건의 미디어를 행으로 추가했습니다.{cancel_note}\n"
            "각 행을 선택해 종명 등 나머지 정보를 입력한 뒤 '선택 행 수정'으로 저장해 주세요.",
        )
        self.set_save_status(f"폴더 불러오기 완료: {added_count}건 (나머지 정보 입력 필요)")

    @staticmethod
    def scan_media_folder(folder: Path) -> list[tuple[Path, str]]:
        media_files = []
        for path in sorted(folder.iterdir()):
            if not path.is_file():
                continue
            media_type = MEDIA_EXTENSION_TYPES.get(path.suffix.lower())
            if media_type:
                media_files.append((path, media_type))
        return media_files

    def select_media_source_file(self):
        media_type = self.get_form_value("미디어타입")
        file_filter = self.media_file_filter(media_type)
        file_path, _ = QFileDialog.getOpenFileName(self, "원본 미디어 선택", "", file_filter)
        if not file_path:
            return

        source_path = Path(file_path)
        # 미디어타입은 세션 내내 유지되는 값(STICKY_FIELDS)이라, 이전 행에서 "동영상"을
        # 고른 채로 남아있으면 실제로 고른 파일이 이미지여도 안 바뀜 -> 저장 시 엉뚱하게
        # 동영상 필수값(제목/재생시간 등) 검증에 걸려서 조용히 저장이 안 되던 원인.
        # 고른 파일 확장자로 다시 맞춰줌.
        detected_type = MEDIA_EXTENSION_TYPES.get(source_path.suffix.lower())
        if detected_type and detected_type != media_type:
            media_type = detected_type
            self.set_form_value("미디어타입", media_type)

        self.set_form_value("파일경로", file_path)
        # 새로 고른 원본 파일 정보를 바로 기록 (저장 시 파일경로/파일명만 새 저장위치로 바뀜)
        self.set_form_value("원본경로", file_path)
        self.set_form_value("원본파일명", source_path.stem)
        self.apply_location_from_media(source_path, media_type)
        self.apply_duration_from_media(source_path, media_type)
        self.apply_date_from_media(source_path, media_type)
        self.schedule_autosave()

    def apply_location_from_media(
        self, source_path: Path, media_type: str, ffprobe_format: dict | None = None
    ):
        latitude, longitude, message = self.extract_location_from_media(source_path, media_type, ffprobe_format)
        if latitude is None or longitude is None:
            self.set_save_status(message)
            return

        self.set_form_value("위도", f"{latitude:.7f}")
        self.set_form_value("경도", f"{longitude:.7f}")
        self.set_save_status(f"미디어 위치정보 적용: 위도 {latitude:.7f}, 경도 {longitude:.7f}")

    def apply_duration_from_media(
        self, source_path: Path, media_type: str, ffprobe_format: dict | None = None
    ):
        duration_field = {"동영상": "동영상 재생시간", "사운드": "사운드 재생시간"}.get(media_type)
        if not duration_field:
            return
        duration_text = self.extract_media_duration(source_path, ffprobe_format)
        if duration_text:
            self.set_form_value(duration_field, duration_text)

    @staticmethod
    def run_json_command(command: list[str], timeout: int, empty_json: str) -> Any:
        # ponytail: ffprobe/exiftool emit UTF-8 JSON; on Korean Windows the platform
        # default (cp949) can't decode it, so encoding must be pinned here explicitly.
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                check=False,
            )
            return json.loads(result.stdout or empty_json)
        except Exception:
            return json.loads(empty_json)

    @staticmethod
    def probe_ffprobe_format(file_path: Path) -> dict:
        # 위치/재생시간/촬영일자가 모두 이 하나의 ffprobe 출력에서 나오므로, 호출부에서
        # 한 번만 실행해 공유하면 파일당 ffprobe 프로세스 기동을 최대 3->1회로 줄일 수 있음.
        ffprobe = find_media_tool("ffprobe")
        if not ffprobe:
            return {}
        data = UploadTab.run_json_command(
            [ffprobe, "-v", "quiet", "-print_format", "json", "-show_format", str(file_path)],
            timeout=10,
            empty_json="{}",
        )
        return data.get("format", {})

    @staticmethod
    def extract_media_duration(file_path: Path, ffprobe_format: dict | None = None) -> str:
        media_format = ffprobe_format if ffprobe_format is not None else UploadTab.probe_ffprobe_format(file_path)
        duration = UploadTab.to_float(media_format.get("duration"))
        if duration is None:
            return ""

        total_seconds = int(round(duration))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours:
            return f"{hours}:{minutes:02d}:{seconds:02d}"
        return f"{minutes}:{seconds:02d}"

    def apply_date_from_media(self, source_path: Path, media_type: str, ffprobe_format: dict | None = None):
        date_text = self.extract_media_date(source_path, media_type, ffprobe_format)
        if date_text:
            self.set_form_value("관찰일", date_text)

    @staticmethod
    def extract_media_date(file_path: Path, media_type: str, ffprobe_format: dict | None = None) -> str:
        if media_type == "이미지":
            date_text = UploadTab.extract_image_date(file_path)
            if date_text:
                return date_text
        return UploadTab.extract_media_date_with_ffprobe(file_path, ffprobe_format)

    @staticmethod
    def extract_image_date(file_path: Path) -> str:
        if Image is None or ExifTags is None:
            return ""
        try:
            with Image.open(file_path) as image:
                exif = image.getexif()
                raw = exif.get(36867) or exif.get(306)  # DateTimeOriginal / DateTime
        except Exception:
            return ""
        if not raw:
            return ""
        date_part = str(raw).split(" ")[0]  # "YYYY:MM:DD HH:MM:SS" -> "YYYY:MM:DD"
        if date_part.count(":") != 2:
            return ""
        return date_part.replace(":", "-")

    @staticmethod
    def extract_media_date_with_ffprobe(file_path: Path, ffprobe_format: dict | None = None) -> str:
        media_format = ffprobe_format if ffprobe_format is not None else UploadTab.probe_ffprobe_format(file_path)
        creation_time = media_format.get("tags", {}).get("creation_time")
        if not creation_time:
            return ""
        return str(creation_time).split("T")[0]

    def prepare_media_for_record(self, values: dict[str, Any]) -> dict[str, Any]:
        if getattr(self, "_defer_media_copy", False):
            return values
        media_path_text = normalize_text(values.get("파일경로"))
        if not media_path_text:
            return values

        source_path = Path(media_path_text)
        if not source_path.exists():
            return values

        saved_path = self.save_media_file_to_storage(source_path, values)
        if saved_path:
            values = values.copy()
            if not normalize_text(values.get("원본경로")):
                # 최초 저장 시 1번만 기록 -> 재저장해도 진짜 원본 경로가 덮어써지지 않음
                values["원본경로"] = str(source_path)
                values["원본파일명"] = source_path.stem
                self.set_form_value("원본경로", str(source_path))
                self.set_form_value("원본파일명", source_path.stem)
            values["파일경로"] = str(saved_path)
            values["파일명"] = saved_path.name
            self.set_form_value("파일경로", str(saved_path))
            self.apply_location_from_media(saved_path, values.get("미디어타입", ""))
            values["위도"] = self.get_form_value("위도")
            values["경도"] = self.get_form_value("경도")

            filename_field = {"동영상": "동영상 파일명", "사운드": "사운드 파일명"}.get(values.get("미디어타입"))
            if filename_field:
                values[filename_field] = saved_path.name
                self.set_form_value(filename_field, saved_path.name)
        return values

    def apply_media_copy_to_row_values(self, values: dict[str, Any]) -> dict[str, Any]:
        # prepare_media_for_record와 같은 복사 로직이지만 폼 위젯을 건드리지 않음
        # (현재 선택된 행이 아닌 다른 행을 일괄 처리할 때 쓰기 위함)
        media_path_text = normalize_text(values.get("파일경로"))
        if not media_path_text:
            return values

        source_path = Path(media_path_text)
        if not source_path.is_file():
            source_path = Path(normalize_text(values.get("원본경로")))
            if not source_path.is_file():
                return values

        saved_path = self.save_media_file_to_storage(source_path, values)
        if not saved_path:
            return values

        values = values.copy()
        if not normalize_text(values.get("원본경로")):
            values["원본경로"] = str(source_path)
            values["원본파일명"] = source_path.stem
        values["파일경로"] = str(saved_path)
        values["파일명"] = saved_path.name
        latitude, longitude, _ = self.extract_location_from_media(saved_path, values.get("미디어타입", ""))
        if latitude is not None and longitude is not None:
            values["위도"] = f"{latitude:.7f}"
            values["경도"] = f"{longitude:.7f}"
        filename_field = {"동영상": "동영상 파일명", "사운드": "사운드 파일명"}.get(values.get("미디어타입"))
        if filename_field:
            values[filename_field] = saved_path.name
        return values

    def copy_pending_media_for_all_rows(self, strict=False):
        # "입력"/"선택 행 수정"으로 개별 저장한 행만 저장 위치로 복사되고 나머지 행(예:
        # 미디어 폴더로 한 번에 불러온 뒤 아직 손 안 댄 행)은 원본 경로 그대로 남아있던 문제 ->
        # 임시저장/엑셀저장 시 모든 행을 훑어서 아직 안 옮겨진 미디어까지 마저 복사
        for row in range(self.data_table.rowCount()):
            values = {
                column: (self.data_table.item(row, col).text() if self.data_table.item(row, col) else "")
                for col, column in enumerate(ROW_COLUMNS)
            }
            updated = self.apply_media_copy_to_row_values(values)
            if strict and normalize_text(values.get("파일경로")) and updated is values:
                return False
            if updated is not values:
                self.write_data_row(row, updated)

        return True

    def prepare_media_for_excel_export(self, specimen, capture=True, choose_destination=False):
        table = self.specimen_data_table if specimen else self.data_table
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        media_rows = []
        for row in range(table.rowCount()):
            values = {field: table.item(row, col).text() if table.item(row, col) else ""
                      for col, field in enumerate(columns)}
            if not normalize_text(values.get("파일경로")):
                continue
            db_number = normalize_text(values.get("데이터베이스번호"))
            if not db_number:
                raise ValueError(f"{row + 1}행: 미디어 저장에 필요한 DB 번호가 없습니다.")
            if self.safe_filename(db_number) != db_number:
                raise ValueError(f"{row + 1}행: DB 번호에 파일명으로 사용할 수 없는 문자가 있습니다.")
            source = Path(values["파일경로"])
            original = Path(normalize_text(values.get("원본경로")))
            if not source.is_file() and not original.is_file():
                raise FileNotFoundError(f"{row + 1}행: 원본 미디어 파일을 찾을 수 없습니다.\n{source}")
            if not specimen:
                media_destination(Path("."), values)
            media_rows.append(row)
        if not media_rows:
            return True
        if choose_destination:
            previous = self.specimen_media_storage_dir if specimen else self.media_storage_dir
            title = "이번 엑셀의 이미지 저장 위치 선택" if specimen else "이번 엑셀의 미디어 저장 위치 선택"
            selected = QFileDialog.getExistingDirectory(self, title, previous or "")
            if not selected:
                return False
            if specimen:
                self.specimen_media_storage_dir = selected
                self.specimen_media_storage_path_input.setText(selected)
            else:
                self.media_storage_dir = selected
                self.media_storage_path_input.setText(selected)
        ensure = self.ensure_specimen_media_storage_dir if specimen else self.ensure_media_storage_dir
        destination = ensure()
        if destination is None:
            return False
        copier = self.copy_pending_media_for_all_specimen_rows if specimen else self.copy_pending_media_for_all_rows
        if not copier(strict=True):
            QMessageBox.warning(self, "엑셀 저장 중단", "미디어 파일을 DB 번호로 저장하지 못했습니다. 저장 위치와 DB 번호를 확인해 주세요. 작성 내용은 유지됩니다.")
            return False
        for row in media_rows:
            path = Path(table.item(row, columns.index("파일경로")).text())
            db_number = table.item(row, columns.index("데이터베이스번호")).text().strip()
            values = {field: table.item(row, col).text() if table.item(row, col) else ""
                      for col, field in enumerate(columns)}
            expected_dir = destination if specimen else media_destination(destination, values)
            if not path.is_file() or path.parent.resolve() != expected_dir.resolve() or path.stem != db_number:
                raise ValueError(f"{row + 1}행: 미디어 저장 경로 또는 파일명이 DB 번호와 일치하지 않습니다.\n{path}")
        if not specimen and capture:
            _, _, failed, _ = self.capture_pending_video_thumbnails()
            if failed:
                raise RuntimeError(f"동영상 캡처 {failed}건을 생성하지 못했습니다. ffmpeg 설치와 동영상 파일을 확인해 주세요. 작성 내용은 유지됩니다.")
        return True

    def select_media_storage_dir(self):
        selected_dir = QFileDialog.getExistingDirectory(self, "미디어 저장 위치 선택", self.media_storage_dir or "")
        if not selected_dir:
            return
        self.media_storage_dir = selected_dir
        self.media_storage_path_input.setText(selected_dir)
        self.autosave_draft()

    def ensure_media_storage_dir(self) -> Path | None:
        if self.media_storage_dir:
            return Path(self.media_storage_dir)

        selected_dir = QFileDialog.getExistingDirectory(self, "미디어 저장 위치 선택", "")
        if not selected_dir:
            return None
        self.media_storage_dir = selected_dir
        self.media_storage_path_input.setText(selected_dir)
        return Path(selected_dir)

    def save_media_file_to_storage(self, source_path: Path, values: dict[str, Any] | None = None) -> Path | None:
        values = values or self.collect_form_values()
        db_number = normalize_text(values.get("데이터베이스번호"))
        if not db_number:
            QMessageBox.warning(self, "미디어 저장", "DB번호를 먼저 만들 수 있도록 미디어타입, 분류코드, 사업, 기관 값이 필요합니다.")
            return None

        target_dir = self.ensure_media_storage_dir()
        if target_dir is None:
            self.set_save_status("미디어 저장 위치 선택이 취소되었습니다.")
            return None

        try:
            target_dir = media_destination(target_dir, values)
        except ValueError as exc:
            QMessageBox.warning(self, "동영상 저장", str(exc))
            return None

        expected_stem = self.safe_filename(db_number)
        same_folder = source_path.parent.resolve() == target_dir.resolve()
        if same_folder and source_path.stem == expected_stem:
            # 이미 이 폴더에 DB번호 이름으로 저장돼 있는 파일 -> 복사할 필요 없음
            return source_path
        original_path = normalize_text(values.get("원본경로"))
        already_copied = bool(original_path and Path(original_path).resolve() != source_path.resolve())
        if same_folder and not already_copied:
            QMessageBox.warning(
                self,
                "미디어 저장",
                "미디어 저장 위치가 원본 파일이 있는 폴더와 같습니다. "
                "같은 폴더로 복사하면 원본과 뒤섞이니, 원본과 다른 폴더를 저장 위치로 선택해 주세요.",
            )
            self.set_save_status("미디어 저장 실패: 저장 위치가 원본 폴더와 동일합니다.")
            return None

        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            t0 = time.perf_counter()
            target_path = copy_database_media(source_path, target_dir, expected_stem)
            logger.info(
                "save_media_file_to_storage: copy %.3fs (%d bytes) -> %s",
                time.perf_counter() - t0,
                source_path.stat().st_size,
                target_path,
            )
            # ponytail: 동영상 캡처(ffmpeg 프로세스 실행)는 행 하나 저장할 때마다 부르면
            # 무거워서 UI가 멈춤 -> 여기서는 안 하고, "동영상 미리보기 생성" 버튼으로
            # 한번에 일괄 처리(generate_video_thumbnails)하도록 분리.
            self.set_save_status(f"미디어 저장 완료: {target_path}")
            return target_path
        except Exception as exc:
            QMessageBox.critical(self, "미디어 저장 실패", str(exc))
            self.set_save_status(f"미디어 저장 실패: {exc}")
            return None

    @staticmethod
    def capture_video_thumbnail(video_path: Path, output_path: Path, at_seconds: float = 1.0) -> bool:
        ffmpeg = find_media_tool("ffmpeg")
        if not ffmpeg:
            return False
        output_path.parent.mkdir(parents=True, exist_ok=True)
        for position in dict.fromkeys((at_seconds, 0.0)):
            try:
                result = subprocess.run(
                    [ffmpeg, "-y", "-ss", str(position), "-i", str(video_path), "-frames:v", "1", str(output_path)],
                    capture_output=True, timeout=20, check=False,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                if result.returncode == 0 and output_path.is_file() and output_path.stat().st_size > 0:
                    return True
            except (OSError, subprocess.TimeoutExpired):
                continue
        return False

    @staticmethod
    def video_thumbnail_path(video_path: Path) -> Path:
        # Saved videos and captures share the species-ID directory and database stem.
        return video_path.with_suffix(".jpg")

    @staticmethod
    def video_preview_thumbnail_path(video_path: Path) -> Path:
        return video_path.parent / "썸네일" / f"{video_path.stem}.jpg"

    def generate_video_thumbnails(self):
        try:
            deferred = getattr(self, "_defer_media_copy", False)
            self._defer_media_copy = True
            try:
                self.sync_current_form_to_table()
            finally:
                self._defer_media_copy = deferred
            success, skipped, failed, total = self.capture_pending_video_thumbnails(preview=True)
        except Exception as exc:
            QMessageBox.critical(self, "동영상 미리보기 생성 실패", str(exc))
            return
        if total == 0:
            QMessageBox.information(self, "동영상 미리보기 생성", "미리보기를 생성할 동영상 행이 없습니다.")
            return

        QMessageBox.information(
            self, "동영상 미리보기 생성", f"완료: {success}건 생성, {skipped}건 이미 있음, {failed}건 실패(ffmpeg 필요)."
        )
        self.set_save_status(f"동영상 미리보기 생성 완료: {success}건 생성 / {failed}건 실패")
        if hasattr(self, "media_thumbnail_label"):
            self.update_media_thumbnail()

    def capture_pending_video_thumbnails(self, preview=False) -> tuple[int, int, int, int]:
        # "작성 내용 저장"/"엑셀로 저장"에서 동영상 행을 저장할 때 캡처본도 같이 남기기
        # 위해 호출됨(silent). "동영상 미리보기 생성" 버튼도 이걸 재사용.
        media_column = ROW_COLUMNS.index("파일경로")
        type_column = ROW_COLUMNS.index("미디어타입")

        video_paths = []
        invalid = 0
        for row in range(self.data_table.rowCount()):
            type_item = self.data_table.item(row, type_column)
            media_item = self.data_table.item(row, media_column)
            if not type_item or type_item.text() != "동영상":
                continue
            if not media_item or not media_item.text():
                continue
            video_path = Path(media_item.text())
            if preview:
                if video_path.is_file():
                    video_paths.append(video_path)
                else:
                    invalid += 1
                continue
            values = {field: self.data_table.item(row, col).text() if self.data_table.item(row, col) else ""
                      for col, field in enumerate(ROW_COLUMNS)}
            try:
                expected_dir = media_destination(self.media_storage_dir, values)
            except ValueError:
                invalid += 1
                continue
            if (self.media_storage_dir and video_path.is_file()
                    and video_path.parent.resolve() == expected_dir.resolve()
                    and video_path.stem == normalize_text(values.get("데이터베이스번호"))):
                video_paths.append(video_path)
            else:
                invalid += 1

        success, skipped, failed = 0, 0, 0
        for video_path in video_paths:
            thumbnail_path = self.video_preview_thumbnail_path(video_path) if preview else self.video_thumbnail_path(video_path)
            if thumbnail_path.is_file() and thumbnail_path.stat().st_size > 0:
                skipped += 1
                continue
            thumbnail_path.parent.mkdir(parents=True, exist_ok=True)
            if self.capture_video_thumbnail(video_path, thumbnail_path):
                success += 1
            else:
                failed += 1
        return success, skipped, failed + invalid, len(video_paths) + invalid

    @staticmethod
    def safe_filename(value: str) -> str:
        text = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
        return text or "media"

    @staticmethod
    def unique_media_path(target_dir: Path, stem: str, suffix: str) -> Path:
        suffix = suffix or ".bin"
        target_path = target_dir / f"{stem}{suffix}"
        counter = 2
        while target_path.exists():
            target_path = target_dir / f"{stem}_{counter}{suffix}"
            counter += 1
        return target_path

    @staticmethod
    def media_file_filter(media_type: str) -> str:
        if media_type == "이미지":
            return "Image Files (*.jpg *.jpeg *.tif *.tiff *.heic *.png);;All Files (*)"
        if media_type == "동영상":
            return "Video Files (*.mp4 *.mov *.m4v *.avi *.mts *.3gp);;All Files (*)"
        if media_type == "사운드":
            return "Audio Files (*.wav *.mp3 *.m4a *.aac *.flac);;All Files (*)"
        return "Media Files (*.jpg *.jpeg *.tif *.tiff *.heic *.png *.mp4 *.mov *.m4v *.wav *.mp3 *.m4a);;All Files (*)"

    def extract_location_from_media(
        self, file_path: Path, media_type: str, ffprobe_format: dict | None = None
    ) -> tuple[float | None, float | None, str]:
        if media_type == "이미지":
            latitude, longitude = self.extract_image_gps(file_path)
            if latitude is not None and longitude is not None:
                return latitude, longitude, "이미지 EXIF 위치정보를 적용했습니다."

        latitude, longitude = self.extract_location_with_exiftool(file_path)
        if latitude is not None and longitude is not None:
            return latitude, longitude, "미디어 메타데이터 위치정보를 적용했습니다."

        latitude, longitude = self.extract_location_with_ffprobe(file_path, ffprobe_format)
        if latitude is not None and longitude is not None:
            return latitude, longitude, "동영상 메타데이터 위치정보를 적용했습니다."

        if media_type in {"동영상", "사운드"} and not shutil.which("exiftool") and not find_media_tool("ffprobe"):
            return None, None, "동영상/사운드 위치정보를 읽으려면 exiftool 또는 ffprobe가 필요합니다."
        return None, None, "선택한 파일에 읽을 수 있는 위도/경도 메타데이터가 없습니다."

    @staticmethod
    def extract_image_gps(file_path: Path) -> tuple[float | None, float | None]:
        if Image is None or ExifTags is None:
            return None, None
        try:
            with Image.open(file_path) as image:
                exif = image.getexif()
                gps_ifd = exif.get_ifd(ExifTags.IFD.GPSInfo) if hasattr(ExifTags, "IFD") else {}
        except Exception:
            return None, None

        if not gps_ifd:
            return None, None

        latitude = UploadTab.gps_to_decimal(gps_ifd.get(2), gps_ifd.get(1))
        longitude = UploadTab.gps_to_decimal(gps_ifd.get(4), gps_ifd.get(3))
        return latitude, longitude

    @staticmethod
    def gps_to_decimal(value: Any, reference: Any) -> float | None:
        if not value or len(value) < 3:
            return None
        try:
            degrees = float(value[0])
            minutes = float(value[1])
            seconds = float(value[2])
            decimal = degrees + minutes / 60 + seconds / 3600
            if str(reference).upper() in {"S", "W"}:
                decimal *= -1
            return decimal
        except Exception:
            return None

    @staticmethod
    def extract_location_with_exiftool(file_path: Path) -> tuple[float | None, float | None]:
        exiftool = shutil.which("exiftool")
        if not exiftool:
            return None, None
        data = UploadTab.run_json_command([exiftool, "-j", "-n", str(file_path)], timeout=10, empty_json="[]")

        if not data:
            return None, None
        item = data[0]
        return UploadTab.to_float(item.get("GPSLatitude")), UploadTab.to_float(item.get("GPSLongitude"))

    @staticmethod
    def extract_location_with_ffprobe(
        file_path: Path, ffprobe_format: dict | None = None
    ) -> tuple[float | None, float | None]:
        media_format = ffprobe_format if ffprobe_format is not None else UploadTab.probe_ffprobe_format(file_path)
        tags = media_format.get("tags", {})
        location = tags.get("location") or tags.get("com.apple.quicktime.location.ISO6709")
        if not location:
            return None, None
        match = re.match(r"([+-]\d+(?:\.\d+)?)([+-]\d+(?:\.\d+)?)", str(location))
        if not match:
            return None, None
        return UploadTab.to_float(match.group(1)), UploadTab.to_float(match.group(2))

    @staticmethod
    def to_float(value: Any) -> float | None:
        try:
            if value in ("", None):
                return None
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def validate_latlon(latitude_text: str, longitude_text: str) -> str:
        # 빈 문자열이면 "" 반환(정상), 그 외엔 문제를 설명하는 한글 메시지를 반환.
        # 셀 배경색 표시와 저장 전 팝업 둘 다 이 메시지를 그대로 재사용함.
        lat_text = normalize_text(latitude_text)
        lon_text = normalize_text(longitude_text)
        if not lat_text and not lon_text:
            return ""
        if bool(lat_text) != bool(lon_text):
            return "위도/경도 중 한쪽만 입력됨"

        latitude = UploadTab.to_float(lat_text)
        longitude = UploadTab.to_float(lon_text)
        if latitude is None or longitude is None:
            return "위도/경도가 숫자가 아님"
        lat_min, lat_max = KOREA_LAT_RANGE
        lon_min, lon_max = KOREA_LON_RANGE
        in_korea = lat_min <= latitude <= lat_max and lon_min <= longitude <= lon_max
        looks_swapped = lat_min <= longitude <= lat_max and lon_min <= latitude <= lon_max
        if not in_korea and looks_swapped:
            return "위도/경도가 뒤바뀐 것 같습니다"
        if not (-90 <= latitude <= 90) or not (-180 <= longitude <= 180):
            return "위도/경도 범위 오류 (위도 -90~90, 경도 -180~180)"
        if latitude == 0 or longitude == 0:
            return "0 좌표: 미입력 대체값 또는 실제 위치인지 확인 필요"
        if not in_korea:
            return "국내 범위를 벗어남 (해외 건이면 무시해도 됩니다)"
        return ""

    def create_validation_button(self, table: QTableWidget, row_columns: list[str]) -> QPushButton:
        button = QPushButton("데이터 검사")
        button.setToolTip("현재 탭의 전체 데이터에서 검사할 항목을 선택합니다.")
        menu = QMenu(button)
        number_action = menu.addAction("DB 번호 오류 검사 및 수정")
        number_action.triggered.connect(self.repair_invalid_database_numbers)
        location_action = menu.addAction("좌표 검사")
        location_action.triggered.connect(lambda: self.show_location_validation(table, row_columns))
        button.setMenu(menu)
        return button

    def show_location_validation(self, table: QTableWidget, row_columns: list[str]):
        self.refresh_location_validation(table, row_columns)
        issues = self.location_validation_issues(table, row_columns)
        if not issues:
            QMessageBox.information(self, "좌표 검사", "입력된 좌표에서 형식·범위 이상을 찾지 못했습니다.\n좌표가 모두 비어 있는 행은 검사에서 제외합니다.\n실제 관찰 위치와 일치하는지는 지도에서 확인해 주세요.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("좌표 검사 결과")
        dialog.resize(900, 480)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f"확인 필요 {len(issues)}건 · 결과를 더블클릭하면 해당 행으로 이동합니다.\n해외 좌표는 국내 범위 경고를 무시할 수 있습니다. 좌표는 자동 수정하지 않습니다."))
        results = QTableWidget(len(issues), 5, dialog)
        results.setHorizontalHeaderLabels(["행", "NO.", "위도", "경도", "확인 사유"])
        results.setEditTriggers(QAbstractItemView.NoEditTriggers)
        results.setSelectionBehavior(QAbstractItemView.SelectRows)
        entries = []
        for index, (row, label, message) in enumerate(issues):
            coordinates = []
            for field in ("위도", "경도"):
                item = table.item(row, row_columns.index(field))
                coordinates.append(item.text() if item else "")
            entry = (str(row + 1), label, *coordinates, message)
            entries.append(entry)
            for column, value in enumerate(entry):
                results.setItem(index, column, QTableWidgetItem(value))
        results.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        results.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        layout.addWidget(results)

        def jump_to_row(index, _column):
            row = issues[index][0]
            table.setCurrentCell(row, row_columns.index("위도"))
            table.selectRow(row)
            table.scrollToItem(table.item(row, row_columns.index("위도")))
            dialog.accept()
            table.setFocus()

        results.cellDoubleClicked.connect(jump_to_row)
        buttons = QHBoxLayout()
        copy_button = QPushButton("결과 복사")
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText(
            "행\tNO.\t위도\t경도\t확인 사유\n" + "\n".join("\t".join(entry) for entry in entries)
        ))
        close_button = QPushButton("닫기")
        close_button.clicked.connect(dialog.accept)
        buttons.addStretch()
        buttons.addWidget(copy_button)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)
        dialog.exec()

    def location_validation_issues(self, table: QTableWidget, row_columns: list[str]) -> list[tuple[int, str, str]]:
        lat_col = row_columns.index("위도")
        lon_col = row_columns.index("경도")
        no_col = row_columns.index("NO.")
        issues = []
        for row in range(table.rowCount()):
            lat_item = table.item(row, lat_col)
            lon_item = table.item(row, lon_col)
            message = self.validate_latlon(lat_item.text() if lat_item else "", lon_item.text() if lon_item else "")
            if not message:
                continue
            no_item = table.item(row, no_col)
            label = f"NO. {no_item.text()}" if no_item and no_item.text() else f"{row + 1}행"
            issues.append((row, label, message))
        return issues

    def revalidate_row_location(self, table: QTableWidget, row_columns: list[str], row: int):
        lat_col = row_columns.index("위도")
        lon_col = row_columns.index("경도")
        lat_item = table.item(row, lat_col)
        lon_item = table.item(row, lon_col)
        message = self.validate_latlon(lat_item.text() if lat_item else "", lon_item.text() if lon_item else "")
        if message:
            if lat_item is None:
                lat_item = QTableWidgetItem("")
                table.setItem(row, lat_col, lat_item)
            if lon_item is None:
                lon_item = QTableWidgetItem("")
                table.setItem(row, lon_col, lon_item)
            lat_item.setBackground(LOCATION_ERROR_COLOR)
            lon_item.setBackground(LOCATION_ERROR_COLOR)
            lat_item.setToolTip(message)
            lon_item.setToolTip(message)
        else:
            if lat_item is not None:
                lat_item.setBackground(QBrush())
                lat_item.setToolTip("")
            if lon_item is not None:
                lon_item.setBackground(QBrush())
                lon_item.setToolTip("")

    def refresh_location_validation(self, table: QTableWidget, row_columns: list[str]):
        was_blocked = table.blockSignals(True)
        try:
            for row in range(table.rowCount()):
                self.revalidate_row_location(table, row_columns, row)
        finally:
            table.blockSignals(was_blocked)

    def confirm_location_issues_before_save(self, table: QTableWidget, row_columns: list[str]) -> bool:
        issues = self.location_validation_issues(table, row_columns)
        if not issues:
            return True
        lines = "\n".join(f"- {label}: {message}" for _, label, message in issues[:20])
        more_note = f"\n... 외 {len(issues) - 20}건" if len(issues) > 20 else ""
        reply = QMessageBox.warning(
            self,
            "위치 확인 필요",
            f"위도/경도에 문제가 있어 보이는 행이 {len(issues)}건 있습니다:\n{lines}{more_note}\n\n그래도 저장하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
        )
        return reply == QMessageBox.Yes

    def add_data_row(self, values: dict[str, Any]):
        row = self.data_table.rowCount()
        self.data_table.insertRow(row)
        self.write_data_row(row, values)

    def write_data_row(self, row: int, values: dict[str, Any]):
        old_sequence = self.data_table.item(row, ROW_COLUMNS.index("사업별 순번"))
        sequence_changed = old_sequence is not None and old_sequence.text().strip() != str(values.get("사업별 순번", "")).strip()
        values = self.preserve_row_database_number(self.data_table, ROW_COLUMNS, row, values)
        was_blocked = self.data_table.blockSignals(True)
        try:
            for col, column_name in enumerate(ROW_COLUMNS):
                value = values.get(column_name, "")
                item = QTableWidgetItem("" if value is None else str(value))
                if column_name == "데이터베이스번호":
                    item.setData(DB_NUMBER_MANUAL_ROLE, bool(values.get("_database_number_manual")))
                self.data_table.setItem(row, col, item)
        finally:
            self.data_table.blockSignals(was_blocked)

        if not self.autosave_suspended:
            self.recalculate_project_sequences(self.data_table, ROW_COLUMNS)

        if sequence_changed:
            self.renumber_project_sequences_below(self.data_table, ROW_COLUMNS, row, False)

    def current_data_row(self) -> int | None:
        indexes = self.data_table.selectedIndexes()
        if not indexes:
            return None
        # selectedIndexes()는 정렬 보장이 없어서(드래그 방향에 따라 순서가 달라짐)
        # indexes[0]을 쓰면 셀 범위를 드래그로 선택했을 때 엉뚱한 행이 잡힘 -> 항상 맨 위 행으로
        return min(index.row() for index in indexes)

    def preserve_row_database_number(self, table, columns, row, values):
        sequence_item = table.item(row, columns.index("사업별 순번"))
        sequence_changed = sequence_item is not None and sequence_item.text().strip() != str(values.get("사업별 순번", "")).strip()
        if self.preserve_database_numbers.isChecked() and not sequence_changed:
            item = table.item(row, columns.index("데이터베이스번호"))
            if item and (item.text().strip() or item.data(DB_NUMBER_MANUAL_ROLE)):
                values = dict(values)
                values["데이터베이스번호"] = item.text()
                values["_database_number_manual"] = bool(item.data(DB_NUMBER_MANUAL_ROLE))
        return values

    def project_sequence_for_form(self, specimen, new=False):
        widgets = self.specimen_form_widgets if specimen else self.form_widgets
        if not all(field in widgets for field in ("사업년도", "순번")):
            return 1
        name = "specimen_data_table" if specimen else "data_table"
        if not hasattr(self, name):
            return 1
        table = getattr(self, name)
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        group = tuple(str(self.get_form_value(field, widgets)) for field in ("사업년도", "순번"))
        selected = None if new else (self.current_specimen_data_row() if specimen else self.current_data_row())
        if selected is not None:
            item = table.item(selected, columns.index("사업별 순번"))
            if item and item.text().isdigit():
                return int(item.text())
        maximum = 0
        for row in range(table.rowCount()):
            key = tuple(table.item(row, columns.index(field)).text().strip() if table.item(row, columns.index(field)) else ""
                        for field in ("사업년도", "순번"))
            item = table.item(row, columns.index("사업별 순번"))
            if key == group and item and item.text().isdigit():
                maximum = max(maximum, int(item.text()))
        return maximum + 1

    def commit_project_sequence(self, specimen):
        suspended = self.specimen_autosave_suspended if specimen else self.autosave_suspended
        if suspended:
            return
        row = self._specimen_last_selected_row if specimen else self._last_selected_row
        if row is None:
            return
        table = self.specimen_data_table if specimen else self.data_table
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        widgets = self.specimen_form_widgets if specimen else self.form_widgets
        value = str(self.get_form_value("사업별 순번", widgets)).strip()
        old = table.item(row, columns.index("사업별 순번"))
        if not value.isdigit() or not 1 <= int(value) <= 99999 or (old and old.text().strip() == value):
            return
        values = (self.collect_specimen_form_values if specimen else self.collect_form_values)()
        (self.write_specimen_data_row if specimen else self.write_data_row)(row, values)
        (self.schedule_specimen_autosave if specimen else self.schedule_autosave)()

    def renumber_project_sequences_below(self, table, columns, start_row, specimen):
        def cell(row, name):
            item = table.item(row, columns.index(name))
            return item.text().strip() if item else ""

        value = cell(start_row, "사업별 순번")
        if not value.isdigit() or not 1 <= int(value) <= 99999:
            return
        group = tuple(cell(start_row, field) for field in ("사업년도", "순번"))
        if not all(group):
            return
        rows = [row for row in range(start_row, table.rowCount())
                if tuple(cell(row, field) for field in ("사업년도", "순번")) == group]
        if int(value) + len(rows) - 1 > 99999:
            QMessageBox.warning(self, "사업별 순번", "아래 행까지 번호를 매기면 99999를 초과합니다. 시작 번호를 낮춰 주세요.")
            return
        blocked = table.blockSignals(True)
        try:
            for sequence, row in enumerate(rows, int(value)):
                table.setItem(row, columns.index("사업별 순번"), QTableWidgetItem(str(sequence)))
                expected, _, error = self._database_number_rule(table, columns, row, specimen)
                table.setItem(row, columns.index("데이터베이스번호"), QTableWidgetItem("" if error else expected))
        finally:
            table.blockSignals(blocked)

    def recalculate_project_sequences(self, table, columns):
        if "사업별 순번" not in columns:
            return
        counts = {}
        blocked = table.blockSignals(True)
        try:
            for row in range(table.rowCount()):
                group = tuple(table.item(row, columns.index(field)).text().strip() if table.item(row, columns.index(field)) else ""
                              for field in ("사업년도", "순번"))
                index = columns.index("사업별 순번")
                item = table.item(row, index)
                text = item.text().strip() if item else ""
                if text:
                    if text.isdigit():
                        counts[group] = max(counts.get(group, 0), int(text))
                    continue
                counts[group] = counts.get(group, 0) + 1
                if all(group):
                    table.setItem(row, index, QTableWidgetItem(str(counts[group])))
        finally:
            table.blockSignals(blocked)

    def prepare_new_project_number(self, values, specimen):
        values = dict(values)
        raw_sequence = str(values.get("사업별 순번", ""))
        sequence = int(raw_sequence) if raw_sequence.isdigit() and int(raw_sequence) > 0 else self.project_sequence_for_form(specimen, new=True)
        values["사업별 순번"] = str(sequence)
        media = "SP" if specimen else MEDIA_TYPE_CODES.get(values.get("미디어타입"), "")
        fields = [values.get(field) for field in ("기관코드", "분류코드", "사업년도", "순번")]
        if all(fields) and media:
            org, taxon, year, order = fields
            values["데이터베이스번호"] = f"{org}-{media}-{taxon}-{int(year) % 100:02d}{int(order):02d}{sequence:05d}"
        return values

    def restore_existing_database_number(self, specimen):
        if not self.preserve_database_numbers.isChecked():
            return False
        table_name = "specimen_data_table" if specimen else "data_table"
        if not hasattr(self, table_name):
            return False
        table = getattr(self, table_name)
        row = self._specimen_last_selected_row if specimen else self._last_selected_row
        if row is None:
            return False
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        item = table.item(row, columns.index("데이터베이스번호"))
        if not item or not (item.text().strip() or item.data(DB_NUMBER_MANUAL_ROLE)):
            return False
        widgets = self.specimen_form_widgets if specimen else self.form_widgets
        sequence_item = table.item(row, columns.index("사업별 순번"))
        if sequence_item and sequence_item.text().strip() != str(self.get_form_value("사업별 순번", widgets)).strip():
            return False
        self.set_form_value("데이터베이스번호", item.text(), widgets)
        return True

    @staticmethod
    def mark_database_number_manual(table, item):
        blocked = table.blockSignals(True)
        item.setData(DB_NUMBER_MANUAL_ROLE, True)
        table.blockSignals(blocked)

    def edit_database_number(self, specimen):
        table = self.specimen_data_table if specimen else self.data_table
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        widgets = self.specimen_form_widgets if specimen else self.form_widgets
        row = self.current_specimen_data_row() if specimen else self.current_data_row()
        if row is None:
            return
        value = self.get_form_value("데이터베이스번호", widgets)
        item = QTableWidgetItem(value)
        item.setData(DB_NUMBER_MANUAL_ROLE, True)
        blocked = table.blockSignals(True)
        table.setItem(row, columns.index("데이터베이스번호"), item)
        table.blockSignals(blocked)
        (self.autosave_specimen_draft if specimen else self.autosave_draft)()

    def _database_number_rule(self, table, columns, row, specimen):
        def cell(name):
            item = table.item(row, columns.index(name))
            return item.text().strip() if item else ""

        org, taxon, year, order, sequence = [
            cell(field) for field in ("기관코드", "분류코드", "사업년도", "순번", "사업별 순번")
        ]
        media = "SP" if specimen else MEDIA_TYPE_CODES.get(cell("미디어타입"), cell("미디어코드"))
        missing = [
            label
            for label, value in (
                ("기관코드", org),
                ("분류코드", taxon),
                ("미디어타입", media),
                ("사업년도", year),
                ("순번", order),
                ("사업별 순번", sequence),
            )
            if not value
        ]
        if missing:
            return None, None, f"필수값 누락: {', '.join(missing)}"
        try:
            year_code = int(year) % 100
            order_code = int(order)
            sequence_code = int(sequence)
            if order_code < 0 or sequence_code < 1:
                raise ValueError
        except (TypeError, ValueError):
            return None, None, "사업년도·순번·사업별 순번이 올바른 숫자가 아님"

        prefix = f"{org}-{media}-{taxon}-{year_code:02d}{order_code:02d}"
        expected = f"{prefix}{sequence_code:05d}"
        pattern = re.compile(rf"{re.escape(prefix)}\d{{5}}")
        return expected, pattern, ""

    def repair_invalid_database_numbers(self):
        specimen = self.tabs.currentIndex() == 1
        if self.tabs.currentIndex() not in (0, 1):
            QMessageBox.information(self, "DB 번호 검사", "관찰 정보작성 또는 표본정보 작성 탭에서 실행해 주세요.")
            return
        table = self.specimen_data_table if specimen else self.data_table
        columns = SPECIMEN_ROW_COLUMNS if specimen else ROW_COLUMNS
        if table.rowCount() == 0:
            QMessageBox.information(self, "DB 번호 검사", "검사할 데이터가 없습니다.")
            return

        selected_row = self.current_specimen_data_row() if specimen else self.current_data_row()
        if selected_row is not None:
            (self.sync_current_specimen_form_to_table if specimen else self.sync_current_form_to_table)()
        self.recalculate_project_sequences(table, columns)

        db_column = columns.index("데이터베이스번호")
        numbers = []
        for row in range(table.rowCount()):
            item = table.item(row, db_column)
            numbers.append(item.text().strip() if item else "")

        changes = []
        failures = []
        valid_count = 0
        seen_numbers = set()
        blocked = table.blockSignals(True)
        try:
            for row in range(table.rowCount()):
                old_number = numbers[row]
                expected, pattern, error = self._database_number_rule(table, columns, row, specimen)
                duplicate = bool(old_number and old_number in seen_numbers)
                if old_number:
                    seen_numbers.add(old_number)

                if pattern is not None and pattern.fullmatch(old_number) and not duplicate:
                    valid_count += 1
                    continue

                reason = "번호 없음" if not old_number else ("중복 번호" if duplicate else "번호 형식 또는 행 정보와 불일치")
                if error:
                    failures.append((row + 1, old_number, f"{reason}; {error}"))
                    continue
                if expected in seen_numbers:
                    failures.append((row + 1, old_number, f"{reason}; 생성할 번호 {expected}가 이미 존재함"))
                    continue

                new_item = QTableWidgetItem(expected)
                new_item.setData(DB_NUMBER_MANUAL_ROLE, False)
                table.setItem(row, db_column, new_item)
                seen_numbers.add(expected)
                changes.append((row + 1, old_number, expected, reason))
        finally:
            table.blockSignals(blocked)

        if selected_row is not None:
            (self.load_selected_specimen_row_to_form if specimen else self.load_selected_row_to_form)()
        if changes:
            (self.autosave_specimen_draft if specimen else self.autosave_draft)()
        self.show_database_number_repair_log(
            "표본" if specimen else "관찰", valid_count, changes, failures
        )

    def show_database_number_repair_log(self, data_type, valid_count, changes, failures):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{data_type} DB 번호 검사 결과")
        dialog.resize(980, 520)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            f"정상 {valid_count}건 · 수정 {len(changes)}건 · 수정 불가 {len(failures)}건"
        ))

        result_table = QTableWidget(len(changes) + len(failures), 5, dialog)
        result_table.setHorizontalHeaderLabels(["결과", "행", "변경 전", "변경 후", "사유"])
        result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        result_table.verticalHeader().setVisible(False)
        entries = [
            ("수정", row, old or "(빈 값)", new, reason)
            for row, old, new, reason in changes
        ] + [
            ("수정 불가", row, old or "(빈 값)", "-", reason)
            for row, old, reason in failures
        ]
        for row_index, entry in enumerate(entries):
            for column_index, value in enumerate(entry):
                result_table.setItem(row_index, column_index, QTableWidgetItem(str(value)))
        header = result_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        layout.addWidget(result_table)

        button_layout = QHBoxLayout()
        copy_button = QPushButton("결과 복사")
        close_button = QPushButton("닫기")
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText("\n".join(
            "\t".join(map(str, entry)) for entry in entries
        )))
        close_button.clicked.connect(dialog.accept)
        button_layout.addStretch()
        button_layout.addWidget(copy_button)
        button_layout.addWidget(close_button)
        layout.addLayout(button_layout)
        dialog.exec()

    def regenerate_selected_database_number(self):
        """이전 호출 경로와의 호환을 위해 전체 검사 기능으로 연결한다."""
        self.repair_invalid_database_numbers()
        (self.autosave_specimen_draft if specimen else self.autosave_draft)()

    def paste_clipboard_into_table(self):
        text = QApplication.clipboard().text()
        if not text:
            return

        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if lines and lines[-1] == "":
            lines.pop()
        if not lines:
            return
        rows = [line.split("\t") for line in lines]

        # currentColumn()은 드래그로 범위를 선택했을 때 "마지막으로 지나간 셀"의 열이라
        # 목적지 왼쪽 위가 아닐 수 있음 -> 선택 범위 중 가장 왼쪽 위 셀로 고정.
        # 선택된 셀이 하나도 없으면 (0,0)으로 잘못 덮어쓰지 않도록 그냥 붙여넣기를 취소함
        indexes = self.data_table.selectedIndexes()
        if not indexes:
            self.set_save_status("붙여넣기 취소: 붙여넣을 셀을 먼저 선택해 주세요.")
            return
        start_row = min(index.row() for index in indexes)
        start_col = min(index.column() for index in indexes)

        was_blocked = self.data_table.blockSignals(True)
        try:
            for row_offset, row_values in enumerate(rows):
                target_row = start_row + row_offset
                while target_row >= self.data_table.rowCount():
                    self.data_table.insertRow(self.data_table.rowCount())
                for col_offset, value in enumerate(row_values):
                    target_col = start_col + col_offset
                    if target_col >= len(ROW_COLUMNS):
                        break
                    self.data_table.setItem(target_row, target_col, QTableWidgetItem(value))
        finally:
            self.data_table.blockSignals(was_blocked)

        # 붙여넣기 중엔 신호를 막아뒀으니, 현재 폼에 로드된 행이 이번에 붙여넣은 범위에
        # 포함됐다면 국명->학명 자동조회/날짜 형식변환을 셀 직접 수정 때와 동일하게 적용
        current_row = self.current_data_row()
        if current_row is not None and start_row <= current_row < start_row + len(rows):
            row_values = rows[current_row - start_row]
            touched_columns = {
                ROW_COLUMNS[start_col + offset]
                for offset in range(len(row_values))
                if start_col + offset < len(ROW_COLUMNS)
            }
            self.load_selected_row_to_form()
            if "국명" in touched_columns:
                self.fill_species_from_korean_name()
            for date_field in DATE_FIELDS & touched_columns:
                self.normalize_date_field(date_field)
            self.write_data_row(current_row, self.collect_form_values())

        self.refresh_location_validation(self.data_table, ROW_COLUMNS)
        self.autosave_draft()
        self.set_save_status(f"붙여넣기 완료: {len(rows)}행")

    def go_to_next_row(self):
        total = self.data_table.rowCount()
        if total == 0:
            self.set_save_status("작성된 행이 없습니다.")
            return

        row = self.current_data_row()
        if row is None and self._last_selected_row is None:
            values = self.collect_form_values()
            if self.has_form_content(values):
                reply = QMessageBox.question(
                    self,
                    "다음 행",
                    "'입력'으로 저장하지 않은 작성 중인 내용이 있습니다. "
                    "행을 선택하면 지금 작성 중인 내용은 사라집니다. 계속할까요?",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if reply != QMessageBox.Yes:
                    return

        next_row = 0 if row is None else row + 1
        if next_row >= total:
            self.set_save_status("마지막 행입니다.")
            return

        self.data_table.selectRow(next_row)

    def delete_selected_data_rows(self):
        rows = sorted({index.row() for index in self.data_table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.warning(self, "선택 행 삭제", "삭제할 행을 먼저 선택해 주세요.")
            return
        self.delete_selected_rows(self.data_table)
        self.autosave_draft()
        self.set_save_status(f"{len(rows)}개 행을 삭제했습니다.")

    def clear_data_rows(self):
        if self.data_table.rowCount() == 0:
            return
        reply = QMessageBox.question(self, "작성 내용 비우기", "작성 중인 행을 모두 비울까요?")
        if reply == QMessageBox.Yes:
            self.data_table.setRowCount(0)
            self.autosave_draft()

    def add_project(self):
        name = self.project_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "사업 추가", "사업명을 입력해 주세요.")
            return
        row = {
            "사업년도": self.project_year_input.value(),
            "순번": self.project_order_input.value(),
            "사업명": name,
        }
        self.add_reference_row(self.project_table, row, PROJECT_COLUMNS)
        self.project_name_input.clear()
        self.save_projects_from_table(show_message=False)

    def delete_selected_projects(self):
        self.delete_selected_rows(self.project_table)
        self.save_projects_from_table(show_message=False)

    def save_projects_from_table(self, show_message: bool = True):
        self.projects = self.table_to_records(self.project_table, PROJECT_COLUMNS)
        self.projects.sort(
            key=lambda row: (
                self.safe_sort_int(row.get("사업년도")),
                self.safe_sort_int(row.get("순번")),
                str(row.get("사업명", "")),
            )
        )
        self.populate_reference_table(self.project_table, self.projects, PROJECT_COLUMNS)
        self.write_json(PROJECTS_FILE, self.projects)
        self.refresh_combos()
        if show_message:
            QMessageBox.information(self, "사업 저장", "사업 기준 데이터를 저장했습니다.")

    def add_organization(self):
        name = self.organization_name_input.text().strip()
        code = self.organization_code_input.text().strip().upper()
        if not name or not code:
            QMessageBox.warning(self, "기관 추가", "기관명과 기관코드를 모두 입력해 주세요.")
            return
        self.add_reference_row(self.organization_table, {"기관명": name, "기관코드": code}, ORGANIZATION_COLUMNS)
        self.organization_name_input.clear()
        self.organization_code_input.clear()
        self.save_organizations_from_table(show_message=False)

    def delete_selected_organizations(self):
        self.delete_selected_rows(self.organization_table)
        self.save_organizations_from_table(show_message=False)

    def save_organizations_from_table(self, show_message: bool = True):
        self.organizations = self.table_to_records(self.organization_table, ORGANIZATION_COLUMNS)
        self.organizations.sort(key=lambda row: str(row.get("기관명", "")))
        self.populate_reference_table(self.organization_table, self.organizations, ORGANIZATION_COLUMNS)
        self.write_json(ORGANIZATIONS_FILE, self.organizations)
        self.refresh_combos()
        if show_message:
            QMessageBox.information(self, "기관 저장", "기관 기준 데이터를 저장했습니다.")

    def save_draft(self):
        self.sync_current_form_to_table()
        self.copy_pending_media_for_all_rows()
        self.capture_pending_video_thumbnails()
        records = self.table_to_records(self.data_table, ROW_COLUMNS)
        self.write_draft(records, self.collect_form_values())
        self.set_save_status(f"임시저장 완료: {len(records)}건")
        QMessageBox.information(self, "작성 내용 저장", f"작성 중인 신규 데이터 {len(records)}건을 저장했습니다.\n{DRAFT_FILE}")

    def load_draft(self):
        if not DRAFT_FILE.exists():
            return
        draft = self.read_draft()
        self.autosave_suspended = True
        try:
            for record in draft["rows"]:
                self.add_data_row(record)
            current_form = draft.get("current_form", {})
            if current_form:
                for column_name in ROW_COLUMNS:
                    self.set_form_value(column_name, current_form.get(column_name, ""))
            self.media_storage_dir = draft.get("media_storage_dir", "")
            if self.media_storage_dir and hasattr(self, "media_storage_path_input"):
                self.media_storage_path_input.setText(self.media_storage_dir)
        finally:
            self.autosave_suspended = False
        self.refresh_location_validation(self.data_table, ROW_COLUMNS)

    def reload_draft(self):
        self.autosave_suspended = True
        try:
            self.data_table.setRowCount(0)
        finally:
            self.autosave_suspended = False
        self.load_draft()
        self.set_save_status("임시저장을 다시 불러왔습니다.")

    def export_rows_to_excel(self):
        deferred = getattr(self, "_defer_media_copy", False)
        self._defer_media_copy = True
        try:
            self.sync_current_form_to_table()
        finally:
            self._defer_media_copy = deferred
        records = self.table_to_records(self.data_table, ROW_COLUMNS)
        if not records:
            QMessageBox.warning(self, "엑셀 저장", "저장할 행이 없습니다.")
            return
        if not self.confirm_location_issues_before_save(self.data_table, ROW_COLUMNS):
            return

        default_name = f"new_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "신규 데이터 엑셀 저장", default_name, "Excel Files (*.xlsx)")
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            if not self.prepare_media_for_excel_export(False, choose_destination=True):
                return
            records = self.table_to_records(self.data_table, ROW_COLUMNS)
            self.write_excel_file(file_path, records)
            final_path = self.final_excel_path(file_path)
            self.write_final_excel_file(str(final_path), records)
            gbif_path = self.gbif_excel_path(file_path)
            write_gbif_template(str(gbif_path), records, "observation")
            self.reset_after_export()
            QMessageBox.information(
                self,
                "엑셀 저장",
                f"신규 데이터 엑셀 3개를 저장했습니다.\n{file_path}\n{final_path}\n{gbif_path}\n\n"
                "작성 중이던 내용은 초기화되어 NO.가 1부터 다시 시작됩니다.",
            )
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 저장 실패", str(exc))

    def reset_after_export(self):
        self.autosave_suspended = True
        try:
            self.data_table.setRowCount(0)
        finally:
            self.autosave_suspended = False
        self.clear_form()
        self.write_draft([], self.collect_form_values())
        self.set_save_status("엑셀 저장 완료. 새 작성을 시작합니다 (NO. 1부터).")

    def map_excel_columns(self, df: pd.DataFrame, columns: list[str], export_columns, kind: str) -> pd.DataFrame | None:
        """엑셀 컬럼을 프로그램 컬럼으로 매핑. 자동 매칭 후 매핑 창에서 확인. 취소 시 None.
        kind("row"/"specimen")별로 사용자가 정한 매핑을 COLUMN_MAPPINGS_FILE에 저장해 다음에 자동 적용."""
        df = df.dropna(axis=1, how="all")  # 헤더만 있고 값이 없는 열(빈 서식 열 포함)은 매핑 대상에서 제외
        df.columns = [str(name).strip() for name in df.columns]
        all_saved = self.read_json(COLUMN_MAPPINGS_FILE, {})
        saved = all_saved.get(kind, {})
        mapping = auto_map(df.columns, columns, export_columns, saved)
        dialog = ColumnMappingDialog(df, mapping, columns, self)
        if dialog.exec() != QDialog.Accepted:
            return None
        mapping = dialog.mapping()
        all_saved[kind] = {**saved, **mapping}
        self.write_json(COLUMN_MAPPINGS_FILE, all_saved)
        result = apply_mapping(df, mapping, columns)
        epoch = df.attrs.get("excel_epoch", datetime(1899, 12, 30))
        for field in DATE_FIELDS.intersection(result.columns):
            result[field] = result[field].map(lambda value: UploadTab.clean_import_date(value, epoch))
        return result

    @staticmethod
    def clean_import_date(value: Any, epoch: datetime = datetime(1899, 12, 30)) -> str:
        if pd.isna(value):
            return ""
        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")
        text = UploadTab.clean_master_value(value)
        # Six/eight digit calendar dates retain their existing interpretation.
        if text.isdigit() and len(text) in (6, 8):
            return UploadTab.normalize_date_text(text)
        if re.fullmatch(r"\d{1,5}(?:\.\d+)?", text):
            serial = float(text)
            if 1 <= serial < 100000:
                # Excel's 1900 calendar includes a fictitious leap day.
                if epoch == datetime(1899, 12, 30) and serial < 60:
                    serial += 1
                return (epoch + timedelta(days=serial)).strftime("%Y-%m-%d")
        return UploadTab.normalize_date_text(text)

    def select_excel_sheet(self, file_path):
        with open_excel_lenient(file_path) as workbook:
            sheet_name, accepted = QInputDialog.getItem(
                self, "엑셀 시트 선택", "불러올 시트를 선택하세요.", workbook.sheet_names, 0, False
            )
            if not accepted:
                return None
            return parse_excel_sheet(workbook, sheet_name)

    def load_rows_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "신규 데이터 엑셀 불러오기", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            df = self.select_excel_sheet(file_path)
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 불러오기 실패", str(exc))
            return

        if df is None:
            return

        df = self.map_excel_columns(df, ROW_COLUMNS, FINAL_EXPORT_COLUMNS, "row")
        if df is None:
            return

        if self.data_table.rowCount() > 0:
            reply = QMessageBox.question(self, "엑셀 불러오기", "현재 작성 중인 내용을 지우고 불러올까요?")
            if reply != QMessageBox.Yes:
                return

        self.autosave_suspended = True
        try:
            self.data_table.setRowCount(0)
            for _, row in df.iterrows():
                record = {column: self.clean_master_value(row[column]) for column in ROW_COLUMNS}
                self.add_data_row(record)
        finally:
            self.autosave_suspended = False

        self.refresh_location_validation(self.data_table, ROW_COLUMNS)
        self.autosave_draft()
        self.set_save_status(f"엑셀 불러오기 완료: {file_path} ({len(df)}건)")

    def sync_current_form_to_table(self):
        # 선택된 행이 있을 때만 그 행에 지금 폼 내용을 반영. 새 행 추가는 "입력" 버튼으로만
        # 하도록 함 - 예전엔 선택 없을 때 폼 내용으로 새 행을 추측해서 만들었는데, 그게
        # 저장(임시저장/엑셀로 저장) 시 의도치 않은 중복 행이 생기는 원인이었음.
        row = self.current_data_row()
        if row is None:
            return
        values = self.collect_form_values()
        if not self.has_form_content(values):
            return
        values = self.prepare_media_for_record(values)
        self.write_data_row(row, values)

    def has_form_content(self, values: dict[str, Any]) -> bool:
        # 값이 비어있지 않다고 해서 "내용 있음"은 아님 - 세션 유지값(STICKY_FIELDS)이나
        # 자동생성 값(NO./미디어코드/데이터베이스번호)일 수도 있고, 드롭다운/국가명처럼
        # 기본값이 채워진 상태(clear_form 직후)일 수도 있음. 이 둘을 놓치면 선택 없이
        # 저장할 때 빈 폼이 새 행으로 중복 추가되는 버그가 생김(실제로 두 번 발생했던 문제) ->
        # clear_form()이 실제로 채우는 기본값(default_field_value)과 직접 비교해서 판단.
        ignored_fields = STICKY_FIELDS | {"NO.", "미디어코드", "데이터베이스번호"}
        for key, value in values.items():
            if key in ignored_fields:
                continue
            normalized = value if value is not None else ""
            if str(normalized) != str(self.default_field_value(key)):
                return True
        return False

    def is_duplicate_record(self, values: dict[str, Any]) -> bool:
        current = {column: "" if values.get(column) is None else str(values.get(column, "")) for column in ROW_COLUMNS}
        for record in self.table_to_records(self.data_table, ROW_COLUMNS):
            existing = {column: "" if record.get(column) is None else str(record.get(column, "")) for column in ROW_COLUMNS}
            if existing == current:
                return True
        return False

    def connect_autosave_signal(self, widget: QLineEdit | QComboBox | QSpinBox):
        if isinstance(widget, QLineEdit):
            widget.textChanged.connect(self.schedule_autosave)
        elif isinstance(widget, QComboBox):
            widget.currentTextChanged.connect(self.schedule_autosave)
        elif isinstance(widget, QSpinBox):
            widget.valueChanged.connect(self.schedule_autosave)

    def schedule_autosave(self, *_):
        if self.autosave_suspended:
            return
        self.autosave_timer.start(700)

    def autosave_draft(self):
        if self.autosave_suspended or not hasattr(self, "data_table") or not self.form_widgets:
            return
        start = time.perf_counter()
        self.autosave_suspended = True
        try:
            # ponytail: 현재 선택 행에 폼 내용을 반영하는 건 여기서 하지 않음 -> 행 선택이
            # 바뀔 때(load_selected_row_to_form)만 반영. 타이핑 중 매번 표에 되쓰면 행이
            # 많을 때 컬럼폭 재계산 등으로 느려짐. 임시저장 JSON엔 현재 폼 값이 별도 저장되어
            # 유실 없음.
            t_row_done = time.perf_counter()

            records = self.table_to_records(self.data_table, ROW_COLUMNS)
            t_records_done = time.perf_counter()

            # ponytail: 엑셀(원본+완성본) 재작성은 여기서 하지 않음 - 입력할 때마다(700ms
            # 디바운스) 전체 표를 xlsx로 다시 쓰면 행이 많아질수록 저장 중 UI가 멈춤.
            # 임시저장 JSON만으로 복구는 충분하고, 엑셀 반영은 "작성 내용 저장"/"엑셀로 저장"
            # 같은 명시적 클릭 시에만 수행.
            self.write_draft(
                records,
                self.collect_form_values(),
            )
            t_draft_done = time.perf_counter()

            self.set_save_status("자동 임시저장 완료")
            logger.info(
                "autosave_draft: rows=%d row_write=%.3fs table_to_records=%.3fs draft_write=%.3fs total=%.3fs",
                len(records),
                t_row_done - start,
                t_records_done - t_row_done,
                t_draft_done - t_records_done,
                t_draft_done - start,
            )
        finally:
            self.autosave_suspended = False

    def write_draft(self, rows: list[dict[str, Any]], current_form: dict[str, Any]):
        payload = {
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "rows": rows,
            "current_form": current_form,
            "media_storage_dir": self.media_storage_dir,
        }
        self.write_json(DRAFT_FILE, payload)

    def read_draft(self) -> dict[str, Any]:
        data = self.read_json(DRAFT_FILE, {})
        if isinstance(data, list):
            return {"rows": data, "current_form": {}}
        if not isinstance(data, dict):
            return {"rows": [], "current_form": {}}
        rows = data.get("rows", [])
        current_form = data.get("current_form", {})
        return {
            "rows": rows if isinstance(rows, list) else [],
            "current_form": current_form if isinstance(current_form, dict) else {},
            "media_storage_dir": str(data.get("media_storage_dir", "")),
        }


    @staticmethod
    def write_excel_file(file_path: str, records: list[dict[str, Any]]):
        pd.DataFrame(records, columns=ROW_COLUMNS).to_excel(file_path, index=False)

    @staticmethod
    def final_excel_path(file_path: str) -> Path:
        path = Path(file_path)
        return path.with_name(f"{path.stem}_완성본{path.suffix}")

    @staticmethod
    def gbif_excel_path(file_path: str) -> Path:
        path = Path(file_path)
        return path.with_name(f"{path.stem}_GBIF{path.suffix}")

    @staticmethod
    def write_final_excel_file(file_path: str, records: list[dict[str, Any]]):
        write_final_template(file_path, records, FINAL_EXPORT_COLUMNS, "observation")

    def populate_reference_table(self, table: QTableWidget, records: list[dict[str, Any]], columns: list[str]):
        table.setRowCount(0)
        for record in records:
            self.add_reference_row(table, record, columns)

    def add_reference_row(self, table: QTableWidget, values: dict[str, Any], columns: list[str]):
        row = table.rowCount()
        table.insertRow(row)
        for col, column_name in enumerate(columns):
            value = values.get(column_name, "")
            table.setItem(row, col, QTableWidgetItem("" if value is None else str(value)))

    def table_to_records(self, table: QTableWidget, columns: list[str]) -> list[dict[str, Any]]:
        self.recalculate_project_sequences(table, columns)
        records = []
        for row in range(table.rowCount()):
            record = {}
            has_value = False
            for col, column_name in enumerate(columns):
                item = table.item(row, col)
                value = item.text().strip() if item else ""
                if column_name in {"NO.", "사업년도", "순번", "관찰개체수"} and value:
                    value = self.safe_int(value)
                record[column_name] = value
                if column_name == "데이터베이스번호" and item and item.data(DB_NUMBER_MANUAL_ROLE):
                    record["_database_number_manual"] = True
                has_value = has_value or bool(value)
            if has_value:
                records.append(record)
        return records

    def get_form_value(self, column_name: str, widgets: dict[str, Any] | None = None) -> Any:
        widget = (widgets if widgets is not None else self.form_widgets)[column_name]
        if isinstance(widget, QSpinBox):
            return widget.value()
        if isinstance(widget, QComboBox):
            text = widget.currentText().strip()
            if column_name == "사업명":
                return self.project_lookup.get(text, {}).get("사업명", text)
            if column_name == "기관명":
                return self.organization_lookup.get(text, {}).get("기관명", text)
            return text
        return widget.text().strip()

    def set_form_value(self, column_name: str, value: Any, widgets: dict[str, Any] | None = None):
        widget = (widgets if widgets is not None else self.form_widgets).get(column_name)
        if widget is None:
            return

        if isinstance(widget, QSpinBox):
            widget.setValue(self.safe_sort_int(value) if value not in ("", None) else 0)
            return

        text = "" if value is None else str(value)
        if isinstance(widget, QComboBox):
            index = self.find_combo_index(widget, text, column_name)
            widget.setCurrentIndex(max(index, 0))
            return

        widget.setText(text)

    @staticmethod
    def find_combo_index(widget: QComboBox, text: str, column_name: str) -> int:
        index = widget.findText(text)
        if index >= 0:
            return index

        if column_name == "사업명":
            for row in range(widget.count()):
                if widget.itemText(row).endswith(f" {text}"):
                    return row
        if column_name == "기관명":
            for row in range(widget.count()):
                if widget.itemText(row).startswith(f"{text} ("):
                    return row
        return -1

    def delete_selected_rows(self, table: QTableWidget):
        rows = sorted({index.row() for index in table.selectedIndexes()}, reverse=True)
        for row in rows:
            table.removeRow(row)

    @staticmethod
    def read_json(path: Path, fallback: Any) -> Any:
        try:
            with path.open("r", encoding="utf-8") as file:
                data = json.load(file)
            return data
        except Exception:
            return fallback

    @staticmethod
    def write_json(path: Path, data: Any):
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=2)

    @staticmethod
    def project_label(project: dict[str, Any]) -> str:
        return f"{project.get('사업년도', '')}-{project.get('순번', '')} {project.get('사업명', '')}".strip()

    @staticmethod
    def organization_label(organization: dict[str, Any]) -> str:
        return f"{organization.get('기관명', '')} ({organization.get('기관코드', '')})"

    @staticmethod
    def safe_int(value: Any) -> int | str:
        try:
            return int(value)
        except (TypeError, ValueError):
            return "" if value is None else str(value)

    @staticmethod
    def safe_sort_int(value: Any) -> int:
        try:
            return int(value)
        except (TypeError, ValueError):
            return 999999

    @staticmethod
    def clean_master_value(value: Any) -> str:
        if pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def set_species_status(self, message: str):
        if hasattr(self, "species_status_label"):
            self.species_status_label.setText(message)
            self.species_status_label.setVisible(bool(message.strip()))

    def set_save_status(self, message: str):
        if hasattr(self, "save_status_label"):
            self.save_status_label.setText(message)
            self.save_status_label.setToolTip(message)
