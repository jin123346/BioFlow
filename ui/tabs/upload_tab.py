from __future__ import annotations

import json
import logging
import re
import shutil
import subprocess
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from PySide6.QtCore import Qt, QTimer

logger = logging.getLogger("bioflow.upload_tab")
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QCompleter,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from config.paths import DATA_DIR
from services.address_master_service import AddressMasterService
from services.species_master_service import SpeciesMasterService
from utils.text_utils import normalize_text

try:
    from PIL import ExifTags, Image
except ImportError:
    ExifTags = None
    Image = None


REFERENCE_DIR = DATA_DIR / "reference"
PROJECTS_FILE = REFERENCE_DIR / "upload_projects.json"
ORGANIZATIONS_FILE = REFERENCE_DIR / "organizations.json"
DRAFT_FILE = REFERENCE_DIR / "new_data_draft.json"

ROW_COLUMNS = [
    "NO.",
    "종고유ID",
    "미디어타입",
    "미디어코드",
    "미디어파일",
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
    "분류코드": ["IN", "PL", "BI", "PI", "MM"],
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

AUTO_FIELDS = {"NO.", "미디어코드", "사업년도", "순번", "기관코드", "데이터베이스번호"}
DEFAULT_ORGANIZATION_NAME = "국립중앙과학관"

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


class DataTableWidget(QTableWidget):
    """Ctrl+V로 엑셀에서 복사한 여러 셀을 한 번에 붙여넣을 수 있는 표."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.paste_requested = None

    def keyPressEvent(self, event):
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_V and self.paste_requested:
            self.paste_requested()
            return
        super().keyPressEvent(event)


class UploadTab(QWidget):
    def __init__(self):
        super().__init__()
        self.projects: list[dict[str, Any]] = []
        self.organizations: list[dict[str, Any]] = []
        self.project_lookup: dict[str, dict[str, Any]] = {}
        self.organization_lookup: dict[str, dict[str, Any]] = {}
        self.form_widgets: dict[str, QLineEdit | QComboBox | QSpinBox] = {}
        self.species_master_df = pd.DataFrame()
        self.species_name_lookup: dict[str, list[dict[str, Any]]] = {}
        self.master_service = SpeciesMasterService()
        self.address_master = AddressMasterService()
        self._last_selected_row: int | None = None
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setSingleShot(True)
        self.autosave_timer.timeout.connect(self.autosave_draft)
        self.autosave_suspended = False
        self.excel_file_path = ""
        self.media_storage_dir = ""

        self._ensure_reference_files()
        self.load_reference_data()
        self.load_species_master_lookup()
        self.init_ui()
        self.load_draft()

    def init_ui(self):
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        title_label = QLabel("신규 데이터 직접 작성")
        main_layout.addWidget(title_label)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.tabs.addTab(self._build_row_editor_tab(), "행 작성")
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
        lookup_species_button = QPushButton("국명 조회")
        lookup_species_button.clicked.connect(self.fill_species_from_korean_name)
        lookup_address_button = QPushButton("위도/경도로 주소 조회")
        lookup_address_button.clicked.connect(self.fill_address_from_latlon)
        self.species_status_label = QLabel("")

        control_layout.addWidget(new_form_button)
        control_layout.addWidget(import_folder_button)
        control_layout.addWidget(update_row_button)
        control_layout.addWidget(delete_row_button)
        control_layout.addWidget(lookup_species_button)
        control_layout.addWidget(lookup_address_button)
        control_layout.addWidget(self.species_status_label)
        control_layout.addStretch()
        layout.addLayout(control_layout)

        media_storage_layout = QHBoxLayout()
        self.media_storage_path_input = QLineEdit()
        self.media_storage_path_input.setReadOnly(True)
        self.media_storage_path_input.setPlaceholderText("미디어 파일을 DB번호 이름으로 저장할 폴더")
        media_storage_button = QPushButton("미디어 저장 위치 선택")
        media_storage_button.clicked.connect(self.select_media_storage_dir)
        media_storage_layout.addWidget(QLabel("미디어 저장 위치"))
        media_storage_layout.addWidget(self.media_storage_path_input, 1)
        media_storage_layout.addWidget(media_storage_button)
        layout.addLayout(media_storage_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        form_container = QWidget()
        self.row_form = QFormLayout(form_container)
        self.row_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self.build_row_form()
        scroll_area.setWidget(form_container)
        layout.addWidget(scroll_area, 2)

        next_row_layout = QHBoxLayout()
        next_row_layout.addStretch()
        add_row_button = QPushButton("입력")
        add_row_button.clicked.connect(self.add_data_row_from_form)
        next_row_button = QPushButton("다음 행")
        next_row_button.clicked.connect(self.go_to_next_row)
        next_row_layout.addWidget(add_row_button)
        next_row_layout.addWidget(next_row_button)
        layout.addLayout(next_row_layout)

        self.data_table = DataTableWidget(0, len(ROW_COLUMNS))
        self.data_table.setHorizontalHeaderLabels(ROW_COLUMNS)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.data_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.data_table.itemSelectionChanged.connect(self.load_selected_row_to_form)
        self.data_table.itemChanged.connect(self.on_data_table_item_changed)
        self.data_table.paste_requested = self.paste_clipboard_into_table
        # ponytail: ResizeToContents는 셀 값이 바뀔 때마다 해당 컬럼 전체(수천 행)의
        # 텍스트 폭을 다시 재는 탓에 행이 많아지면 한 번 쓰기에 10초 이상 걸림.
        # Interactive로 바꾸고 최초 1회만 폭을 맞춤 (이후 사용자가 드래그로 조절 가능).
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.data_table.resizeColumnsToContents()
        layout.addWidget(self.data_table, 3)

        save_layout = QHBoxLayout()
        save_draft_button = QPushButton("작성 내용 저장")
        save_draft_button.clicked.connect(self.save_draft)
        load_draft_button = QPushButton("임시저장 불러오기")
        load_draft_button.clicked.connect(self.reload_draft)
        export_button = QPushButton("엑셀로 저장")
        export_button.clicked.connect(self.export_rows_to_excel)
        import_button = QPushButton("엑셀 불러오기")
        import_button.clicked.connect(self.load_rows_from_excel)
        thumbnail_button = QPushButton("동영상 캡처 이미지 생성")
        thumbnail_button.clicked.connect(self.generate_video_thumbnails)
        clear_button = QPushButton("작성 내용 비우기")
        clear_button.clicked.connect(self.clear_data_rows)
        self.save_status_label = QLabel("")

        save_layout.addStretch()
        save_layout.addWidget(save_draft_button)
        save_layout.addWidget(load_draft_button)
        save_layout.addWidget(export_button)
        save_layout.addWidget(import_button)
        save_layout.addWidget(thumbnail_button)
        save_layout.addWidget(clear_button)
        save_layout.addWidget(self.save_status_label)
        layout.addLayout(save_layout)
        self.refresh_combos()
        return tab

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
                if column_name in {"대륙/해양 국문", "대륙/해양 영문"}:
                    widget.setEditable(True)
                widget.currentTextChanged.connect(self.refresh_generated_fields)
                if column_name == "미디어타입":
                    widget.currentTextChanged.connect(self.update_media_type_field_visibility)
            else:
                widget = QLineEdit()
                if column_name in {"동정년월일", "관찰일"}:
                    widget.setPlaceholderText("2016-01-01")
                elif column_name == "위도":
                    widget.setPlaceholderText("36.123456")
                elif column_name == "경도":
                    widget.setPlaceholderText("126.123456")
                elif column_name == "종고유ID":
                    widget.setPlaceholderText("종정보 마스터의 종고유ID")
                elif column_name == "미디어파일":
                    widget.setPlaceholderText("원본 미디어 선택 버튼으로 선택")
                elif column_name in {"동영상 재생시간", "사운드 재생시간"}:
                    widget.setPlaceholderText("예: 0:45")
                elif column_name in {"동영상 파일명", "사운드 파일명"}:
                    widget.setPlaceholderText("미디어 저장 시 자동 입력, 직접 수정 가능")
                elif column_name == "국명":
                    widget.setPlaceholderText("입력 후 Enter 또는 국명 조회")
                    widget.editingFinished.connect(self.fill_species_from_korean_name)
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
            if column_name in AUTO_FIELDS:
                widget.setEnabled(False)
                if isinstance(widget, QLineEdit):
                    widget.setReadOnly(True)
                if isinstance(widget, QSpinBox):
                    widget.setButtonSymbols(QSpinBox.NoButtons)
            self.form_widgets[column_name] = widget
            self.connect_autosave_signal(widget)
            if column_name == "미디어파일":
                media_layout = QHBoxLayout()
                media_layout.addWidget(widget, 1)
                select_media_button = QPushButton("원본 선택")
                select_media_button.clicked.connect(self.select_media_source_file)
                media_layout.addWidget(select_media_button)
                self.row_form.addRow(label, media_layout)
            else:
                self.row_form.addRow(label, widget)
        self.update_media_type_field_visibility()

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

    def load_species_master_lookup(self):
        try:
            df = self.master_service.load_last_master()
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
        lookup: dict[str, list[dict[str, Any]]] = {}
        for _, row in df.iterrows():
            korean_name = normalize_text(row.get("국명"))
            if not korean_name:
                continue
            lookup.setdefault(korean_name, []).append(row.to_dict())
        self.species_name_lookup = lookup

    def rebuild_lookups(self):
        self.project_lookup = {self.project_label(project): project for project in self.projects}
        self.organization_lookup = {self.organization_label(org): org for org in self.organizations}

    def refresh_combos(self):
        if "사업명" not in self.form_widgets or "기관명" not in self.form_widgets:
            return

        self.rebuild_lookups()
        self.refresh_combo_items("사업명", list(self.project_lookup.keys()))
        self.refresh_combo_items("기관명", list(self.organization_lookup.keys()))
        self.set_default_organization()
        self.apply_project_to_form()
        self.apply_organization_to_form()

    def refresh_combo_items(self, column_name: str, labels: list[str]):
        widget = self.form_widgets.get(column_name)
        if not isinstance(widget, QComboBox):
            return

        selected_value = self.get_form_value(column_name) if widget.count() else ""
        widget.blockSignals(True)
        widget.clear()
        widget.addItems(labels)
        selected_index = self.find_combo_index(widget, selected_value, column_name)
        if selected_index >= 0:
            widget.setCurrentIndex(selected_index)
        widget.blockSignals(False)

    def apply_project_to_form(self, label: str = ""):
        if not label and isinstance(self.form_widgets.get("사업명"), QComboBox):
            label = self.form_widgets["사업명"].currentText()
        project = self.project_lookup.get(label, {})
        if not project:
            return
        self.set_form_value("사업년도", project.get("사업년도", datetime.now().year))
        self.set_form_value("순번", project.get("순번", 0))
        self.refresh_generated_fields()

    def apply_organization_to_form(self, label: str = ""):
        if not label and isinstance(self.form_widgets.get("기관명"), QComboBox):
            label = self.form_widgets["기관명"].currentText()
        organization = self.organization_lookup.get(label, {})
        if not organization:
            return
        self.set_form_value("기관코드", organization.get("기관코드", ""))
        self.refresh_generated_fields()

    def set_default_organization(self):
        widget = self.form_widgets.get("기관명")
        if not isinstance(widget, QComboBox):
            return
        for row in range(widget.count()):
            if widget.itemText(row).startswith(f"{DEFAULT_ORGANIZATION_NAME} ("):
                widget.setCurrentIndex(row)
                return

    def refresh_generated_fields(self):
        media_type = self.get_form_value("미디어타입")
        media_code = MEDIA_TYPE_CODES.get(media_type, "")
        self.set_form_value("미디어코드", media_code)

        organization_code = self.get_form_value("기관코드")
        media_code = self.get_form_value("미디어코드")
        taxon_code = self.get_form_value("분류코드")
        year = self.get_form_value("사업년도")
        order = self.get_form_value("순번")
        no = self.get_form_value("NO.")
        if organization_code and media_code and taxon_code and year and order and no:
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
        if item.column() == ROW_COLUMNS.index("NO."):
            self.renumber_rows_below(item.row())
        if item.row() == self.current_data_row():
            self.load_selected_row_to_form()
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

    def clear_form(self):
        for column_name in ROW_COLUMNS:
            if column_name in STICKY_FIELDS:
                continue
            if column_name == "NO.":
                self.set_form_value(column_name, self.next_no_value())
            elif column_name == "관찰개체수":
                self.set_form_value(column_name, 1)
            elif column_name in DROPDOWN_DEFAULTS:
                self.set_form_value(column_name, DROPDOWN_DEFAULTS[column_name])
            else:
                self.set_form_value(column_name, "")
        self.apply_project_to_form()
        self.set_default_organization()
        self.apply_organization_to_form()
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
        self.set_form_value("NO.", self.next_no_value())
        self.refresh_generated_fields()

    def next_no_value(self) -> int:
        max_no = 0
        no_column = ROW_COLUMNS.index("NO.")
        for row in range(self.data_table.rowCount()):
            item = self.data_table.item(row, no_column)
            max_no = max(max_no, self.safe_sort_int(item.text() if item else 0))
        return max_no + 1

    def fill_species_from_korean_name(self):
        korean_name = normalize_text(self.get_form_value("국명"))
        if not korean_name:
            self.set_species_status("")
            return

        if not self.species_name_lookup:
            self.set_species_status("종정보 마스터 없음")
            return

        matches = self.species_name_lookup.get(korean_name, [])
        if not matches:
            self.set_species_status("마스터 매칭 없음")
            return

        match = matches[0]
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

        self.autosave_suspended = True
        try:
            # 폼에 이전에 작성하던 국명/학명 등 내용이 남아있으면 새로 만드는 모든 행에
            # 그대로 찍혀버리므로, 사업명/기관/분류코드/미디어타입만 남기고 비운 뒤 시작
            self.clear_form()
            for file_path, media_type in media_files:
                self.set_form_value("미디어타입", media_type)
                self.set_form_value("NO.", self.next_no_value())
                self.set_form_value("미디어파일", str(file_path))
                self.apply_location_from_media(file_path, media_type)
                self.apply_duration_from_media(file_path, media_type)
                self.apply_date_from_media(file_path, media_type)
                self.add_data_row(self.collect_form_values())
        finally:
            self.autosave_suspended = False

        self.autosave_draft()
        QMessageBox.information(
            self,
            "폴더 불러오기",
            f"{len(media_files)}건의 미디어를 행으로 추가했습니다.\n"
            "각 행을 선택해 종명 등 나머지 정보를 입력한 뒤 '선택 행 수정'으로 저장해 주세요.",
        )
        self.set_save_status(f"폴더 불러오기 완료: {len(media_files)}건 (나머지 정보 입력 필요)")

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

        self.set_form_value("미디어파일", file_path)
        self.apply_location_from_media(Path(file_path), media_type)
        self.apply_duration_from_media(Path(file_path), media_type)
        self.apply_date_from_media(Path(file_path), media_type)
        self.schedule_autosave()

    def apply_location_from_media(self, source_path: Path, media_type: str):
        latitude, longitude, message = self.extract_location_from_media(source_path, media_type)
        if latitude is None or longitude is None:
            self.set_save_status(message)
            return

        self.set_form_value("위도", f"{latitude:.7f}")
        self.set_form_value("경도", f"{longitude:.7f}")
        self.set_save_status(f"미디어 위치정보 적용: 위도 {latitude:.7f}, 경도 {longitude:.7f}")

    def apply_duration_from_media(self, source_path: Path, media_type: str):
        duration_field = {"동영상": "동영상 재생시간", "사운드": "사운드 재생시간"}.get(media_type)
        if not duration_field:
            return
        duration_text = self.extract_media_duration(source_path)
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
    def extract_media_duration(file_path: Path) -> str:
        ffprobe = shutil.which("ffprobe")
        if not ffprobe:
            return ""
        data = UploadTab.run_json_command(
            [ffprobe, "-v", "quiet", "-print_format", "json", "-show_format", str(file_path)],
            timeout=10,
            empty_json="{}",
        )

        duration = UploadTab.to_float(data.get("format", {}).get("duration"))
        if duration is None:
            return ""

        total_seconds = int(round(duration))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours:
            return f"{hours}:{minutes:02d}:{seconds:02d}"
        return f"{minutes}:{seconds:02d}"

    def apply_date_from_media(self, source_path: Path, media_type: str):
        date_text = self.extract_media_date(source_path, media_type)
        if date_text:
            self.set_form_value("관찰일", date_text)

    @staticmethod
    def extract_media_date(file_path: Path, media_type: str) -> str:
        if media_type == "이미지":
            date_text = UploadTab.extract_image_date(file_path)
            if date_text:
                return date_text
        return UploadTab.extract_media_date_with_ffprobe(file_path)

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
    def extract_media_date_with_ffprobe(file_path: Path) -> str:
        ffprobe = shutil.which("ffprobe")
        if not ffprobe:
            return ""
        data = UploadTab.run_json_command(
            [ffprobe, "-v", "quiet", "-print_format", "json", "-show_format", str(file_path)],
            timeout=10,
            empty_json="{}",
        )
        creation_time = data.get("format", {}).get("tags", {}).get("creation_time")
        if not creation_time:
            return ""
        return str(creation_time).split("T")[0]

    def prepare_media_for_record(self, values: dict[str, Any]) -> dict[str, Any]:
        media_path_text = normalize_text(values.get("미디어파일"))
        if not media_path_text:
            return values

        source_path = Path(media_path_text)
        if not source_path.exists():
            return values

        saved_path = self.save_media_file_to_storage(source_path, values)
        if saved_path:
            values = values.copy()
            values["미디어파일"] = str(saved_path)
            self.set_form_value("미디어파일", str(saved_path))
            self.apply_location_from_media(saved_path, values.get("미디어타입", ""))
            values["위도"] = self.get_form_value("위도")
            values["경도"] = self.get_form_value("경도")

            filename_field = {"동영상": "동영상 파일명", "사운드": "사운드 파일명"}.get(values.get("미디어타입"))
            if filename_field:
                values[filename_field] = saved_path.name
                self.set_form_value(filename_field, saved_path.name)
        return values

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

        expected_stem = self.safe_filename(db_number)
        if source_path.parent.resolve() == target_dir.resolve() and source_path.stem == expected_stem:
            return source_path

        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            target_path = self.unique_media_path(target_dir, expected_stem, source_path.suffix.lower())
            t0 = time.perf_counter()
            shutil.copy2(source_path, target_path)
            logger.info(
                "save_media_file_to_storage: copy %.3fs (%d bytes) -> %s",
                time.perf_counter() - t0,
                source_path.stat().st_size,
                target_path,
            )
            # ponytail: 동영상 캡처(ffmpeg 프로세스 실행)는 행 하나 저장할 때마다 부르면
            # 무거워서 UI가 멈춤 -> 여기서는 안 하고, "동영상 캡처 이미지 생성" 버튼으로
            # 한번에 일괄 처리(generate_video_thumbnails)하도록 분리.
            self.set_save_status(f"미디어 저장 완료: {target_path}")
            return target_path
        except Exception as exc:
            QMessageBox.critical(self, "미디어 저장 실패", str(exc))
            self.set_save_status(f"미디어 저장 실패: {exc}")
            return None

    @staticmethod
    def capture_video_thumbnail(video_path: Path, output_path: Path, at_seconds: float = 1.0) -> bool:
        ffmpeg = shutil.which("ffmpeg")
        if not ffmpeg:
            return False
        try:
            result = subprocess.run(
                [ffmpeg, "-y", "-ss", str(at_seconds), "-i", str(video_path), "-frames:v", "1", str(output_path)],
                capture_output=True,
                timeout=20,
                check=False,
            )
        except Exception:
            return False
        return result.returncode == 0 and output_path.exists()

    def generate_video_thumbnails(self):
        media_column = ROW_COLUMNS.index("미디어파일")
        type_column = ROW_COLUMNS.index("미디어타입")

        video_paths = []
        for row in range(self.data_table.rowCount()):
            type_item = self.data_table.item(row, type_column)
            media_item = self.data_table.item(row, media_column)
            if not type_item or type_item.text() != "동영상":
                continue
            if not media_item or not media_item.text():
                continue
            video_path = Path(media_item.text())
            if video_path.exists():
                video_paths.append(video_path)

        if not video_paths:
            QMessageBox.information(self, "동영상 캡처", "캡처할 동영상 행이 없습니다 (미디어 저장이 먼저 필요합니다).")
            return

        success, skipped, failed = 0, 0, 0
        for video_path in video_paths:
            thumbnail_path = video_path.with_suffix(".jpg")
            if thumbnail_path.exists():
                skipped += 1
                continue
            if self.capture_video_thumbnail(video_path, thumbnail_path):
                success += 1
            else:
                failed += 1

        QMessageBox.information(
            self, "동영상 캡처", f"완료: {success}건 생성, {skipped}건 이미 있음, {failed}건 실패(ffmpeg 필요)."
        )
        self.set_save_status(f"동영상 캡처 이미지 생성 완료: {success}건 생성 / {failed}건 실패")

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

    def extract_location_from_media(self, file_path: Path, media_type: str) -> tuple[float | None, float | None, str]:
        if media_type == "이미지":
            latitude, longitude = self.extract_image_gps(file_path)
            if latitude is not None and longitude is not None:
                return latitude, longitude, "이미지 EXIF 위치정보를 적용했습니다."

        latitude, longitude = self.extract_location_with_exiftool(file_path)
        if latitude is not None and longitude is not None:
            return latitude, longitude, "미디어 메타데이터 위치정보를 적용했습니다."

        latitude, longitude = self.extract_location_with_ffprobe(file_path)
        if latitude is not None and longitude is not None:
            return latitude, longitude, "동영상 메타데이터 위치정보를 적용했습니다."

        if media_type in {"동영상", "사운드"} and not shutil.which("exiftool") and not shutil.which("ffprobe"):
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
    def extract_location_with_ffprobe(file_path: Path) -> tuple[float | None, float | None]:
        ffprobe = shutil.which("ffprobe")
        if not ffprobe:
            return None, None
        data = UploadTab.run_json_command(
            [ffprobe, "-v", "quiet", "-print_format", "json", "-show_format", str(file_path)],
            timeout=10,
            empty_json="{}",
        )

        tags = data.get("format", {}).get("tags", {})
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

    def add_data_row(self, values: dict[str, Any]):
        row = self.data_table.rowCount()
        self.data_table.insertRow(row)
        self.write_data_row(row, values)

    def write_data_row(self, row: int, values: dict[str, Any]):
        was_blocked = self.data_table.blockSignals(True)
        try:
            for col, column_name in enumerate(ROW_COLUMNS):
                value = values.get(column_name, "")
                self.data_table.setItem(row, col, QTableWidgetItem("" if value is None else str(value)))
        finally:
            self.data_table.blockSignals(was_blocked)

    def current_data_row(self) -> int | None:
        indexes = self.data_table.selectedIndexes()
        if not indexes:
            return None
        return indexes[0].row()

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

        row = self.current_data_row()
        start_row = 0 if row is None else row
        start_col = self.data_table.currentColumn()
        start_col = 0 if start_col < 0 else start_col

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
        records = self.table_to_records(self.data_table, ROW_COLUMNS)
        self.write_draft(records, self.collect_form_values())
        self.write_current_excel(records)
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
            self.excel_file_path = draft.get("excel_path", "")
            if self.excel_file_path:
                self.set_save_status(f"엑셀 자동반영 대상: {self.excel_file_path}")
            self.media_storage_dir = draft.get("media_storage_dir", "")
            if self.media_storage_dir and hasattr(self, "media_storage_path_input"):
                self.media_storage_path_input.setText(self.media_storage_dir)
        finally:
            self.autosave_suspended = False

    def reload_draft(self):
        self.autosave_suspended = True
        try:
            self.data_table.setRowCount(0)
        finally:
            self.autosave_suspended = False
        self.load_draft()
        self.set_save_status("임시저장을 다시 불러왔습니다.")

    def export_rows_to_excel(self):
        self.sync_current_form_to_table()
        records = self.table_to_records(self.data_table, ROW_COLUMNS)
        if not records:
            QMessageBox.warning(self, "엑셀 저장", "저장할 행이 없습니다.")
            return

        default_name = f"new_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "신규 데이터 엑셀 저장", default_name, "Excel Files (*.xlsx)")
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            self.write_excel_file(file_path, records)
            final_path = self.final_excel_path(file_path)
            self.write_final_excel_file(str(final_path), records)
            self.reset_after_export()
            QMessageBox.information(
                self,
                "엑셀 저장",
                f"신규 데이터 엑셀을 저장했습니다.\n{file_path}\n{final_path}\n\n"
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
        self.excel_file_path = ""
        self.clear_form()
        self.write_draft([], self.collect_form_values())
        self.set_save_status("엑셀 저장 완료. 새 작성을 시작합니다 (NO. 1부터).")

    def load_rows_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "신규 데이터 엑셀 불러오기", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 불러오기 실패", str(exc))
            return

        missing_columns = [column for column in ROW_COLUMNS if column not in df.columns]
        if missing_columns:
            QMessageBox.warning(
                self, "엑셀 불러오기", f"'작성한 그대로' 형식의 엑셀이 아닙니다. 없는 컬럼: {', '.join(missing_columns)}"
            )
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

        self.excel_file_path = file_path
        self.autosave_draft()
        self.set_save_status(f"엑셀 불러오기 완료: {file_path} ({len(df)}건)")

    def sync_current_form_to_table(self):
        values = self.collect_form_values()
        if not self.has_form_content(values):
            return
        values = self.prepare_media_for_record(values)

        row = self.current_data_row()
        if row is not None:
            self.write_data_row(row, values)
            return

        if not self.is_duplicate_record(values):
            self.add_data_row(values)

    def has_form_content(self, values: dict[str, Any]) -> bool:
        ignored_fields = {"사업명", "사업년도", "순번", "기관명", "기관코드", "미디어코드", "데이터베이스번호"}
        for key, value in values.items():
            if key in ignored_fields:
                continue
            if value not in ("", 0):
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
            "excel_path": self.excel_file_path,
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
            "excel_path": str(data.get("excel_path", "")),
            "media_storage_dir": str(data.get("media_storage_dir", "")),
        }

    def write_current_excel(self, records: list[dict[str, Any]]):
        if not self.excel_file_path:
            return
        try:
            self.write_excel_file(self.excel_file_path, records)
            self.write_final_excel_file(str(self.final_excel_path(self.excel_file_path)), records)
            self.set_save_status(f"엑셀 자동반영 완료: {self.excel_file_path}")
        except Exception as exc:
            self.set_save_status(f"엑셀 자동반영 실패: {exc}")

    @staticmethod
    def write_excel_file(file_path: str, records: list[dict[str, Any]]):
        pd.DataFrame(records, columns=ROW_COLUMNS).to_excel(file_path, index=False)

    @staticmethod
    def final_excel_path(file_path: str) -> Path:
        path = Path(file_path)
        return path.with_name(f"{path.stem}_완성본{path.suffix}")

    @staticmethod
    def write_final_excel_file(file_path: str, records: list[dict[str, Any]]):
        final_records = [
            {target: (record.get(source, "") if source else "") for target, source in FINAL_EXPORT_COLUMNS}
            for record in records
        ]
        final_columns = [target for target, _ in FINAL_EXPORT_COLUMNS]
        pd.DataFrame(final_records, columns=final_columns).to_excel(file_path, index=False)

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
                has_value = has_value or bool(value)
            if has_value:
                records.append(record)
        return records

    def get_form_value(self, column_name: str) -> Any:
        widget = self.form_widgets[column_name]
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

    def set_form_value(self, column_name: str, value: Any):
        widget = self.form_widgets.get(column_name)
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

    def set_save_status(self, message: str):
        if hasattr(self, "save_status_label"):
            self.save_status_label.setText(message)
