from __future__ import annotations

import json
import re
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
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

MEDIA_TYPE_CODES = {
    "정보": "OB",
    "이미지": "OI",
    "동영상": "OM",
    "사운드": "OS",
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
        add_row_button = QPushButton("입력")
        add_row_button.clicked.connect(self.add_data_row_from_form)
        update_row_button = QPushButton("선택 행 수정")
        update_row_button.clicked.connect(self.update_selected_data_row)
        delete_row_button = QPushButton("선택 행 삭제")
        delete_row_button.clicked.connect(self.delete_selected_data_rows)
        lookup_species_button = QPushButton("국명 조회")
        lookup_species_button.clicked.connect(self.fill_species_from_korean_name)
        self.species_status_label = QLabel("")

        control_layout.addWidget(new_form_button)
        control_layout.addWidget(add_row_button)
        control_layout.addWidget(update_row_button)
        control_layout.addWidget(delete_row_button)
        control_layout.addWidget(lookup_species_button)
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

        self.data_table = QTableWidget(0, len(ROW_COLUMNS))
        self.data_table.setHorizontalHeaderLabels(ROW_COLUMNS)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.data_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.data_table.itemSelectionChanged.connect(self.load_selected_row_to_form)
        self.data_table.itemChanged.connect(self.on_data_table_item_changed)
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(self.data_table, 3)

        save_layout = QHBoxLayout()
        save_draft_button = QPushButton("작성 내용 저장")
        save_draft_button.clicked.connect(self.save_draft)
        load_draft_button = QPushButton("임시저장 불러오기")
        load_draft_button.clicked.connect(self.reload_draft)
        export_button = QPushButton("엑셀로 저장")
        export_button.clicked.connect(self.export_rows_to_excel)
        clear_button = QPushButton("작성 내용 비우기")
        clear_button.clicked.connect(self.clear_data_rows)
        self.save_status_label = QLabel("")

        save_layout.addStretch()
        save_layout.addWidget(save_draft_button)
        save_layout.addWidget(load_draft_button)
        save_layout.addWidget(export_button)
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
                widget.currentTextChanged.connect(self.refresh_generated_fields)
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
                elif column_name == "국명":
                    widget.setPlaceholderText("입력 후 Enter 또는 국명 조회")
                    widget.editingFinished.connect(self.fill_species_from_korean_name)
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
        row = self.current_data_row()
        if row is None:
            return
        self.autosave_suspended = True
        for col, column_name in enumerate(ROW_COLUMNS):
            item = self.data_table.item(row, col)
            self.set_form_value(column_name, item.text() if item else "")
        self.autosave_suspended = False

    def on_data_table_item_changed(self, item: QTableWidgetItem):
        if self.autosave_suspended:
            return
        if item.row() == self.current_data_row():
            self.load_selected_row_to_form()
        self.schedule_autosave()

    def clear_form(self):
        for column_name in ROW_COLUMNS:
            if column_name in {"사업명", "사업년도", "순번", "기관명", "기관코드"}:
                continue
            if column_name == "NO.":
                self.set_form_value(column_name, self.next_no_value())
            elif column_name == "관찰개체수":
                self.set_form_value(column_name, 1)
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

    def select_media_source_file(self):
        media_type = self.get_form_value("미디어타입")
        file_filter = self.media_file_filter(media_type)
        file_path, _ = QFileDialog.getOpenFileName(self, "원본 미디어 선택", "", file_filter)
        if not file_path:
            return

        self.set_form_value("미디어파일", file_path)
        self.apply_location_from_media(Path(file_path), media_type)
        self.schedule_autosave()

    def apply_location_from_media(self, source_path: Path, media_type: str):
        latitude, longitude, message = self.extract_location_from_media(source_path, media_type)
        if latitude is None or longitude is None:
            self.set_save_status(message)
            return

        self.set_form_value("위도", f"{latitude:.7f}")
        self.set_form_value("경도", f"{longitude:.7f}")
        self.set_save_status(f"미디어 위치정보 적용: 위도 {latitude:.7f}, 경도 {longitude:.7f}")

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
            shutil.copy2(source_path, target_path)
            self.set_save_status(f"미디어 저장 완료: {target_path}")
            return target_path
        except Exception as exc:
            QMessageBox.critical(self, "미디어 저장 실패", str(exc))
            self.set_save_status(f"미디어 저장 실패: {exc}")
            return None

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
        try:
            result = subprocess.run(
                [exiftool, "-j", "-n", str(file_path)],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            data = json.loads(result.stdout or "[]")
        except Exception:
            return None, None

        if not data:
            return None, None
        item = data[0]
        return UploadTab.to_float(item.get("GPSLatitude")), UploadTab.to_float(item.get("GPSLongitude"))

    @staticmethod
    def extract_location_with_ffprobe(file_path: Path) -> tuple[float | None, float | None]:
        ffprobe = shutil.which("ffprobe")
        if not ffprobe:
            return None, None
        try:
            result = subprocess.run(
                [ffprobe, "-v", "quiet", "-print_format", "json", "-show_format", str(file_path)],
                capture_output=True,
                text=True,
                timeout=10,
                check=False,
            )
            data = json.loads(result.stdout or "{}")
        except Exception:
            return None, None

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
            self.excel_file_path = file_path
            self.write_draft(records, self.collect_form_values())
            self.set_save_status(f"엑셀 저장 및 자동반영 설정: {file_path}")
            QMessageBox.information(self, "엑셀 저장", f"신규 데이터 엑셀을 저장했습니다.\n{file_path}")
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 저장 실패", str(exc))

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
        self.autosave_suspended = True
        try:
            row = self.current_data_row()
            if row is not None:
                values = self.collect_form_values()
                if self.has_form_content(values):
                    self.write_data_row(row, values)
            records = self.table_to_records(self.data_table, ROW_COLUMNS)
            self.write_draft(
                records,
                self.collect_form_values(),
            )
            self.write_current_excel(records)
            self.set_save_status("자동 임시저장 완료")
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
            self.set_save_status(f"엑셀 자동반영 완료: {self.excel_file_path}")
        except Exception as exc:
            self.set_save_status(f"엑셀 자동반영 실패: {exc}")

    @staticmethod
    def write_excel_file(file_path: str, records: list[dict[str, Any]]):
        pd.DataFrame(records, columns=ROW_COLUMNS).to_excel(file_path, index=False)

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
