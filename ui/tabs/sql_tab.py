"""NARIS SQL authoring, with no database execution or NAS access."""

from collections import Counter
from copy import deepcopy
import json
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QDialog, QHBoxLayout, QHeaderView, QInputDialog,
    QLabel, QLineEdit, QMessageBox, QPushButton, QSpinBox, QSplitter, QTableWidget,
    QTableWidgetItem, QTabWidget, QTextEdit, QVBoxLayout, QWidget,
)

from config.paths import SETTINGS_DIR
from config.sql_media_profiles import PROFILES, default_profile
from services.sql_excel_service import ExcelSource, export_sql_review
from services.sql_generation_service import auto_map_profile, generate_sql
from ui.widgets.recent_file_dialog import RecentFileDialog as QFileDialog
from ui.widgets.sql_mapping_dialog import SQLMappingDialog

SETTINGS_FILE = SETTINGS_DIR / "sql_media_mappings.json"


def fill_table(table, headers, rows):
    table.clear()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    for row, values in enumerate(rows):
        for col, value in enumerate(values):
            table.setItem(row, col, QTableWidgetItem("" if value is None else str(value)))
    table.resizeColumnsToContents()


class SqlTab(QWidget):
    def __init__(self):
        super().__init__()
        self.source = self.sheet_data = self.results = self.profile = None
        try:
            saved = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            self.saved_profiles = saved if isinstance(saved, dict) else {}
        except (OSError, ValueError):
            self.saved_profiles = {}
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("NARIS 미디어 INSERT SQL 생성"))
        row = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        upload = QPushButton("엑셀 선택")
        upload.clicked.connect(self.choose_file)
        row.addWidget(self.file_path, 1)
        row.addWidget(upload)
        layout.addLayout(row)
        row = QHBoxLayout()
        self.sheets = QComboBox()
        self.sheets.currentIndexChanged.connect(self.change_sheet)
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 1048576)
        self.header_row.valueChanged.connect(self.refresh_sheet)
        self.preview_count = QSpinBox()
        self.preview_count.setRange(1, 500)
        self.preview_count.setValue(20)
        self.preview_count.valueChanged.connect(self.refresh_preview)
        for label, widget in (("시트", self.sheets), ("헤더 행", self.header_row), ("미리보기 행 수", self.preview_count)):
            row.addWidget(QLabel(label))
            row.addWidget(widget)
        layout.addLayout(row)
        self.info = QLabel("엑셀을 선택해 주세요.")
        self.info.setWordWrap(True)
        layout.addWidget(self.info)
        self.columns_info = QTextEdit()
        self.columns_info.setReadOnly(True)
        self.columns_info.setMaximumHeight(65)
        layout.addWidget(self.columns_info)
        buttons = QHBoxLayout()
        self.generate_button = QPushButton("SQL 작성")
        self.generate_button.setEnabled(False)
        self.generate_button.clicked.connect(self.generate)
        self.save_button = QPushButton("TXT + 검수 XLSX 저장")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_results)
        buttons.addWidget(self.generate_button)
        buttons.addWidget(self.save_button)
        buttons.addStretch()
        layout.addLayout(buttons)
        self.summary = QLabel("실제 DB 접속·SQL 실행·NAS 파일 이동은 수행하지 않습니다.")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)
        self.tabs = QTabWidget()
        self.preview = QTableWidget()
        self.preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabs.addTab(self.preview, "데이터 미리보기")
        output = QSplitter(Qt.Vertical)
        self.result_table = QTableWidget()
        self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_table.itemSelectionChanged.connect(self.show_sql)
        self.sql_preview = QTextEdit()
        self.sql_preview.setReadOnly(True)
        self.sql_preview.setLineWrapMode(QTextEdit.NoWrap)
        output.addWidget(self.result_table)
        output.addWidget(self.sql_preview)
        self.tabs.addTab(output, "검증 결과 / SQL 미리보기")
        layout.addWidget(self.tabs, 1)

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "SQL용 엑셀 선택", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.load_file(path)

    def load_file(self, path):
        self.info.setText("엑셀 파일을 확인하는 중...")
        self.generate_button.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        QApplication.processEvents()
        try:
            source = ExcelSource(path)
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "엑셀 불러오기 실패", str(exc))
            return
        try:
            self.source = source
            self.file_path.setText(str(source.path))
            self.sheets.blockSignals(True)
            self.sheets.clear()
            for name, hidden in source.sheets:
                self.sheets.addItem(name + (" (숨김)" if hidden else ""), name)
            self.sheets.setCurrentIndex(0)
            self.sheets.blockSignals(False)
            self.change_sheet()
        finally:
            QApplication.restoreOverrideCursor()

    def change_sheet(self, *_):
        if self.source is None or self.sheets.currentData() is None:
            return
        try:
            suggested = self.source.suggest_header_row(self.sheets.currentData())
        except Exception as exc:
            self.clear_results()
            self.sheet_data = None
            self.generate_button.setEnabled(False)
            self.info.setText(f"헤더 행 확인 실패: {exc}")
            return
        self.header_row.blockSignals(True)
        self.header_row.setValue(suggested)
        self.header_row.blockSignals(False)
        self.refresh_sheet()

    def clear_results(self):
        self.results = self.profile = None
        self.result_table.clear()
        self.result_table.setRowCount(0)
        self.sql_preview.clear()
        self.save_button.setEnabled(False)
        self.summary.setText("DB 연결 키·기존 DB 중복·실제 NAS 파일 존재 여부는 정적 검증 범위에 포함되지 않습니다.")

    def refresh_sheet(self, *_):
        self.clear_results()
        self.sheet_data = None
        self.generate_button.setEnabled(False)
        if self.source is None or self.sheets.currentData() is None:
            return
        try:
            self.sheet_data = self.source.read_sheet(self.sheets.currentData(), self.header_row.value())
        except Exception as exc:
            self.info.setText(f"시트 읽기 실패: {exc}")
            fill_table(self.preview, [], [])
            self.columns_info.clear()
            return
        sheet = self.sheet_data
        legacy = " · xls는 시트/셀값을 XLSX로 변환하며 수식은 계산 결과로 유지" if self.source.legacy else ""
        self.info.setText(f"{self.source.path.name} · 시트: {sheet.name} · 데이터 {len(sheet.rows):,}행 · {len(sheet.headers)}컬럼{legacy}")
        self.columns_info.setPlainText("컬럼: " + " / ".join(sheet.headers))
        self.refresh_preview()
        self.generate_button.setEnabled(bool(sheet.rows))
        self.tabs.setCurrentIndex(0)

    def refresh_preview(self, *_):
        if self.sheet_data is not None:
            sheet = self.sheet_data
            rows = [[number, *[row.get(header) for header in sheet.headers]] for number, row in sheet.rows[:self.preview_count.value()]]
            fill_table(self.preview, ["엑셀 행", *sheet.headers], rows)

    def generate(self):
        if self.sheet_data is None:
            return
        kinds = list(PROFILES)
        names = [PROFILES[kind]["label"] for kind in kinds]
        name, ok = QInputDialog.getItem(self, "미디어 유형 선택", "SQL을 생성할 미디어 유형", names, 0, False)
        if not ok:
            return
        kind = kinds[names.index(name)]
        profile = deepcopy(self.saved_profiles.get(kind, default_profile(kind)))
        try:
            profile = auto_map_profile(profile, self.sheet_data.headers)
            dialog = SQLMappingDialog(profile, self.sheet_data.headers, self)
            dialog.set_preview_rows(self.sheet_data.rows)
            if dialog.exec() != QDialog.Accepted:
                return
            profile = dialog.values()
            results = generate_sql(self.sheet_data, profile)
        except Exception as exc:
            QMessageBox.critical(self, "SQL 생성 실패", str(exc))
            return
        self.profile, self.results = profile, results
        self.saved_profiles[kind] = deepcopy(profile)
        try:
            SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
            temporary = SETTINGS_FILE.with_suffix(".tmp")
            temporary.write_text(json.dumps(self.saved_profiles, ensure_ascii=False, indent=2), encoding="utf-8")
            temporary.replace(SETTINGS_FILE)
        except OSError as exc:
            QMessageBox.warning(self, "매핑 설정 저장", f"SQL은 생성했지만 설정을 저장하지 못했습니다: {exc}")
        counts = Counter(result.status for result in results)
        self.summary.setText(f"{name} · SUCCESS {counts['SUCCESS']} / ERROR {counts['ERROR']} / SKIP {counts['SKIP']} · 표는 최대 500행 표시. 전체 결과는 검수 XLSX에 저장됩니다.")
        fill_table(self.result_table, ["엑셀 행", "상태", "오류 / 제외 사유"],
                   [[item.excel_row, item.status, item.error] for item in results[:500]])
        self.result_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.save_button.setEnabled(True)
        self.tabs.setCurrentIndex(1)
        if results:
            self.result_table.selectRow(0)

    def show_sql(self):
        row = self.result_table.currentRow()
        if self.results is not None and 0 <= row < len(self.results):
            result = self.results[row]
            self.sql_preview.setPlainText(result.sql or result.error)

    def save_results(self):
        if self.results is None or self.sheet_data is None:
            return
        default = f"NARIS_{self.profile['label']}_{self.source.path.stem}_SQL.txt"
        path, _ = QFileDialog.getSaveFileName(self, "실행용 SQL 및 검수 엑셀 저장", default, "SQL text (*.txt *.sql)")
        if not path:
            return
        if Path(path).suffix.lower() not in (".txt", ".sql"):
            path += ".txt"
        review = Path(path).with_name(f"{Path(path).stem}_검수.xlsx")
        existing = [str(item) for item in (Path(path), review) if item.exists()]
        if existing and QMessageBox.question(self, "결과 파일 덮어쓰기", "다음 결과 파일을 덮어쓸까요?\n" + "\n".join(existing),
                                            QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        try:
            text, workbook = export_sql_review(self.source, self.sheet_data, self.results, self.profile, path, overwrite=bool(existing))
            QMessageBox.information(self, "SQL 저장 완료", f"실행용 SQL과 검수용 엑셀을 저장했습니다.\n{text}\n{workbook}\n\n하나의 DB 세션에서 순차 실행하세요. DB 실행은 프로그램에서 수행하지 않습니다.")
        except Exception as exc:
            QMessageBox.critical(self, "SQL 저장 실패", str(exc))
