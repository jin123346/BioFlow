from __future__ import annotations

from datetime import datetime
import re

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QInputDialog,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ui.widgets.recent_file_dialog import RecentFileDialog as QFileDialog
from services.species_master_service import SpeciesMasterService
from services.classification_service import ClassificationService, RANK_FIELDS, import_reference_excel, read_version
from ui.tabs.upload_tab import ROW_COLUMNS
from utils.text_utils import clean_master_value, normalize_text

TEMPLATE_COLUMNS = [
    "기존ID값",
    "이전 국명 및 학명",
    "이전 국명",
    "이전 학명",
    "국명",
    "학명",
    "clss_id",
    "clssc_sstem_kor",
    "clssc_sstem_latin",
    "일반적특징(크기,체색과무늬,주요형질 등)",
    "생태적특징(서식지,먹이습성,행동습성,국내분포,국외분포 등)",
    "계 영명",
    "계 국문",
    "문 영명",
    "문 국문",
    "강 영명",
    "강 국문",
    "목 영명",
    "목 국문",
    "과 영명",
    "과 국문",
    "속 영명(동정되는 부분까지 입력)",
    "속 국문",
    "종 영명(동정되는 부분까지 입력)",
    "종 국문",
]

STATUS_COLUMN = "구분"
STATUS_OPTIONS = ["신규", "업데이트 필요"]
DISPLAY_COLUMNS = [STATUS_COLUMN] + TEMPLATE_COLUMNS


class SpeciesReviewTab(QWidget):
    def __init__(self, upload_tab):
        super().__init__()
        self.upload_tab = upload_tab
        self.master_service = SpeciesMasterService()
        self.classification_service = ClassificationService()
        self._master_df: pd.DataFrame | None = None
        self.form_widgets: dict[str, QWidget] = {}
        self._last_selected_row: int | None = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        title_label = QLabel("신규/업데이트 종 정리")
        layout.addWidget(title_label)

        control_layout = QHBoxLayout()
        collect_button = QPushButton("관찰정보에서 미매칭 종 가져오기")
        collect_button.clicked.connect(self.collect_unmatched_species)
        import_button = QPushButton("엑셀 불러오기·컬럼 매핑")
        import_button.clicked.connect(self.import_previous_names_from_excel)
        add_button = QPushButton("행 직접 추가")
        add_button.clicked.connect(self.add_blank_row)
        update_button = QPushButton("선택 행 수정")
        update_button.clicked.connect(self.update_selected_row_from_form)
        delete_button = QPushButton("선택 행 삭제")
        delete_button.clicked.connect(self.delete_selected_rows)
        export_button = QPushButton("엑셀로 저장")
        export_button.clicked.connect(self.export_to_excel)
        self.status_label = QLabel("")
        self.upload_tab.configure_status_label(self.status_label)

        control_layout.addWidget(collect_button)
        control_layout.addWidget(import_button)
        control_layout.addWidget(add_button)
        control_layout.addWidget(update_button)
        control_layout.addWidget(delete_button)
        control_layout.addWidget(export_button)
        control_layout.addStretch()
        layout.addLayout(control_layout)
        layout.addWidget(self.status_label)
        reference_layout = QHBoxLayout()
        reference_button = QPushButton("분류군 기준자료 업데이트")
        reference_button.clicked.connect(self.import_classification_reference)
        self.reference_label = QLabel()
        self.reference_label.setWordWrap(True)
        reference_layout.addWidget(reference_button)
        reference_layout.addWidget(self.reference_label, 1)
        layout.addLayout(reference_layout)
        self.refresh_reference_label()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        form_container = QWidget()
        self.review_form = QFormLayout(form_container)
        self.review_form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self.build_input_form()
        scroll_area.setWidget(form_container)

        add_row_layout = QHBoxLayout()
        add_row_layout.addStretch()
        add_form_row_button = QPushButton("입력")
        add_form_row_button.clicked.connect(self.add_row_from_form)
        add_row_layout.addWidget(add_form_row_button)

        top_pane = QWidget()
        top_pane_layout = QVBoxLayout(top_pane)
        top_pane_layout.setContentsMargins(0, 0, 0, 0)
        top_pane_layout.addWidget(scroll_area, 1)
        top_pane_layout.addLayout(add_row_layout)

        self.table = QTableWidget(0, len(DISPLAY_COLUMNS))
        self.table.setHorizontalHeaderLabels(DISPLAY_COLUMNS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.itemSelectionChanged.connect(self.load_selected_row_to_form)
        self.table.resizeColumnsToContents()
        for column_name in ("일반적특징(크기,체색과무늬,주요형질 등)", "생태적특징(서식지,먹이습성,행동습성,국내분포,국외분포 등)"):
            self.table.setColumnWidth(DISPLAY_COLUMNS.index(column_name), 320)

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(top_pane)
        splitter.addWidget(self.table)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter, 1)

    def build_input_form(self):
        status_combo = QComboBox()
        status_combo.addItems(STATUS_OPTIONS)
        self.form_widgets[STATUS_COLUMN] = status_combo
        self.review_form.addRow(STATUS_COLUMN, status_combo)
        for column_name in TEMPLATE_COLUMNS:
            widget = QLineEdit()
            self.form_widgets[column_name] = widget
            self.review_form.addRow(column_name, widget)
            if column_name in {field for pair in RANK_FIELDS for field in pair}:
                widget.editingFinished.connect(self.refresh_form_classification)
            if column_name in {"clss_id", "clssc_sstem_kor", "clssc_sstem_latin"}:
                widget.setReadOnly(True)
            if column_name in {"이전 국명", "이전 학명", "이전 국명 및 학명"}:
                # 이전 국명/이전 학명을 적고 다른 칸으로 넘어가는 순간, 클릭해서
                # "입력"을 누르기 전에도 미리 마스터 매칭 결과를 폼에서 바로 보여줌.
                widget.editingFinished.connect(self.refresh_form_match)

    def collect_form_values(self) -> dict[str, str]:
        values = {}
        for column_name, widget in self.form_widgets.items():
            values[column_name] = widget.currentText() if isinstance(widget, QComboBox) else widget.text().strip()
        return values

    def clear_input_form(self):
        for column_name, widget in self.form_widgets.items():
            if isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
            else:
                widget.clear()

    def add_row_from_form(self):
        values = self.collect_form_values()
        if not any(values[column_name] for column_name in TEMPLATE_COLUMNS):
            QMessageBox.information(self, "행 추가", "입력된 값이 없습니다.")
            return

        # 엑셀 불러오기와 동일하게, 이전 국명/이전 학명을 적었으면 "구분"을 직접
        # 안 골라도 마스터에서 매칭되는 순간 자동으로 "업데이트 필요"로 잡아줌.
        if self.match_previous(values) is not None:
            values[STATUS_COLUMN] = "업데이트 필요"

        row = self._insert_row(values)
        if values[STATUS_COLUMN] == "업데이트 필요":
            self.apply_update_match(row)
        else:
            self.status_label.setText("입력한 값으로 행을 추가했습니다.")
        self.clear_input_form()

    def refresh_form_match(self):
        # 폼의 이전 국명/이전 학명만으로 마스터를 조회해서, 아직 표에 행을 넣기 전이라도
        # 기존ID값/clssc값을 폼에 바로 채워줌 (apply_update_match의 표 버전과 동일한 매칭 규칙).
        scientific_name = self.form_widgets["이전 학명"].text().strip()
        matched = self.match_previous(self.collect_form_values())
        if matched is None:
            return

        self.form_widgets[STATUS_COLUMN].setCurrentText("업데이트 필요")
        self.form_widgets["기존ID값"].setText(self.matched_species_id(matched))
        if not scientific_name:
            self.form_widgets["이전 학명"].setText(normalize_text(matched.get("학명")))
        self.refresh_form_classification()

    def current_table_row(self) -> int | None:
        rows = {index.row() for index in self.table.selectedIndexes()}
        if len(rows) != 1:
            return None
        return rows.pop()

    def write_row(self, row: int, values: dict[str, str]):
        values = self.classify_values(values)
        combo = self.table.cellWidget(row, 0)
        if isinstance(combo, QComboBox):
            was_blocked = combo.blockSignals(True)
            combo.setCurrentText(values.get(STATUS_COLUMN, STATUS_OPTIONS[0]))
            combo.blockSignals(was_blocked)

        was_blocked = self.table.blockSignals(True)
        try:
            for offset, column_name in enumerate(TEMPLATE_COLUMNS, start=1):
                self.table.setItem(row, offset, QTableWidgetItem(values.get(column_name, "")))
        finally:
            self.table.blockSignals(was_blocked)

    def load_selected_row_to_form(self):
        previous_row = self._last_selected_row
        row = self.current_table_row()

        # 선택 행이 바뀌는 시점에만 이전에 폼에서 편집하던 내용을 표에 반영
        # (upload_tab의 load_selected_row_to_form과 동일한 패턴)
        if previous_row is not None and previous_row != row and previous_row < self.table.rowCount():
            values = self.collect_form_values()
            if any(values[column_name] for column_name in TEMPLATE_COLUMNS):
                self.write_row(previous_row, values)

        self._last_selected_row = row
        if row is None:
            return

        combo = self.table.cellWidget(row, 0)
        self.form_widgets[STATUS_COLUMN].setCurrentText(combo.currentText() if isinstance(combo, QComboBox) else STATUS_OPTIONS[0])
        for column_name in TEMPLATE_COLUMNS:
            offset = TEMPLATE_COLUMNS.index(column_name) + 1
            item = self.table.item(row, offset)
            self.form_widgets[column_name].setText(item.text() if item else "")

    def update_selected_row_from_form(self):
        row = self.current_table_row()
        if row is None:
            QMessageBox.warning(self, "선택 행 수정", "수정할 행을 먼저 선택해 주세요.")
            return
        values = self.collect_form_values()
        self.write_row(row, values)
        if values[STATUS_COLUMN] == "업데이트 필요":
            self.apply_update_match(row)
        else:
            self.status_label.setText("선택 행을 수정했습니다.")

    def _ensure_master_loaded(self) -> pd.DataFrame:
        if self._master_df is None:
            try:
                self._master_df = self.master_service.load_last_master()
            except Exception:
                self._master_df = pd.DataFrame()
        return self._master_df

    def collect_unmatched_species(self):
        records = self.upload_tab.table_to_records(self.upload_tab.data_table, ROW_COLUMNS)

        seen: set[str] = set()
        added = 0
        for record in records:
            korean_name = normalize_text(record.get("국명"))
            species_id = normalize_text(record.get("종고유ID"))
            if not korean_name or species_id or korean_name in seen:
                continue
            seen.add(korean_name)
            self.add_species_row(korean_name, normalize_text(record.get("학명")))
            added += 1

        if added == 0:
            QMessageBox.information(self, "가져오기", "매칭 실패한(종고유ID 없는) 국명이 없습니다.")
            return
        self.status_label.setText(f"{added}건 가져옴")

    def add_blank_row(self):
        self.add_species_row()

    def _insert_row(self, values: dict[str, str]) -> int:
        values = self.classify_values(values)
        row = self.table.rowCount()
        self.table.insertRow(row)

        status_combo = QComboBox()
        status_combo.addItems(STATUS_OPTIONS)
        status_combo.setCurrentText(values.get(STATUS_COLUMN, STATUS_OPTIONS[0]))
        status_combo.currentTextChanged.connect(self.on_status_combo_changed)
        self.table.setCellWidget(row, 0, status_combo)

        was_blocked = self.table.blockSignals(True)
        try:
            for offset, column_name in enumerate(TEMPLATE_COLUMNS, start=1):
                self.table.setItem(row, offset, QTableWidgetItem(values.get(column_name, "")))
        finally:
            self.table.blockSignals(was_blocked)
        return row

    def add_species_row(self, korean_name: str = "", scientific_name: str = ""):
        values = {column: "" for column in TEMPLATE_COLUMNS}
        values["국명"] = korean_name
        values["학명"] = scientific_name
        self._insert_row(values)

    def on_status_combo_changed(self, text: str):
        combo = self.sender()
        for row in range(self.table.rowCount()):
            if self.table.cellWidget(row, 0) is combo:
                # 국명/학명 칸은 "수정 후 값"을 적는 칸이라 여기선 절대 조회 트리거로 안 씀.
                # 이미 "이전 국명"/"이전 학명"에 값이 있으면 그 값으로만 조회.
                if text == "업데이트 필요":
                    self.apply_update_match(row)
                break

    def on_item_changed(self, item: QTableWidgetItem):
        column_name = DISPLAY_COLUMNS[item.column()] if item.column() < len(DISPLAY_COLUMNS) else None
        # 기존ID값 조회는 "이전 국명"/"이전 학명" 칸을 적었을 때만 일어남.
        # 국명/학명 칸은 수정된(새) 값을 적는 칸이라 여기서 바뀌어도 기존ID값에 영향 없음.
        if column_name in {"이전 국명", "이전 학명", "이전 국명 및 학명"}:
            combo = self.table.cellWidget(item.row(), 0)
            if isinstance(combo, QComboBox) and combo.currentText() == "업데이트 필요":
                self.apply_update_match(item.row())
        if column_name in {field for pair in RANK_FIELDS for field in pair}:
            self.apply_classification_match(item.row())

        # 표에서 직접 고친 값(매칭 결과 포함)을 위 폼에도 바로 반영 -- 지금 그 행이
        # 선택돼 있을 때만. write-back 루프 방지를 위해 load_selected_row_to_form이
        # previous_row == row일 땐 표에 다시 쓰지 않고 그냥 폼만 새로고침함.
        if item.row() == self._last_selected_row:
            self.load_selected_row_to_form()

    def previous_name_values(self, row: int) -> tuple[str, str]:
        korean_col = TEMPLATE_COLUMNS.index("이전 국명") + 1
        latin_col = TEMPLATE_COLUMNS.index("이전 학명") + 1
        korean_item = self.table.item(row, korean_col)
        latin_item = self.table.item(row, latin_col)
        return (
            korean_item.text().strip() if korean_item else "",
            latin_item.text().strip() if latin_item else "",
        )

    def apply_update_match(self, row: int):
        korean_name, scientific_name = self.previous_name_values(row)

        values = self.row_values(row)
        combined = values.get("이전 국명 및 학명", "")
        if not korean_name and not scientific_name and not combined:
            return
        matched = self.match_previous(values)
        if matched is None:
            self.status_label.setText(f"'{korean_name or scientific_name or combined}' 기존 종을 유일하게 확인하지 못했습니다. 기존ID값을 확인해 주세요.")
            return

        self.fill_matched_row(row, matched)
        matched_id = self.matched_species_id(matched)
        self.status_label.setText(f"'{korean_name or scientific_name}' → 기존ID {matched_id} 자동 매칭 (정확히 일치)")

    def fill_matched_row(self, row: int, matched):
        # 이전 국명/이전 학명 칸에 다시 쓰지는 않음(itemChanged 재귀 방지 + 원본 유지),
        # 매칭으로 얻는 값(기존ID/분류체계명)만 채우고 이전 학명은 비어있을 때만 채움.
        id_col = TEMPLATE_COLUMNS.index("기존ID값") + 1
        latin_col = TEMPLATE_COLUMNS.index("이전 학명") + 1

        was_blocked = self.table.blockSignals(True)
        try:
            self.table.setItem(row, id_col, QTableWidgetItem(self.matched_species_id(matched)))

            latin_item = self.table.item(row, latin_col)
            if not (latin_item and latin_item.text().strip()):
                self.table.setItem(row, latin_col, QTableWidgetItem(normalize_text(matched.get("학명"))))

        finally:
            self.table.blockSignals(was_blocked)
        self.apply_classification_match(row)

    def import_previous_names_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "이전 국명/학명 엑셀 불러오기", "", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        try:
            df = self.upload_tab.select_excel_sheet(file_path)
            if df is None:
                return
            aliases = [("목 국명", "목 국문"), ("과 국명", "과 국문"),
                       ("속 국명", "속 국문"), ("종 국명", "종 국문"),
                       ("clssc_sstem_info_id", "clss_id")]
            df = self.upload_tab.map_excel_columns(df, DISPLAY_COLUMNS, aliases, "species_review")
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 불러오기 실패", str(exc))
            return

        if df is None:
            return

        added = 0
        for row_data in df.to_dict("records"):
            values = {column: clean_master_value(row_data.get(column, "")) for column in TEMPLATE_COLUMNS}
            if not any(values.values()):
                continue
            matched = self.match_previous(values)
            requested = clean_master_value(row_data.get(STATUS_COLUMN, ""))
            previous = any(values.get(key) for key in ("이전 국명", "이전 학명", "이전 국명 및 학명", "기존ID값"))
            values[STATUS_COLUMN] = "업데이트 필요" if previous or requested == "업데이트 필요" else "신규"
            table_row = self._insert_row(values)

            if matched is not None:
                self.fill_matched_row(table_row, matched)
            added += 1

        if added == 0:
            QMessageBox.information(self, "엑셀 불러오기", "가져올 국명이 없습니다.")
            return
        matched_count = sum(bool(self.row_values(row).get("clss_id")) for row in range(self.table.rowCount() - added, self.table.rowCount()))
        self.status_label.setText(f"엑셀에서 {added}건 가져옴. 과 기준 분류ID 매칭 {matched_count}/{added}건. 빈 ID는 분류정보를 확인해 주세요.")

    def classify_values(self, values):
        values = dict(values)
        match, reason = self.classification_service.match(values)
        values["clss_id"] = str(match["clssc_sstem_info_id"]) if match else ""
        values["clssc_sstem_kor"] = match["taxonomy_name"] if match else ""
        values["clssc_sstem_latin"] = match["taxonomy_latin"] if match else ""
        return values

    def refresh_form_classification(self):
        values = self.collect_form_values()
        match, reason = self.classification_service.match(values)
        classified = self.classify_values(values)
        for field in ("clss_id", "clssc_sstem_kor", "clssc_sstem_latin"):
            self.form_widgets[field].setText(classified[field])
            self.form_widgets[field].setToolTip(reason)

    def apply_classification_match(self, row):
        values = self.row_values(row)
        _, reason = self.classification_service.match(values)
        classified = self.classify_values(values)
        blocked = self.table.blockSignals(True)
        try:
            for field in ("clss_id", "clssc_sstem_kor", "clssc_sstem_latin"):
                item = QTableWidgetItem(classified[field])
                item.setToolTip(reason)
                self.table.setItem(row, TEMPLATE_COLUMNS.index(field) + 1, item)
        finally:
            self.table.blockSignals(blocked)

    def row_values(self, row):
        return {column: (self.table.item(row, index + 1).text() if self.table.item(row, index + 1) else "")
                for index, column in enumerate(TEMPLATE_COLUMNS)}

    def match_previous(self, values):
        korean = values.get("이전 국명", "")
        latin = values.get("이전 학명", "")
        combined = values.get("이전 국명 및 학명", "").strip()
        probes = [(korean, latin)] if korean or latin else []
        if combined:
            candidates = [combined] + re.split(r"[/,;|\n]+", combined)
            words = combined.split()
            candidates += [" ".join(words[:i]) for i in range(1, len(words))]
            candidates += [" ".join(words[i:]) for i in range(1, len(words))]
            probes += [(text.strip(), text.strip()) for text in candidates if text.strip()]
        matches = {}
        for kr, sci in probes:
            match = self.match_master(kr, sci)
            if match is not None:
                matches[self.matched_species_id(match)] = match
        return next(iter(matches.values())) if len(matches) == 1 else None

    def refresh_reference_label(self):
        self.reference_label.setText(f"분류군 자료 버전: {read_version() or '미등록'} · 과와 입력된 상위 분류 기준으로 clss_id 자동 매칭")

    def import_classification_reference(self):
        path, _ = QFileDialog.getOpenFileName(self, "분류군 기준자료 선택", "", "Excel Files (*.xlsx)")
        if not path:
            return
        version, accepted = QInputDialog.getText(self, "분류군 자료 버전", "새 자료 버전 (예: 2026-10-13)", text=datetime.now().strftime("%Y-%m-%d"))
        if not accepted:
            return
        try:
            result = import_reference_excel(path, version)
            self.refresh_reference_label()
            self.status_label.setText(f"분류군 기준자료 {result['version']} 등록: {result['rows']}건")
        except Exception as exc:
            QMessageBox.warning(self, "분류군 등록 실패", str(exc))

    @staticmethod
    def matched_species_id(matched) -> str:
        # 마스터에서 "종고유ID"가 미설정일 때 빈칸 대신 0으로 채워둔 행이 있어서,
        # ""뿐 아니라 "0"도 무효로 보고 "종학명정보ID"로 대체 조회 (upload_tab과 동일 규칙)
        primary_id = clean_master_value(matched.get("종고유ID", ""))
        if primary_id not in {"", "0"}:
            return primary_id
        return clean_master_value(matched.get("종학명정보ID", ""))

    def match_master(self, korean_name: str, scientific_name: str):
        # 국명/학명 모두 정확히 일치할 때만 매칭. 조금이라도 다르면 None을 반환해서
        # 호출부가 기존ID값/이전 학명/clssc_sstem_* 값을 빈칸으로 두게 함(유사매칭 금지).
        master_df = self._ensure_master_loaded()
        if master_df is None or master_df.empty:
            return None

        korean_name = normalize_text(korean_name)
        scientific_name = normalize_text(scientific_name)

        if getattr(self, "_indexed_master", None) is not master_df:
            self._indexed_master = master_df
            self._name_indexes = {"국명": {}, "학명": {}}
            for record in master_df.to_dict("records"):
                for field, index in self._name_indexes.items():
                    key = normalize_text(record.get(field))
                    if key:
                        index.setdefault(key, []).append(record)
        matches = {}
        for field, value in (("국명", korean_name), ("학명", scientific_name)):
            for record in self._name_indexes[field].get(value, []):
                matches[self.matched_species_id(record)] = record
        return next(iter(matches.values())) if len(matches) == 1 else None

    def delete_selected_rows(self):
        rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.table.removeRow(row)

    def export_to_excel(self):
        current = self.current_table_row()
        if current is not None:
            self.write_row(current, self.collect_form_values())
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "엑셀 저장", "저장할 행이 없습니다.")
            return

        records = []
        for row in range(self.table.rowCount()):
            self.apply_classification_match(row)
            record = {}
            for offset, column_name in enumerate(TEMPLATE_COLUMNS, start=1):
                item = self.table.item(row, offset)
                record[column_name] = item.text() if item else ""
            records.append(record)

        default_name = f"species_new_update_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "신규/업데이트 종 엑셀 저장", default_name, "Excel Files (*.xlsx)")
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            pd.DataFrame(records, columns=TEMPLATE_COLUMNS).to_excel(file_path, index=False)
            QMessageBox.information(self, "엑셀 저장", f"저장 완료.\n{file_path}")
            self.status_label.setText(f"엑셀 저장 완료: {file_path}")
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 저장 실패", str(exc))
