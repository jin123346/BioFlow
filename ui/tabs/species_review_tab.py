from __future__ import annotations

from datetime import datetime

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from services.species_master_service import SpeciesMasterService
from ui.tabs.upload_tab import ROW_COLUMNS
from utils.text_utils import clean_master_value, normalize_text

TEMPLATE_COLUMNS = [
    "기존ID값",
    "이전 국명",
    "이전 학명",
    "국명",
    "학명",
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
        import_button = QPushButton("이전 국명/학명 엑셀 불러오기")
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

        control_layout.addWidget(collect_button)
        control_layout.addWidget(import_button)
        control_layout.addWidget(add_button)
        control_layout.addWidget(update_button)
        control_layout.addWidget(delete_button)
        control_layout.addWidget(export_button)
        control_layout.addWidget(self.status_label)
        control_layout.addStretch()
        layout.addLayout(control_layout)

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
            if column_name in {"이전 국명", "이전 학명"}:
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
        korean_name = values.get("이전 국명", "")
        scientific_name = values.get("이전 학명", "")
        if (korean_name or scientific_name) and self.match_master(korean_name, scientific_name or korean_name) is not None:
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
        korean_name = self.form_widgets["이전 국명"].text().strip()
        scientific_name = self.form_widgets["이전 학명"].text().strip()
        if not korean_name and not scientific_name:
            return

        matched = self.match_master(korean_name, scientific_name or korean_name)
        if matched is None:
            return

        self.form_widgets[STATUS_COLUMN].setCurrentText("업데이트 필요")
        self.form_widgets["기존ID값"].setText(self.matched_species_id(matched))
        if not scientific_name:
            self.form_widgets["이전 학명"].setText(normalize_text(matched.get("학명")))
        self.form_widgets["clssc_sstem_kor"].setText(normalize_text(matched.get("분류체계명(국문)")))
        self.form_widgets["clssc_sstem_latin"].setText(normalize_text(matched.get("분류체계명(라틴)")))

    def current_table_row(self) -> int | None:
        rows = {index.row() for index in self.table.selectedIndexes()}
        if len(rows) != 1:
            return None
        return rows.pop()

    def write_row(self, row: int, values: dict[str, str]):
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
        if column_name in {"이전 국명", "이전 학명"}:
            combo = self.table.cellWidget(item.row(), 0)
            if isinstance(combo, QComboBox) and combo.currentText() == "업데이트 필요":
                self.apply_update_match(item.row())

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
        if not korean_name and not scientific_name:
            return

        matched = self.match_master(korean_name, scientific_name or korean_name)
        if matched is None:
            self.status_label.setText(f"'{korean_name or scientific_name}' 마스터와 정확히 일치하는 항목이 없어 비워둠")
            return

        self.fill_matched_row(row, matched)
        matched_id = self.matched_species_id(matched)
        self.status_label.setText(f"'{korean_name or scientific_name}' → 기존ID {matched_id} 자동 매칭 (정확히 일치)")

    def fill_matched_row(self, row: int, matched):
        # 이전 국명/이전 학명 칸에 다시 쓰지는 않음(itemChanged 재귀 방지 + 원본 유지),
        # 매칭으로 얻는 값(기존ID/분류체계명)만 채우고 이전 학명은 비어있을 때만 채움.
        id_col = TEMPLATE_COLUMNS.index("기존ID값") + 1
        latin_col = TEMPLATE_COLUMNS.index("이전 학명") + 1
        clssc_kor_col = TEMPLATE_COLUMNS.index("clssc_sstem_kor") + 1
        clssc_latin_col = TEMPLATE_COLUMNS.index("clssc_sstem_latin") + 1

        was_blocked = self.table.blockSignals(True)
        try:
            self.table.setItem(row, id_col, QTableWidgetItem(self.matched_species_id(matched)))

            latin_item = self.table.item(row, latin_col)
            if not (latin_item and latin_item.text().strip()):
                self.table.setItem(row, latin_col, QTableWidgetItem(normalize_text(matched.get("학명"))))

            self.table.setItem(row, clssc_kor_col, QTableWidgetItem(normalize_text(matched.get("분류체계명(국문)"))))
            self.table.setItem(row, clssc_latin_col, QTableWidgetItem(normalize_text(matched.get("분류체계명(라틴)"))))
        finally:
            self.table.blockSignals(was_blocked)

    def import_previous_names_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "이전 국명/학명 엑셀 불러오기", "", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
        except Exception as exc:
            QMessageBox.critical(self, "엑셀 불러오기 실패", str(exc))
            return

        korean_column = next((c for c in ("이전 국명", "국명") if c in df.columns), None)
        if korean_column is None:
            QMessageBox.warning(self, "엑셀 불러오기", "'국명' 컬럼을 찾을 수 없습니다.")
            return
        latin_column = next((c for c in ("이전 학명", "학명") if c in df.columns), None)

        added = 0
        for _, row_data in df.iterrows():
            korean_name = normalize_text(row_data.get(korean_column, ""))
            if not korean_name:
                continue
            scientific_name = normalize_text(row_data.get(latin_column, "")) if latin_column else ""

            matched = self.match_master(korean_name, scientific_name or korean_name)
            values = {column: "" for column in TEMPLATE_COLUMNS}
            values["이전 국명"] = korean_name
            values["이전 학명"] = scientific_name
            values[STATUS_COLUMN] = "업데이트 필요" if matched is not None else "신규"
            table_row = self._insert_row(values)

            if matched is not None:
                self.fill_matched_row(table_row, matched)
            added += 1

        if added == 0:
            QMessageBox.information(self, "엑셀 불러오기", "가져올 국명이 없습니다.")
            return
        self.status_label.setText(f"엑셀에서 {added}건 가져옴 (기존ID/이전 학명/분류체계명 자동 매칭, 확인 필요)")

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

        if korean_name and "국명" in master_df.columns:
            names = master_df["국명"].apply(normalize_text)
            exact = master_df[names == korean_name]
            if len(exact):
                return exact.iloc[0]

        if scientific_name and "학명" in master_df.columns:
            sci_names = master_df["학명"].apply(normalize_text)
            exact = master_df[sci_names == scientific_name]
            if len(exact):
                return exact.iloc[0]

        return None

    def delete_selected_rows(self):
        rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.table.removeRow(row)

    def export_to_excel(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "엑셀 저장", "저장할 행이 없습니다.")
            return

        records = []
        for row in range(self.table.rowCount()):
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
