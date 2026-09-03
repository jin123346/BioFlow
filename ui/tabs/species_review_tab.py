from __future__ import annotations

from datetime import datetime

import pandas as pd
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
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
        delete_button = QPushButton("선택 행 삭제")
        delete_button.clicked.connect(self.delete_selected_rows)
        export_button = QPushButton("엑셀로 저장")
        export_button.clicked.connect(self.export_to_excel)
        self.status_label = QLabel("")

        control_layout.addWidget(collect_button)
        control_layout.addWidget(import_button)
        control_layout.addWidget(add_button)
        control_layout.addWidget(delete_button)
        control_layout.addWidget(export_button)
        control_layout.addWidget(self.status_label)
        control_layout.addStretch()
        layout.addLayout(control_layout)

        self.table = QTableWidget(0, len(DISPLAY_COLUMNS))
        self.table.setHorizontalHeaderLabels(DISPLAY_COLUMNS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.resizeColumnsToContents()
        for column_name in ("일반적특징(크기,체색과무늬,주요형질 등)", "생태적특징(서식지,먹이습성,행동습성,국내분포,국외분포 등)"):
            self.table.setColumnWidth(DISPLAY_COLUMNS.index(column_name), 320)
        layout.addWidget(self.table)

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

    def add_species_row(self, korean_name: str = "", scientific_name: str = ""):
        row = self.table.rowCount()
        self.table.insertRow(row)

        status_combo = QComboBox()
        status_combo.addItems(STATUS_OPTIONS)
        status_combo.currentTextChanged.connect(self.on_status_combo_changed)
        self.table.setCellWidget(row, 0, status_combo)

        values = {column: "" for column in TEMPLATE_COLUMNS}
        values["국명"] = korean_name
        values["학명"] = scientific_name
        for offset, column_name in enumerate(TEMPLATE_COLUMNS, start=1):
            self.table.setItem(row, offset, QTableWidgetItem(values[column_name]))

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
        if column_name not in {"이전 국명", "이전 학명"}:
            return

        combo = self.table.cellWidget(item.row(), 0)
        if not isinstance(combo, QComboBox) or combo.currentText() != "업데이트 필요":
            return

        self.apply_update_match(item.row())

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

            table_row = self.table.rowCount()
            self.table.insertRow(table_row)

            matched = self.match_master(korean_name, scientific_name or korean_name)
            status_combo = QComboBox()
            status_combo.addItems(STATUS_OPTIONS)
            status_combo.setCurrentText("업데이트 필요" if matched is not None else "신규")
            status_combo.currentTextChanged.connect(self.on_status_combo_changed)
            self.table.setCellWidget(table_row, 0, status_combo)

            values = {column: "" for column in TEMPLATE_COLUMNS}
            values["이전 국명"] = korean_name
            values["이전 학명"] = scientific_name
            was_blocked = self.table.blockSignals(True)
            try:
                for offset, column_name in enumerate(TEMPLATE_COLUMNS, start=1):
                    self.table.setItem(table_row, offset, QTableWidgetItem(values[column_name]))
            finally:
                self.table.blockSignals(was_blocked)

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
