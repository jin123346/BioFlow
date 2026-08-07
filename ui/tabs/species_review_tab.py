from __future__ import annotations

import difflib
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
    "이전 국명 및 학명",
    "국명",
    "학명",
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
        add_button = QPushButton("행 직접 추가")
        add_button.clicked.connect(self.add_blank_row)
        delete_button = QPushButton("선택 행 삭제")
        delete_button.clicked.connect(self.delete_selected_rows)
        export_button = QPushButton("엑셀로 저장")
        export_button.clicked.connect(self.export_to_excel)
        self.status_label = QLabel("")

        control_layout.addWidget(collect_button)
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
                # 이미 "이전 국명 및 학명"에 값이 있으면 그 값으로만 조회.
                if text == "업데이트 필요":
                    previous_col = TEMPLATE_COLUMNS.index("이전 국명 및 학명") + 1
                    previous_item = self.table.item(row, previous_col)
                    if previous_item and previous_item.text():
                        self.apply_update_match_from_previous_name(row, previous_item.text())
                break

    def on_item_changed(self, item: QTableWidgetItem):
        column_name = DISPLAY_COLUMNS[item.column()] if item.column() < len(DISPLAY_COLUMNS) else None
        # 기존ID값 조회는 "이전 국명 및 학명" 칸을 적었을 때만 일어남.
        # 국명/학명 칸은 수정된(새) 값을 적는 칸이라 여기서 바뀌어도 기존ID값에 영향 없음.
        if column_name != "이전 국명 및 학명":
            return

        combo = self.table.cellWidget(item.row(), 0)
        if not isinstance(combo, QComboBox) or combo.currentText() != "업데이트 필요":
            return

        self.apply_update_match_from_previous_name(item.row(), item.text())

    def apply_update_match_from_previous_name(self, row: int, text: str):
        id_col = TEMPLATE_COLUMNS.index("기존ID값") + 1

        korean_name, _, scientific_name = text.partition("/")
        korean_name = korean_name.strip()
        scientific_name = scientific_name.strip()
        if not korean_name and not scientific_name:
            return
        if not scientific_name:
            # "국명 / 학명" 구분자 없이 이름 하나만 적은 경우 국명·학명 양쪽으로 시도
            scientific_name = korean_name

        matched = self.match_master(korean_name, scientific_name)
        if matched is None:
            self.status_label.setText(f"'{text}' 마스터에서 비슷한 항목을 찾지 못했습니다.")
            return

        matched_id = self.matched_species_id(matched)
        self.table.setItem(row, id_col, QTableWidgetItem(matched_id))
        self.status_label.setText(f"'{text}' → 기존ID {matched_id} 자동 매칭 (확인 필요)")

    @staticmethod
    def matched_species_id(matched) -> str:
        # 마스터에서 "종고유ID"가 미설정일 때 빈칸 대신 0으로 채워둔 행이 있어서,
        # ""뿐 아니라 "0"도 무효로 보고 "종학명정보ID"로 대체 조회 (upload_tab과 동일 규칙)
        primary_id = clean_master_value(matched.get("종고유ID", ""))
        if primary_id not in {"", "0"}:
            return primary_id
        return clean_master_value(matched.get("종학명정보ID", ""))

    def match_master(self, korean_name: str, scientific_name: str):
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
            candidates = difflib.get_close_matches(korean_name, names.tolist(), n=1, cutoff=0.6)
            if candidates:
                return master_df[names == candidates[0]].iloc[0]

        if scientific_name and "학명" in master_df.columns:
            sci_names = master_df["학명"].apply(normalize_text)
            exact = master_df[sci_names == scientific_name]
            if len(exact):
                return exact.iloc[0]
            candidates = difflib.get_close_matches(scientific_name, sci_names.tolist(), n=1, cutoff=0.6)
            if candidates:
                return master_df[sci_names == candidates[0]].iloc[0]

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
