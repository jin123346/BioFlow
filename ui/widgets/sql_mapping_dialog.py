"""Editable NARIS column mappings and explicit confirmation of unprovided DB rules."""

from copy import deepcopy

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHeaderView, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QScrollArea, QSpinBox, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)
from services.sql_generation_service import FILENAME_COLUMNS, apply_image_sequence_rule, mapped_value, mapped_nas_path, number, filter_row
from services.sql_file_lookup_service import MediaFileIndex
from ui.widgets.recent_file_dialog import RecentFileDialog as QFileDialog


def choices(headers, selected=None):
    widget = QComboBox()
    widget.addItem("(매핑 없음 / 고정값 사용)", None)
    for header in headers:
        widget.addItem(header, header)
    widget.setCurrentIndex(max(0, widget.findData(selected)))
    return widget


def items(text):
    return [value.strip() for value in text.split(",") if value.strip()]


class SQLMappingDialog(QDialog):
    def __init__(self, profile, headers, parent=None):
        super().__init__(parent)
        profile = apply_image_sequence_rule(deepcopy(profile))
        self.profile = profile
        self.setWindowTitle(f"{profile['label']} · 매핑 및 DB 규칙 확인")
        self.resize(1150, 800)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"대상 테이블: {profile['table']}"))
        hint = QLabel("DB 타입과 최대 길이는 제공된 스키마로 고정됩니다. 관찰·표본 이미지는 각각 관찰정보순번·표본정보순번을 연결 키로 사용합니다.")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        self.filter_notice = QLabel()
        self.filter_notice.setWordWrap(True)
        layout.addWidget(self.filter_notice)
        tabs = QTabWidget()
        self.tabs = tabs
        self.preview_rows = []
        layout.addWidget(tabs, 1)
        mapping_page = QWidget()
        mapping_layout = QVBoxLayout(mapping_page)
        folder_row = QHBoxLayout()
        folder_row.addWidget(QLabel("파일명 매칭 폴더"))
        self.lookup_folder = QLineEdit(profile.get("media_lookup_dir", ""))
        self.lookup_folder.setPlaceholderText("폴더에서 DB 번호로 실제 파일명을 가져올 경우 지정")
        folder_row.addWidget(self.lookup_folder, 1)
        folder_button = QPushButton("폴더 선택")
        folder_button.clicked.connect(self.choose_lookup_folder)
        folder_row.addWidget(folder_button)
        mapping_layout.addLayout(folder_row)
        options = QHBoxLayout()
        self.recursive = QCheckBox("종고유ID 등 하위 폴더도 검색")
        self.recursive.setChecked(profile.get("lookup_recursive", True))
        self.skip_previews = QCheckBox("미리보기용 ‘썸네일’ 하위 폴더 제외")
        self.skip_previews.setChecked(profile.get("lookup_skip_previews", True))
        options.addWidget(self.recursive)
        options.addWidget(self.skip_previews)
        mapping_layout.addLayout(options)
        self.mapping = QTableWidget(len(profile["columns"]), 9)
        self.mapping.setHorizontalHeaderLabels(["DB 컬럼", "엑셀 컬럼", "고정값 / 개별 값 입력", "DB 타입", "최대 길이", "길이 단위", "필수", "확장자 (선택)", "파일명 가져오기"])
        self.mapping.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.mapping.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.controls = []
        self.extension_controls = {}
        self.lookup_controls = {}
        for row, column in enumerate(profile["columns"]):
            label = QTableWidgetItem(column["name"])
            label.setFlags(label.flags() & ~Qt.ItemIsEditable)
            self.mapping.setItem(row, 0, label)
            source = choices(headers, column.get("source"))
            constant = QLineEdit("" if column.get("constant") is None else str(column["constant"]))
            data_type = QComboBox()
            data_type.addItems(["NUMBER", "VARCHAR2", "CHAR", "DATE"])
            data_type.setCurrentText(column["type"])
            if "image" in profile["kind"] and column["name"] == profile["key"]:
                data_type.setEnabled(False)
            length = QSpinBox()
            length.setRange(0, 32767)
            length.setSpecialValueText("해당 없음" if column["type"] in ("NUMBER", "DATE") else "미확정")
            length.setValue(column.get("length") or 0)
            if column["name"] == "LVBNG_MVPCT_RPRSN_IMAGE":
                length.setMaximum(100)
            unit = QComboBox()
            unit.addItems(["BYTE", "CHAR"])
            unit.setCurrentText(column.get("length_unit", "BYTE"))
            data_type.setEnabled(False)
            length.setEnabled(False)
            unit.setEnabled(False)
            required = QCheckBox()
            required.setChecked(column.get("required", False))
            if column.get("required"):
                required.setEnabled(False)
            computed = column.get("computed")
            if computed in ("nas", "sysdate", "null") or column["name"] == "FRST_REGISTER_ID":
                source.setCurrentIndex(0)
                source.setEnabled(False)
                constant.setEnabled(False)
                constant.setPlaceholderText({"nas": "NAS 경로 설정 사용", "sysdate": "SYSDATE", "null": "NULL"}.get(computed, "narisorg"))
            if computed == "sequence":
                constant.setEnabled(False)
                constant.setPlaceholderText("SEQ 규칙 설정 사용")
                if column["name"] == "IMAGE_SEQ":
                    source.setCurrentIndex(0)
                    source.setEnabled(False)
                    data_type.setEnabled(False)
                    length.setEnabled(False)
                    unit.setEnabled(False)
            controls = (source, constant, data_type, length, unit, required)
            for col, control in enumerate(controls, 1):
                self.mapping.setCellWidget(row, col, control)
            self.controls.append(controls)
            extension = QComboBox()
            extension.setEditable(True)
            extension.addItems(["", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".heic", ".mp4", ".mov", ".avi", ".mp3", ".wav", ".m4a", ".flac", ".ogg", ".3gp"])
            extension.setCurrentText(column.get("filename_extension") or "")
            extension.lineEdit().setPlaceholderText("추가 안 함")
            extension.setToolTip("DB 번호를 파일명으로 매핑할 때 확장자를 지정하세요. 비워두면 그대로 사용하며, 기존 확장자가 있으면 유지합니다.")
            extension.setEnabled(column["name"] in FILENAME_COLUMNS)
            self.mapping.setCellWidget(row, 7, extension)
            self.extension_controls[column["name"]] = extension
            lookup = QComboBox()
            lookup.addItem("엑셀 / 고정값", False)
            lookup.addItem("폴더 ↔ DB 번호", True)
            lookup.setCurrentIndex(1 if column.get("folder_lookup") else 0)
            lookup.setEnabled(column["name"] in FILENAME_COLUMNS)
            if column["name"] in FILENAME_COLUMNS:
                source.setEnabled(not lookup.currentData())
                constant.setEnabled(not lookup.currentData())
                lookup.currentIndexChanged.connect(
                    lambda _, combo=lookup, source=source, constant=constant:
                    (source.setEnabled(not combo.currentData()), constant.setEnabled(not combo.currentData()))
                )
            lookup.setToolTip("폴더 모드는 엑셀 DB 번호와 확장자를 제외한 파일명을 대조합니다. 후보가 여러 개면 오류로 표시하며, 확장자 옵션은 후보를 좁히는 데 사용합니다.")
            self.mapping.setCellWidget(row, 8, lookup)
            self.lookup_controls[column["name"]] = lookup
        mapping_layout.addWidget(self.mapping)
        tabs.addTab(mapping_page, "컬럼 매핑·타입·길이")
        settings_page = QWidget()
        form = QFormLayout(settings_page)
        self.nas = QLineEdit(profile["nas_template"])
        self.nas.setPlaceholderText("실제 NAS 폴더 경로 입력")
        form.addRow("NAS 경로 템플릿", self.nas)
        self.organization_column = choices(headers, profile.get("organization_code_column"))
        self.organization_code = QLineEdit(profile.get("organization_code", ""))
        self.organization_code.setPlaceholderText("기관코드 컬럼을 매핑하지 않을 때 사용")
        form.addRow("기관코드 컬럼", self.organization_column)
        form.addRow("고정 기관코드", self.organization_code)
        if "image" in profile["kind"]:
            self.organization_column.setEnabled(False)
            self.organization_code.clear()
            self.organization_code.setEnabled(False)
            self.organization_code.setPlaceholderText("엑셀의 기관코드 컬럼 사용")
        hint = QLabel("치환값: {SPECS_ESSNT_ID}, {OBSRV_INFO_SEQ}, {SMPL_INFO_SEQ}, {기관코드}. 이미지는 엑셀의 기관코드 컬럼 값을 사용합니다. 로컬 파일 이동·복사 및 NAS 접속은 수행하지 않습니다.")
        hint.setWordWrap(True)
        form.addRow(hint)
        self.type_column = choices(headers, profile.get("type_column"))
        self.code_column = choices(headers, profile.get("code_column"))
        form.addRow("미디어타입 컬럼", self.type_column)
        form.addRow("미디어코드 컬럼", self.code_column)
        self.target_codes = QLineEdit(", ".join(profile.get("target_codes", [])))
        self.other_codes = QLineEdit(", ".join(profile.get("other_codes", [])))
        form.addRow("이 유형에 해당하는 코드 (쉼표 구분)", self.target_codes)
        form.addRow("다른 유형으로 제외할 코드", self.other_codes)
        hint = QLabel("OB 등의 코드는 임의 해석하지 않습니다. 코드 컬럼을 매핑하면 허용/제외 코드를 입력해야 합니다. 미디어타입만 사용할 경우 코드 컬럼을 ‘매핑 없음’으로 지정하세요.")
        hint.setWordWrap(True)
        form.addRow(hint)
        self.all_rows = QCheckBox("유형 구분 컬럼이 없는 시트이며 전체 행이 선택한 유형임을 확인")
        self.all_rows.setChecked(profile.get("all_rows_confirmed", False))
        self.type_column.currentIndexChanged.connect(self.sync_type_filter)
        self.code_column.currentIndexChanged.connect(self.sync_type_filter)
        self.sync_type_filter()
        form.addRow(self.all_rows)
        self.seq_mode = QComboBox()
        for label, key in (("고정 순번", "fixed"), ("연결 키별 MAX + 1", "max_per_key"), ("엑셀 SEQ 컬럼 사용", "source")):
            self.seq_mode.addItem(label, key)
        self.seq_mode.setCurrentIndex(max(0, self.seq_mode.findData(profile.get("sequence_mode"))))
        self.seq_value = QLineEdit(str(profile.get("sequence_value", "1")))
        self.seq_confirmed = QCheckBox("실제 PK/UNIQUE 구성과 위 SEQ 규칙을 확인했습니다")
        self.seq_confirmed.setChecked(profile.get("sequence_confirmed", False))
        self.seq_mode.currentIndexChanged.connect(lambda: self.seq_confirmed.setChecked(False))
        self.seq_value.textEdited.connect(lambda: self.seq_confirmed.setChecked(False))
        if profile["kind"] in ("observation_image", "specimen_image"):
            self.seq_mode.setEnabled(False)
            self.seq_value.setReadOnly(True)
            self.seq_confirmed.setEnabled(False)
        form.addRow("SEQ 생성 방식", self.seq_mode)
        form.addRow("고정 SEQ 값", self.seq_value)
        form.addRow(self.seq_confirmed)
        self.value_codes = QLineEdit(", ".join(profile.get("value_codes", ["03"]) if "image" in profile["kind"] else profile.get("use_codes", [])))
        form.addRow("이미지구분 / 사용 여부 허용 코드", self.value_codes)
        self.codes_confirmed = QCheckBox("실제 DB 코드값을 확인했습니다 (Y/N 등을 임의 가정하지 않음)")
        self.codes_confirmed.setChecked("image" in profile["kind"] or profile.get("use_codes_confirmed", False))
        self.value_codes.textEdited.connect(lambda: self.codes_confirmed.setChecked(False))
        form.addRow(self.codes_confirmed)
        self.database_number = choices(headers, profile.get("database_number_column"))
        form.addRow("DB 번호 컬럼 (폴더 매칭·동영상 내용)", self.database_number)
        self.vernacular_name = choices(headers, profile.get("vernacular_name_column"))
        self.scientific_name = choices(headers, profile.get("scientific_name_column"))
        form.addRow("빈 동영상 국문 제목·설명 → 국명", self.vernacular_name)
        form.addRow("빈 동영상 영문 제목·설명 → 학명", self.scientific_name)
        self.encoding = QComboBox()
        self.encoding.addItems(["utf-8", "cp949", "euc-kr"])
        self.encoding.setCurrentText(profile.get("encoding", "utf-8"))
        form.addRow("BYTE 길이 검사 시 DB 인코딩", self.encoding)
        self.lock = QCheckBox("실행용 SQL에 테이블 잠금 포함 (동시 INSERT 직렬화, COMMIT은 직접 수행)")
        self.lock.setChecked(profile.get("lock_table", True))
        form.addRow(self.lock)
        settings_scroll = QScrollArea()
        settings_scroll.setWidgetResizable(True)
        settings_scroll.setWidget(settings_page)
        tabs.addTab(settings_scroll, "유형 필터·NAS·SEQ·코드")
        preview_page = QWidget()
        preview_layout = QVBoxLayout(preview_page)
        notice = QLabel("현재 매핑으로 최대 20행을 확인합니다. 설정을 바꾼 뒤 ‘매핑 미리보기 새로고침’을 누르세요. 로컬 폴더는 파일명 검색용이며 NAS 경로와 별도로 설정합니다. DB 길이 확인 전에도 미리볼 수 있습니다. NAS 파일 존재 여부는 검사하지 않습니다.")
        notice.setWordWrap(True)
        preview_layout.addWidget(notice)
        refresh = QPushButton("매핑 미리보기 새로고침")
        refresh.clicked.connect(self.refresh_mapping_preview)
        preview_layout.addWidget(refresh)
        self.mapping_preview = QTableWidget()
        self.mapping_preview.setEditTriggers(QTableWidget.NoEditTriggers)
        preview_layout.addWidget(self.mapping_preview)
        tabs.addTab(preview_page, "파일명·경로 미리보기")
        tabs.currentChanged.connect(lambda index: self.refresh_mapping_preview() if tabs.widget(index) is preview_page else None)
        self.schema_confirmed = QCheckBox("실제 DB 스키마의 컬럼 타입·최대 길이·필수값·연결 키를 확인했습니다")
        self.schema_confirmed.setChecked(profile.get("schema_confirmed", False))
        self.schema_confirmed.setEnabled(False)
        layout.addWidget(self.schema_confirmed)
        hint = QLabel("정적 검증만 수행합니다. DB 연결 키 존재 여부, 기존 DB 중복, 실제 NAS 파일 존재 여부는 확인하지 않습니다.")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("검증 및 SQL 생성")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def set_preview_rows(self, rows):
        self.preview_rows = rows[:20]

    def refresh_mapping_preview(self):
        profile = self.values()
        index, index_error = None, ""
        if any(col.get("folder_lookup") for col in profile["columns"]):
            try:
                index = MediaFileIndex(profile["media_lookup_dir"], profile["lookup_recursive"], profile["lookup_skip_previews"])
            except ValueError as exc:
                index_error = str(exc)
        headers = ["엑셀 행", "연결 키", "DB 번호", "파일명 컬럼 / 가져오는 방식", "적용 파일명", "매칭된 로컬 파일", "SQL NAS 경로", "SQL 저장 경로 + 파일명", "확인 결과"]
        records = []
        for excel_row, row in self.preview_rows:
            status, reason = filter_row(row, profile)
            if status != "SUCCESS":
                records.append([excel_row, "", "", "", "", "", "", "", f"{status}: {reason}"])
                continue
            errors, key, nas = [], "", ""
            try:
                key_col = next(col for col in profile["columns"] if col["name"] == profile["key"])
                value = mapped_value(key_col, row)
                key = str(value).strip() if value is not None and "image" in profile["kind"] else number(value, True) or ""
                if not key:
                    raise ValueError("연결 키가 비어 있습니다.")
            except ValueError as exc:
                errors.append(f"{profile['key']}: {exc}")
            try:
                nas = mapped_nas_path(profile, row, key)
            except ValueError as exc:
                errors.append(str(exc))
            dbno = row.get(profile.get("database_number_column"))
            for col in profile["columns"]:
                if col["name"] not in FILENAME_COLUMNS:
                    continue
                filename, local, problems = "", "", list(errors)
                mode = "폴더 ↔ DB 번호" if col.get("folder_lookup") else (col.get("source") or "고정값")
                try:
                    if col.get("folder_lookup"):
                        if index_error:
                            raise ValueError(index_error)
                        path = index.filepath(dbno, profile["kind"], col)
                        filename, local = path.name, str(path)
                    else:
                        filename = str(mapped_value(col, row) or "")
                    if not filename and col.get("required"):
                        problems.append("필수 파일명이 비어 있습니다.")
                except ValueError as exc:
                    problems.append(str(exc))
                records.append([excel_row, key, dbno, f"{col['name']} / {mode}", filename, local or "로컬 파일 매칭 안 함", nas, nas + filename if nas and filename else "", "; ".join(problems) or "매핑 확인 (DB 검증 전)"])
        table = self.mapping_preview
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(records))
        for r, record in enumerate(records):
            for c, value in enumerate(record):
                table.setItem(r, c, QTableWidgetItem("" if value is None else str(value)))
        table.resizeColumnsToContents()

    def choose_lookup_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "SQL 파일명 매칭 폴더 선택", self.lookup_folder.text())
        if folder:
            self.lookup_folder.setText(folder)

    def sync_type_filter(self, *_):
        missing = not self.type_column.currentData() and not self.code_column.currentData()
        self.all_rows.setChecked(missing)
        self.filter_notice.setText(f"미디어 구분 컬럼이 없어 선택한 유형 ‘{self.profile['label']}’으로 전체 행을 처리합니다. 다른 유형이 섞여 있다면 필터 컬럼을 지정하세요." if missing else "미디어 구분 컬럼으로 행별 유형을 판별합니다.")

    def values(self):
        result = deepcopy(self.profile)
        for column, controls in zip(result["columns"], self.controls):
            source, constant, data_type, length, unit, required = controls
            column.update(source=source.currentData(), constant=constant.text() or None,
                          type=data_type.currentText(), length=length.value() or None,
                          length_unit=unit.currentText(), required=required.isChecked())
            column["filename_extension"] = self.extension_controls[column["name"]].currentText().strip() if column["name"] in FILENAME_COLUMNS else ""
            column["folder_lookup"] = bool(self.lookup_controls[column["name"]].currentData()) if column["name"] in FILENAME_COLUMNS else False
        result.update(nas_template=self.nas.text().strip(),
                      organization_code_column=self.organization_column.currentData(), organization_code=self.organization_code.text().strip(),
                      type_column=self.type_column.currentData(),
                      code_column=self.code_column.currentData(), target_codes=items(self.target_codes.text()),
                      other_codes=items(self.other_codes.text()), all_rows_confirmed=self.all_rows.isChecked(),
                      sequence_mode=self.seq_mode.currentData(), sequence_value=self.seq_value.text().strip(),
                      sequence_confirmed=self.seq_confirmed.isChecked(), value_codes=items(self.value_codes.text()),
                      use_codes=items(self.value_codes.text()), use_codes_confirmed=self.codes_confirmed.isChecked(),
                      database_number_column=self.database_number.currentData(), encoding=self.encoding.currentText(),
                      vernacular_name_column=self.vernacular_name.currentData(), scientific_name_column=self.scientific_name.currentData(),
                      media_lookup_dir=self.lookup_folder.text().strip(), lookup_recursive=self.recursive.isChecked(), lookup_skip_previews=self.skip_previews.isChecked(),
                      lock_table=self.lock.isChecked(), schema_confirmed=self.schema_confirmed.isChecked())
        return result
