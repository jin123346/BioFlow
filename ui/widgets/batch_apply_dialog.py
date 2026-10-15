from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout,
    QLabel, QMessageBox, QSpinBox, QVBoxLayout,
)


class BatchApplyDialog(QDialog):
    def __init__(self, projects, media_codes, row_count, parent=None):
        super().__init__(parent)
        self.setWindowTitle("일괄 적용 설정")
        self.projects = projects
        self.media_codes = media_codes
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"전체 {row_count}행에 적용합니다. 선택한 항목의 기존 값은 덮어씁니다."))
        form = QFormLayout()
        self.species = QCheckBox("국명/학명으로 종정보를 조회하여 적용")
        self.species.setChecked(True)
        self.preserve_numbers = QCheckBox("기존 데이터베이스번호 유지 (빈 번호만 생성)")
        self.preserve_numbers.setChecked(True)
        form.addRow("번호 생성", self.preserve_numbers)
        self.start_number = QSpinBox()
        self.start_number.setRange(1, 99999)
        self.start_number.setValue(1)
        self.start_number.setToolTip("각 사업의 사업별 순번을 이 값부터 매깁니다. DB 번호도 사업별 순번을 사용합니다.")
        form.addRow("사업별 순번 시작 번호", self.start_number)
        form.addRow("종ID·학명", self.species)
        self.taxon = QComboBox()
        self.taxon.addItem("기존 값 유지", "")
        self.taxon.addItems(["IN", "PL", "BI", "PI", "MM", "SP", "MR", "AR", "AM", "SW", "AG"])
        form.addRow("분류코드", self.taxon)
        self.year = QComboBox()
        years = sorted({str(p.get("사업년도", "")) for p in projects if p.get("사업년도")}, reverse=True)
        self.year.addItem("전체 사업년도", "")
        for year in years:
            self.year.addItem(year, year)
        self.project = QComboBox()
        self.project.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.project.setMinimumContentsLength(25)
        self.year.currentIndexChanged.connect(self.refresh_projects)
        form.addRow("사업년도", self.year)
        form.addRow("사업명", self.project)
        self.refresh_projects()
        if media_codes:
            self.media = QComboBox()
            self.media.addItem("기존 값 유지", "")
            for name, code in media_codes.items():
                self.media.addItem(f"{name} ({code})", name)
            form.addRow("미디어타입·코드", self.media)
        layout.addLayout(form)
        layout.addWidget(QLabel("사업년도와 순번은 선택한 사업에 맞춰 적용됩니다."))
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("전체 행에 적용")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.resize(520, 260)

    def refresh_projects(self):
        self.project.clear()
        self.project.addItem("기존 값 유지", None)
        year = self.year.currentData()
        for project in self.projects:
            if not year or str(project.get("사업년도", "")) == year:
                label = f"{project.get('사업년도', '')}-{project.get('순번', '')} {project.get('사업명', '')}"
                self.project.addItem(label, project)

    def values(self):
        return {
            "species": self.species.isChecked(),
            "taxon_code": self.taxon.currentData() or (self.taxon.currentText() if self.taxon.currentIndex() else ""),
            "preserve_database_numbers": self.preserve_numbers.isChecked(),
            "database_start_number": self.start_number.value(),
            "project": self.project.currentData(),
            "media_type": self.media.currentData() if self.media_codes else "",
        }

    def accept(self):
        values = self.values()
        if not (values["species"] or values["project"] or values["media_type"] or values["taxon_code"]
                or not values["preserve_database_numbers"] or values["database_start_number"] != 1):
            QMessageBox.warning(self, "일괄 적용", "적용할 항목을 선택해 주세요.")
            return
        super().accept()
