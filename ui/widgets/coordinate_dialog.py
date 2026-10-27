from PySide6.QtWidgets import (
    QCheckBox, QDialog, QDialogButtonBox, QFormLayout, QLabel, QLineEdit, QRadioButton, QVBoxLayout,
)

from services.coordinate_service import parse_coordinates


class CoordinateDialog(QDialog):
    def __init__(self, selected_count: int, total_count: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle("위도·경도 일괄 입력")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.latitude = QLineEdit()
        self.longitude = QLineEdit()
        form.addRow("위도 (-90~90)", self.latitude)
        form.addRow("경도 (-180~180)", self.longitude)
        layout.addLayout(form)
        self.selected = QRadioButton(f"선택 행 ({selected_count}개)")
        self.all_rows = QRadioButton(f"전체 행 ({total_count}개)")
        self.selected.setChecked(selected_count > 0)
        self.all_rows.setChecked(selected_count == 0)
        self.selected.setEnabled(selected_count > 0)
        layout.addWidget(self.selected)
        layout.addWidget(self.all_rows)
        self.only_empty = QCheckBox("빈 좌표만 채우기")
        self.only_empty.setChecked(True)
        layout.addWidget(self.only_empty)
        self.error = QLabel("")
        layout.addWidget(self.error)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._accept_valid)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _accept_valid(self):
        try:
            self.coordinates = parse_coordinates(self.latitude.text(), self.longitude.text())
        except ValueError as exc:
            self.error.setText(str(exc))
            return
        self.accept()
