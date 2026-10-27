from PySide6.QtWidgets import QDialog, QDialogButtonBox, QLabel, QRadioButton, QVBoxLayout


class SaveOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("저장 방식 선택")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("저장할 내용을 선택하세요."))
        self.excel_only = QRadioButton("Excel만 저장 (기본·완성본·GBIF, 미디어 변경 없음)")
        self.with_media = QRadioButton("Excel과 미디어 함께 저장")
        self.excel_only.setChecked(True)
        layout.addWidget(self.excel_only)
        layout.addWidget(self.with_media)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @property
    def include_media(self):
        return self.with_media.isChecked()
