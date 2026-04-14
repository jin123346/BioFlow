
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QLabel)

class UploadTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        self.setLayout(layout)

        layout.addWidget(QLabel("신규 데이터 업로드 탭"))