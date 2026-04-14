
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QLabel)

class GbifTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        self.setLayout(layout)

        layout.addWidget(QLabel("GBIF 데이터 탭"))