from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel


class SqlTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        self.setLayout(layout)

        layout.addWidget(QLabel("SQL 생성 탭"))