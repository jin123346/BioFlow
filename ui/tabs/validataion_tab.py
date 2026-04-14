from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel


class ValidationTab(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        self.setLayout(layout)

        layout.addWidget(QLabel("데이터검수 탭"))