"""Credential settings and a nonblocking VWorld connection check."""

from datetime import date

from PySide6.QtCore import QDate, QThread, Signal
from PySide6.QtWidgets import (
    QCheckBox, QDateEdit, QDialog, QFormLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QVBoxLayout,
)

from services.api_key_service import ApiKeyError, ApiKeyService
from services.vworld_geocoding_service import GeocodingError, VWorldGeocodingService


class ConnectionWorker(QThread):
    result = Signal(str)

    def __init__(self, service, key, latitude, longitude, parent=None):
        super().__init__(parent)
        self.service, self.key = service, key
        self.latitude, self.longitude = latitude, longitude

    def run(self):
        try:
            self.service.lookup(self.latitude, self.longitude, self.key)
            self.result.emit("연결 성공: 지번주소 응답을 확인했습니다.")
        except GeocodingError as exc:
            self.result.emit(str(exc))
        except Exception:
            self.result.emit("연결 테스트 중 오류가 발생했습니다.")
        finally:
            self.key = ""


class ApiSettingsDialog(QDialog):
    def __init__(self, parent=None, key_service=None, geocoder=None):
        super().__init__(parent)
        self.keys = key_service or ApiKeyService()
        self.geocoder = geocoder or VWorldGeocodingService()
        self.worker = None
        self.setWindowTitle("VWorld API 설정")
        self.resize(540, 300)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.key_input = QLineEdit()
        self.key_input.setEchoMode(QLineEdit.Password)
        self.key_input.setPlaceholderText("새 키 입력 (비워두면 저장된 키 사용)")
        form.addRow("API 키", self.key_input)
        self.expiry_enabled = QCheckBox("발급 정보의 만료일 등록")
        self.expiry = QDateEdit(QDate.currentDate())
        self.expiry.setCalendarPopup(True)
        self.expiry.setDisplayFormat("yyyy-MM-dd")
        self.expiry.setEnabled(False)
        self.expiry_enabled.toggled.connect(self.expiry.setEnabled)
        form.addRow(self.expiry_enabled, self.expiry)
        self.latitude = QLineEdit("37.5665")
        self.longitude = QLineEdit("126.9780")
        form.addRow("테스트 위도", self.latitude)
        form.addRow("테스트 경도", self.longitude)
        layout.addLayout(form)
        self.status = QLabel()
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        buttons = QHBoxLayout()
        self.save_button = QPushButton("저장 / 변경")
        self.delete_button = QPushButton("키 삭제")
        self.test_button = QPushButton("연결 테스트")
        self.close_button = QPushButton("닫기")
        for button, handler in ((self.save_button, self.save_key), (self.delete_button, self.delete_key),
                                (self.test_button, self.test_connection), (self.close_button, self.reject)):
            buttons.addWidget(button)
            button.clicked.connect(handler)
        layout.addLayout(buttons)
        self.load_status()

    def load_status(self):
        try:
            key, expiry = self.keys.load()
            self.expiry_enabled.setChecked(expiry is not None)
            if expiry:
                self.expiry.setDate(QDate(expiry.year, expiry.month, expiry.day))
            message = "저장된 API 키가 있습니다." if key else "API 키를 등록해 주세요."
            if expiry and expiry < date.today():
                message += " 등록한 만료일이 지났습니다. 실제 발급 정보를 확인해 주세요."
            self.status.setText(message)
        except ApiKeyError as exc:
            self.status.setText(str(exc))

    def save_key(self):
        try:
            key = self.key_input.text().strip() or self.keys.load()[0]
            expiry = self.expiry.date().toPython() if self.expiry_enabled.isChecked() else None
            self.keys.save(key, expiry)
            self.key_input.clear()
            self.load_status()
        except ApiKeyError as exc:
            self.status.setText(str(exc))

    def delete_key(self):
        try:
            self.keys.delete()
            self.key_input.clear()
            self.load_status()
        except ApiKeyError as exc:
            self.status.setText(str(exc))

    def test_connection(self):
        if self.worker is not None and self.worker.isRunning():
            return
        try:
            key = self.key_input.text().strip() or self.keys.load()[0]
        except ApiKeyError as exc:
            self.status.setText(str(exc))
            return
        if not key:
            self.status.setText("API 키를 입력하거나 등록해 주세요.")
            return
        self.worker = ConnectionWorker(self.geocoder, key, self.latitude.text(), self.longitude.text(), self)
        self.worker.result.connect(self.status.setText)
        self.worker.finished.connect(self.test_finished)
        self.set_busy(True)
        self.status.setText("VWorld 연결 확인 중…")
        self.worker.start()

    def set_busy(self, busy):
        for widget in (self.save_button, self.delete_button, self.test_button, self.close_button,
                       self.key_input, self.latitude, self.longitude, self.expiry_enabled):
            widget.setEnabled(not busy)
        self.expiry.setEnabled(not busy and self.expiry_enabled.isChecked())

    def test_finished(self):
        self.set_busy(False)

    def reject(self):
        if self.worker is not None and self.worker.isRunning():
            return
        self.key_input.clear()
        super().reject()

    def closeEvent(self, event):
        if self.worker is not None and self.worker.isRunning():
            event.ignore()
        else:
            self.key_input.clear()
            super().closeEvent(event)
