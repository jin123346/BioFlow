"""Modal workflow keeps editing stable while network work runs off the UI thread."""

from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QDialog, QDialogButtonBox, QLabel,
    QProgressBar, QPushButton, QRadioButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout,
)

from services.address_mapping_service import apply_address, map_address
from services.vworld_geocoding_service import GeocodingError
from services.english_address_service import AddressDatabaseError, EnglishAddressService, english_fields


class AddressScopeDialog(QDialog):
    """Choose which table rows should be sent to VWorld."""

    def __init__(self, selected_count=0, parent=None):
        super().__init__(parent)
        self.setWindowTitle("주소 조회 범위")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("조회할 행을 선택해 주세요. 기존 주소는 결과 확인 전까지 변경되지 않습니다."))
        self.empty = QRadioButton("주소가 비어 있는 행만 조회")
        self.selected = QRadioButton(f"현재 선택한 행 조회 ({selected_count}행)")
        self.all_rows = QRadioButton("좌표가 있는 전체 행 다시 조회")
        self.empty.setChecked(True)
        self.selected.setEnabled(selected_count > 0)
        for button in (self.empty, self.selected, self.all_rows):
            layout.addWidget(button)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def mode(self):
        if self.selected.isChecked():
            return "selected"
        if self.all_rows.isChecked():
            return "all"
        return "empty"


class AddressWorker(QThread):
    progress = Signal(int)

    def __init__(self, service, key, targets, parent=None):
        super().__init__(parent)
        self.service, self.key, self.targets = service, key, targets
        self.results, self.errors = {}, {}
        self.references, self.reference_errors = {}, {}

    def run(self):
        try:
            for index, (identifier, record) in enumerate(self.targets):
                if self.isInterruptionRequested():
                    break
                try:
                    address = self.service.lookup(record.get("위도"), record.get("경도"), self.key)
                    if not self.isInterruptionRequested():
                        self.results[identifier] = address
                        try:
                            reference = EnglishAddressService().lookup(address.structure.get("level4LC", ""))
                            if reference:
                                self.references[identifier] = reference
                            else:
                                self.reference_errors[identifier] = "법정동코드가 영문 DB에 없습니다. 기존 영문주소를 유지합니다."
                        except AddressDatabaseError as exc:
                            self.reference_errors[identifier] = str(exc)
                except GeocodingError as exc:
                    self.errors[identifier] = str(exc)
                    if exc.code in {"authentication", "rate_limit", "missing_key"}:
                        break
                except Exception:
                    self.errors[identifier] = "주소 조회 중 오류가 발생했습니다."
                self.progress.emit(index + 1)
                # Sequential calls; cancellation remains responsive during pacing.
                if index + 1 < len(self.targets):
                    for _ in range(10):
                        if self.isInterruptionRequested():
                            break
                        self.msleep(100)
        finally:
            self.key = ""


class AddressLookupDialog(QDialog):
    def __init__(self, service, key, targets, parent=None):
        super().__init__(parent)
        self.setWindowTitle("VWorld 국문주소 조회")
        layout = QVBoxLayout(self)
        self.label = QLabel("주소를 조회하고 있습니다…")
        layout.addWidget(self.label)
        self.bar = QProgressBar()
        self.bar.setRange(0, len(targets))
        layout.addWidget(self.bar)
        self.cancel_button = QPushButton("취소")
        layout.addWidget(self.cancel_button)
        self.cancel_button.clicked.connect(self.reject)
        self.cancelled = False
        self.worker = AddressWorker(service, key, targets, self)
        self.worker.progress.connect(self.bar.setValue)
        self.worker.finished.connect(self.completed)
        QTimer.singleShot(0, self.start_lookup)

    def start_lookup(self):
        if self.cancelled:
            super().reject()
        else:
            self.worker.start()

    def reject(self):
        self.cancelled = True
        self.worker.requestInterruption()
        self.cancel_button.setEnabled(False)
        self.label.setText("취소 중… 현재 요청이 끝나면 닫힙니다. 주소는 적용하지 않습니다.")
        if self.worker.isFinished():
            super().reject()

    def completed(self):
        if self.cancelled:
            super().reject()
        else:
            self.accept()

    def closeEvent(self, event):
        if not self.worker.isFinished():
            self.reject()
            event.ignore()
        else:
            super().closeEvent(event)


class AddressPreviewDialog(QDialog):
    def __init__(self, targets, results, errors, specimen=False, parent=None, references=None,
                 reference_errors=None, differences_only=False):
        super().__init__(parent)
        self.targets, self.results, self.errors = targets, results, errors
        self.specimen = specimen
        self.references = references or {}
        self.reference_errors = reference_errors or {}
        self.differences_only = differences_only
        self._editable_cells = []
        self.setWindowTitle("VWorld 주소 결과 미리보기")
        self.resize(960, 530)
        layout = QVBoxLayout(self)
        if differences_only:
            guide = "기존값과 다른 항목만 표시합니다. 적용 예정값은 더블클릭해 수정할 수 있습니다."
        else:
            guide = "기본값은 빈칸만 채우기입니다. 법정동코드가 일치하는 영문주소만 함께 적용합니다."
        layout.addWidget(QLabel(guide))
        self.overwrite = QCheckBox("기존 국문·영문주소 덮어쓰기 (조회되지 않은 필드는 기존 값 유지)")
        layout.addWidget(self.overwrite)
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["행", "항목", "기존 값", "조회 결과", "적용 예정"])
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        layout.addWidget(self.table)
        buttons = QDialogButtonBox(QDialogButtonBox.Apply | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Apply).clicked.connect(self.accept)
        buttons.button(QDialogButtonBox.Apply).setEnabled(bool(results))
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.overwrite.toggled.connect(self.refresh)
        self.refresh()

    def refresh(self):
        self.table.setRowCount(0)
        self._editable_cells = []
        for identifier, record in self.targets:
            address = self.results.get(identifier)
            if address is None:
                rows = [("조회 상태", "", self.errors.get(identifier, "요청 중단으로 조회하지 않음"), "변경 없음")]
            else:
                fields, _ = map_address(address, self.specimen)
                reference = self.references.get(identifier)
                updated = apply_address(record, address, self.specimen, self.overwrite.isChecked(), reference)
                rows = [(field, record.get(field, ""), value, updated.get(field, "")) for field, value in fields.items()]
                if reference:
                    rows += [(field, record.get(field, ""), value, updated.get(field, ""))
                             for field, value in english_fields(address, reference, self.specimen).items()]
                elif identifier in self.reference_errors:
                    rows += [("영문주소", "", self.reference_errors[identifier], "기존 값 유지")]
                rows += [("법정동코드 (내부)", "", address.structure.get("level4LC", ""), "내부 보관"),
                         ("상세주소 보조값 (내부)", "", address.structure.get("detail", ""), "내부 보관")]
                if self.differences_only:
                    public_fields = set(fields)
                    if reference:
                        public_fields.update(english_fields(address, reference, self.specimen))
                    rows = [row for row in rows if row[0] in public_fields]
                    comparable = lambda value: "".join(str(value or "").split()).casefold()
                    rows = [row for row in rows if comparable(row[1]) and
                            comparable(row[1]) != comparable(row[2])]
            for row in rows:
                index = self.table.rowCount()
                self.table.insertRow(index)
                for col, value in enumerate((str(identifier + 1), *row)):
                    item = QTableWidgetItem(str(value or ""))
                    if col != 4 or row[0] in {"조회 상태", "영문주소", "법정동코드 (내부)", "상세주소 보조값 (내부)"}:
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    else:
                        item.setData(Qt.UserRole, (identifier, row[0]))
                        self._editable_cells.append(item)
                    self.table.setItem(index, col, item)
        self.table.resizeColumnsToContents()

    def planned_records(self):
        """Return mapped records plus any values edited in the preview."""
        updates = {}
        target_map = dict(self.targets)
        for identifier, address in self.results.items():
            if identifier not in target_map:
                continue
            updates[identifier] = apply_address(
                target_map[identifier], address, self.specimen, self.overwrite.isChecked(),
                self.references.get(identifier),
            )
        for item in self._editable_cells:
            identifier, field = item.data(Qt.UserRole)
            if identifier not in updates:
                continue
            value = item.text().strip()
            original = str(target_map[identifier].get(field, "") or "").strip()
            if value == original:
                continue
            updates[identifier][field] = value
            metadata = updates[identifier].get("_vworld_address", {})
            write_fields = list(metadata.get("write_fields", metadata.get("applied_fields", [])))
            if field not in write_fields:
                write_fields.append(field)
            metadata["write_fields"] = write_fields
        return updates
