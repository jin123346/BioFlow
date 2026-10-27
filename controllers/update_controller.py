"""Background update checks and their UI presentation."""

from PySide6.QtCore import QThread, QUrl, Signal, Slot
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QMessageBox

from config.paths import get_app_version
from services.data_update_service import (
    apply_classification_update,
    apply_update,
    check_for_app_update,
    check_for_classification_update,
    check_for_update,
)


class UpdateCheckWorker(QThread):
    checked = Signal(object, object, object)

    def __init__(self, window):
        super().__init__(window)
        self.window = window
        self.checked.connect(self.show_results)

    @Slot(object, object, object)
    def show_results(self, data, release, classification):
        if self.window.isVisible():
            show_update_results(self.window, data, release, classification)

    def run(self):
        self.checked.emit(
            check_for_update(),
            check_for_app_update(get_app_version()),
            check_for_classification_update(),
        )


def start_update_checks(window):
    worker = UpdateCheckWorker(window)
    window.update_check_worker = worker
    worker.start()


def show_update_results(window, data, release, classification=None):
    if data:
        check_for_data_update(window, data)
    if release:
        check_for_app_release(window, release)
    if classification:
        reply = QMessageBox.question(
            window,
            "분류군 기준자료 업데이트",
            f"분류군 기준자료 새 버전 {classification['version']}이 있습니다.\n업데이트할까요?",
        )
        if reply == QMessageBox.Yes:
            try:
                apply_classification_update(classification)
                review = getattr(window, "_loaded_pages", {}).get(3)
                if review is not None:
                    review.refresh_reference_label()
                QMessageBox.information(window, "분류군 업데이트 완료", "새 분류군 기준자료를 적용했습니다.")
            except Exception as exc:
                QMessageBox.warning(window, "분류군 업데이트 실패", str(exc))


def check_for_data_update(window, remote_info=None):
    if not remote_info:
        return
    reply = QMessageBox.question(
        window,
        "업데이트 확인",
        "새로운 기준 데이터(종정보 마스터 등)가 있습니다.\n지금 받아올까요?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply != QMessageBox.Yes:
        return
    try:
        apply_update(remote_info)
        QMessageBox.information(
            window,
            "업데이트 완료",
            "기준 데이터를 최신 버전으로 업데이트했습니다.\n앱을 다시 시작해 주세요.",
        )
    except Exception as exc:
        QMessageBox.warning(window, "업데이트 실패", f"업데이트 중 문제가 발생했습니다: {exc}")


def check_for_app_release(window, remote_info=None):
    if not remote_info:
        return
    reply = QMessageBox.question(
        window,
        "새 버전 안내",
        f"BioFlow 새 버전 {remote_info.get('version')}이(가) 출시됐습니다.\n다운로드 페이지를 열까요?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply == QMessageBox.Yes:
        QDesktopServices.openUrl(QUrl(remote_info["download_url"]))
