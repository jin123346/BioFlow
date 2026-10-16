import logging
import sys

from PySide6.QtCore import QThread, QTimer, QUrl, Signal, Slot
from PySide6.QtGui import QColor, QDesktopServices, QPalette
from PySide6.QtWidgets import QApplication, QMessageBox #PySide6 앱 실행 객체
from config.paths import LOG_DIR, ensure_data_seeded, get_app_version
from services.data_update_service import apply_update, check_for_app_update, check_for_update
from services.data_update_service import check_for_classification_update, apply_classification_update
from ui.main_window import MainWindow #메인 창 클래스
from ui.theme import STYLE_SHEET, install_scroll_wheel_guard

LOG_FILE = LOG_DIR / "app.log"


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
        self.checked.emit(check_for_update(), check_for_app_update(get_app_version()), check_for_classification_update())


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
        reply = QMessageBox.question(window, "분류군 기준자료 업데이트", f"분류군 기준자료 새 버전 {classification['version']}이 있습니다.\n업데이트할까요?")
        if reply == QMessageBox.Yes:
            try:
                apply_classification_update(classification)
                review = getattr(window, "_loaded_pages", {}).get(3)
                if review is not None:
                    review.refresh_reference_label()
                QMessageBox.information(window, "분류군 업데이트 완료", "새 분류군 기준자료를 적용했습니다.")
            except Exception as exc:
                QMessageBox.warning(window, "분류군 업데이트 실패", str(exc))


def install_error_logging():
    # 콘솔 없이(더블클릭 등) 실행하면 에러/타이밍 로그가 화면에 안 보이고 사라지므로 파일로도 남김
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    def log_unhandled(exc_type, exc_value, exc_traceback):
        logging.error("Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))
        sys.__excepthook__(exc_type, exc_value, exc_traceback)

    sys.excepthook = log_unhandled


def main():
    install_error_logging()
    ensure_data_seeded()
    app = QApplication(sys.argv)  #GUI 프로그램 시작 준비
    # Avoid the native white background before the first styled paint.
    palette = app.palette()
    palette.setColor(QPalette.Window, QColor("#f4f6f9"))
    app.setPalette(palette)
    app.setStyleSheet(STYLE_SHEET)
    install_scroll_wheel_guard()

    #창 생성하고 화면에 보여줌
    window = MainWindow()
    window.show()

    # 빌드 시 스플래시를 켰으면(BioFlow.spec) 메인창이 뜬 시점에 닫음. 개발 모드(python
    # app.py)에는 pyi_splash 모듈 자체가 없어서 그냥 넘어감.
    try:
        import pyi_splash
        # Keep the splash until the event loop has painted the main window.
        QTimer.singleShot(100, pyi_splash.close)
    except ImportError:
        pass

    # 창이 뜬 뒤 잠깐 있다가 기준 데이터(종정보 마스터/기관 정보) 새 버전이 있는지 확인
    QTimer.singleShot(1500, lambda: start_update_checks(window))
    app.aboutToQuit.connect(lambda: window.update_check_worker.wait() if hasattr(window, "update_check_worker") else None)

    #이벤트 루프 실행, 버튼 클릭, 창 이동 같은 GUI 동작을 계속 기다림
    sys.exit(app.exec())


def check_for_data_update(window, remote_info=None):
    if not remote_info:
        return
    reply = QMessageBox.question(
        window,
        "업데이트 확인",
        "새로운 기준 데이터(종정보 마스터/기관 정보)가 있습니다.\n지금 받아올까요?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply != QMessageBox.Yes:
        return
    try:
        apply_update(remote_info)
        QMessageBox.information(window, "업데이트 완료", "기준 데이터를 최신 버전으로 업데이트했습니다.\n앱을 다시 시작해 주세요.")
    except Exception as exc:
        QMessageBox.warning(window, "업데이트 실패", f"업데이트 중 문제가 발생했습니다: {exc}")


def check_for_app_release(window, remote_info=None):
    if not remote_info:
        return
    reply = QMessageBox.question(
        window,
        "새 버전 안내",
        f"BioFlow 새 버전 {remote_info.get('version')}이(가) 나왔습니다.\n다운로드 페이지를 열까요?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply == QMessageBox.Yes:
        QDesktopServices.openUrl(QUrl(remote_info["download_url"]))


if __name__=="__main__":
    main()
