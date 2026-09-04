import logging
import sys

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QMessageBox #PySide6 앱 실행 객체
from config.paths import LOG_DIR, ensure_data_seeded
from services.data_update_service import apply_update, check_for_update
from ui.main_window import MainWindow #메인 창 클래스

LOG_FILE = LOG_DIR / "app.log"


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

    #창 생성하고 화면에 보여줌
    window = MainWindow()
    window.show()

    # 빌드 시 스플래시를 켰으면(BioFlow.spec) 메인창이 뜬 시점에 닫음. 개발 모드(python
    # app.py)에는 pyi_splash 모듈 자체가 없어서 그냥 넘어감.
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass

    # 창이 뜬 뒤 잠깐 있다가 기준 데이터(종정보 마스터/기관 정보) 새 버전이 있는지 확인
    QTimer.singleShot(1500, lambda: check_for_data_update(window))

    #이벤트 루프 실행, 버튼 클릭, 창 이동 같은 GUI 동작을 계속 기다림
    sys.exit(app.exec())


def check_for_data_update(window):
    remote_info = check_for_update()
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


if __name__=="__main__":
    main()