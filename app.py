import logging
import sys

from PySide6.QtWidgets import QApplication #PySide6 앱 실행 객체
from config.paths import LOG_DIR, ensure_data_seeded
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

    #이벤트 루프 실행, 버튼 클릭, 창 이동 같은 GUI 동작을 계속 기다림
    sys.exit(app.exec())


if __name__=="__main__":
    main()