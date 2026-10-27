"""BioFlow application entry point."""

import logging
import sys

from PySide6.QtCore import QTimer
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from config.paths import LOG_DIR, ensure_data_seeded
from controllers.update_controller import start_update_checks
from ui.main_window import MainWindow
from ui.theme import STYLE_SHEET, install_scroll_wheel_guard

LOG_FILE = LOG_DIR / "app.log"


def install_error_logging():
    """Persist otherwise invisible errors from windowed executable builds."""
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


def create_application():
    """Create and configure the Qt application and its main window."""
    install_error_logging()
    ensure_data_seeded()

    application = QApplication(sys.argv)
    palette = application.palette()
    palette.setColor(QPalette.Window, QColor("#f4f6f9"))
    application.setPalette(palette)
    application.setStyleSheet(STYLE_SHEET)
    install_scroll_wheel_guard()

    window = MainWindow()
    return application, window


def main():
    application, window = create_application()
    window.show()

    try:
        import pyi_splash

        QTimer.singleShot(100, pyi_splash.close)
    except ImportError:
        pass

    QTimer.singleShot(1500, lambda: start_update_checks(window))
    application.aboutToQuit.connect(
        lambda: window.update_check_worker.wait()
        if hasattr(window, "update_check_worker")
        else None
    )
    sys.exit(application.exec())


if __name__ == "__main__":
    main()
