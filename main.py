"""Kindle AutoCapture エントリーポイント."""

from __future__ import annotations

import sys
import traceback

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QMessageBox

from gui.main_window import MainWindow
from utils.logger import get_logger

log = get_logger("main")


def _excepthook(exc_type, exc_value, exc_tb) -> None:
    """全例外をログに記録しダイアログ表示."""
    msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    log.error("UNCAUGHT EXCEPTION:\n%s", msg)
    try:
        QMessageBox.critical(None, "Kindle AutoCapture - 致命的エラー", msg)
    except Exception:
        pass


def main() -> int:
    sys.excepthook = _excepthook

    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("KindleAutoCapture")
    app.setOrganizationName("KindleAutoCapture")

    app.setStyleSheet(
        """
        QGroupBox {
            border: 1px solid #555;
            border-radius: 6px;
            margin-top: 8px;
            padding-top: 8px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 4px;
        }
        QPushButton {
            padding: 4px 12px;
        }
        """
    )

    window = MainWindow()
    window.show()
    log.info("アプリケーション起動")
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
