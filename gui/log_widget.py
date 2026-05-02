"""色付きログ表示ウィジェット."""

from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QTextCharFormat, QTextCursor
from PyQt6.QtWidgets import QPlainTextEdit


_LEVEL_COLORS = {
    "DEBUG":   QColor("#888888"),
    "INFO":    QColor("#dddddd"),
    "WARNING": QColor("#f0b400"),
    "ERROR":   QColor("#ff5252"),
    "CRITICAL": QColor("#ff5252"),
}


class LogWidget(QPlainTextEdit):
    """シグナル経由でメインスレッドから安全にログを追加するためのウィジェット."""

    MAX_BLOCKS = 5000

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setReadOnly(True)
        self.setMaximumBlockCount(self.MAX_BLOCKS)
        self.setStyleSheet(
            """
            QPlainTextEdit {
                background-color: #1e1e1e;
                color: #dddddd;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 12px;
                border: 1px solid #444;
            }
            """
        )

    def append_log(self, level: str, message: str) -> None:
        color = _LEVEL_COLORS.get(level, _LEVEL_COLORS["INFO"])
        fmt = QTextCharFormat()
        fmt.setForeground(color)

        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(f"[{level}] {message}\n", fmt)
        self.setTextCursor(cursor)
        self.ensureCursorVisible()

    def clear_log(self) -> None:
        self.clear()
