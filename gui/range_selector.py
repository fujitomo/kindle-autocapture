"""画面全体に半透明オーバーレイを出してROIをドラッグ選択するウィジェット.

戻り値はスクリーン座標 (left, top, width, height)。
複数モニタ環境でも仮想スクリーン全体をカバーする。
"""

from __future__ import annotations

from typing import Optional, Tuple

from PyQt6.QtCore import QRect, QPoint, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QPainter, QPen, QGuiApplication, QKeyEvent, QMouseEvent
from PyQt6.QtWidgets import QWidget


class RangeSelector(QWidget):
    """フルスクリーン半透明オーバーレイ + ドラッグ選択."""

    selectionMade = pyqtSignal(int, int, int, int)  # x, y, w, h
    cancelled = pyqtSignal()

    def __init__(self) -> None:
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setCursor(Qt.CursorShape.CrossCursor)

        # 仮想スクリーン全体を覆う
        virtual = QGuiApplication.primaryScreen().virtualGeometry()
        self.setGeometry(virtual)

        self._origin: Optional[QPoint] = None
        self._current: Optional[QPoint] = None
        self._dragging: bool = False

    def show_and_select(self) -> None:
        self.showFullScreen()
        self.raise_()
        self.activateWindow()

    # ---------------- events ----------------
    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Escape:
            self.cancelled.emit()
            self.close()
        else:
            super().keyPressEvent(event)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._origin = event.globalPosition().toPoint()
            self._current = self._origin
            self._dragging = True
            self.update()

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if self._dragging:
            self._current = event.globalPosition().toPoint()
            self.update()

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self._dragging:
            self._dragging = False
            self._current = event.globalPosition().toPoint()
            rect = self._normalized_rect()
            if rect.width() < 10 or rect.height() < 10:
                self.cancelled.emit()
                self.close()
                return
            self.selectionMade.emit(rect.x(), rect.y(), rect.width(), rect.height())
            self.close()

    # ---------------- painting ----------------
    def paintEvent(self, _event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # 全面を半透明黒で塗る
        painter.fillRect(self.rect(), QColor(0, 0, 0, 100))

        if self._origin and self._current:
            local_rect = self._normalized_rect_local()

            # 選択範囲は透明にくり抜く
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_Clear)
            painter.fillRect(local_rect, Qt.GlobalColor.transparent)
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)

            # 枠線
            pen = QPen(QColor(0, 200, 255), 2)
            painter.setPen(pen)
            painter.drawRect(local_rect)

            # サイズ表示
            painter.setPen(QColor(255, 255, 255))
            text = f"{local_rect.width()} x {local_rect.height()}"
            painter.drawText(local_rect.topLeft() + QPoint(4, -6), text)

        # 操作ヒント
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(
            20, 30,
            "ドラッグで範囲を選択 / ESC でキャンセル"
        )

    def _normalized_rect(self) -> QRect:
        if not self._origin or not self._current:
            return QRect()
        return QRect(self._origin, self._current).normalized()

    def _normalized_rect_local(self) -> QRect:
        """ウィジェット座標系での選択範囲（描画用）."""
        global_rect = self._normalized_rect()
        top_left = global_rect.topLeft() - self.geometry().topLeft()
        return QRect(top_left, global_rect.size())
