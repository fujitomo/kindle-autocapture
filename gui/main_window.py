"""メインウィンドウ.

責務:
- ウィンドウ選択・保存先指定・ROI 指定 などの GUI 入力
- CaptureController を別スレッドに置いてシグナル/スロットで連携
- グローバルホットキー（keyboard ライブラリ）と Qt シグナルの橋渡し
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import List, Optional

from PyQt6.QtCore import (
    QObject,
    Qt,
    QThread,
    QTimer,
    pyqtSignal,
    pyqtSlot,
)
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QStatusBar,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from app.controller import CaptureController, CaptureRequest
from app.states import CaptureState, StopReason
from capture.screenshotter import CaptureRegion, ScreenCapturer
from capture.window_finder import WindowFinder, WindowInfo
from config.config_manager import ConfigManager
from gui.log_widget import LogWidget
from gui.range_selector import RangeSelector
from gui.settings_dialog import SettingsDialog
from utils.logger import get_logger

log = get_logger("gui")

APP_TITLE = "Kindle AutoCapture"


class GlobalHotkeys(QObject):
    """`keyboard` ライブラリ → Qt シグナル のブリッジ."""

    startStopPressed = pyqtSignal()
    pausePressed = pyqtSignal()
    emergencyPressed = pyqtSignal()

    def __init__(self) -> None:
        super().__init__()
        self._registered_handles: List[int] = []
        self._kb = None

    def install(self, start_stop: str, pause: str, emergency: str) -> bool:
        try:
            import keyboard  # type: ignore
        except Exception as e:
            log.warning("keyboard モジュールが利用できないためグローバルホットキー無効: %s", e)
            return False

        self._kb = keyboard
        try:
            self.uninstall()
            self._registered_handles.append(
                keyboard.add_hotkey(start_stop, lambda: self.startStopPressed.emit())
            )
            self._registered_handles.append(
                keyboard.add_hotkey(pause, lambda: self.pausePressed.emit())
            )
            self._registered_handles.append(
                keyboard.add_hotkey(emergency, lambda: self.emergencyPressed.emit())
            )
            log.info("グローバルホットキー登録: start=%s pause=%s emergency=%s",
                     start_stop, pause, emergency)
            return True
        except Exception as e:
            log.warning("ホットキー登録失敗（管理者権限が必要な場合あり）: %s", e)
            return False

    def uninstall(self) -> None:
        if not self._kb:
            return
        for h in self._registered_handles:
            try:
                self._kb.remove_hotkey(h)
            except Exception:
                pass
        self._registered_handles.clear()


class MainWindow(QMainWindow):
    requestStart = pyqtSignal(object)  # CaptureRequest

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(900, 720)

        self._cfg_mgr = ConfigManager()
        self._finder = WindowFinder()
        self._windows_cache: List[WindowInfo] = []
        self._selected_roi: Optional[CaptureRegion] = None  # None ならウィンドウ全体
        self._range_selector: Optional[RangeSelector] = None
        self._elapsed_seconds: int = 0
        self._state: CaptureState = CaptureState.IDLE

        self._build_ui()
        self._init_worker()
        self._init_hotkeys()
        self._init_timer()

        self.refresh_window_list()

    # ---------------- UI ----------------
    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(10)

        # --- ターゲット
        target_box = QGroupBox("ターゲットウィンドウ")
        tlay = QHBoxLayout(target_box)
        self.cmb_window = QComboBox()
        self.cmb_window.setMinimumWidth(500)
        self.btn_refresh = QPushButton("再検索")
        self.btn_refresh.clicked.connect(self.refresh_window_list)
        tlay.addWidget(self.cmb_window, 1)
        tlay.addWidget(self.btn_refresh)
        root.addWidget(target_box)

        # --- 保存先
        out_box = QGroupBox("保存先フォルダ")
        olay = QHBoxLayout(out_box)
        self.edit_output = QLineEdit(self._cfg_mgr.config.storage.output_dir)
        self.btn_browse = QPushButton("参照…")
        self.btn_browse.clicked.connect(self._on_browse_output)
        self.btn_open_folder = QPushButton("フォルダを開く")
        self.btn_open_folder.clicked.connect(self._on_open_folder)
        olay.addWidget(self.edit_output, 1)
        olay.addWidget(self.btn_browse)
        olay.addWidget(self.btn_open_folder)
        root.addWidget(out_box)

        # --- キャプチャ設定 (簡易)
        cap_box = QGroupBox("キャプチャ設定（詳細は「設定」から）")
        clay = QHBoxLayout(cap_box)

        self.spin_delay = QDoubleSpinBox()
        self.spin_delay.setRange(0.1, 10.0)
        self.spin_delay.setSingleStep(0.1)
        self.spin_delay.setSuffix(" 秒")
        self.spin_delay.setValue(self._cfg_mgr.config.capture.page_delay_sec)
        self.spin_delay.valueChanged.connect(
            lambda v: self._cfg_mgr.update(capture={"page_delay_sec": v})
        )
        clay.addWidget(QLabel("ページ送り遅延:"))
        clay.addWidget(self.spin_delay)
        clay.addSpacing(20)

        self.spin_max = QSpinBox()
        self.spin_max.setRange(0, 100000)
        self.spin_max.setSpecialValueText("無制限")
        self.spin_max.setValue(0)
        clay.addWidget(QLabel("最大ページ数:"))
        clay.addWidget(self.spin_max)
        clay.addSpacing(20)

        self.chk_skip_dup = QCheckBox("重複スキップ")
        self.chk_skip_dup.setChecked(self._cfg_mgr.config.capture.skip_duplicates)
        self.chk_skip_dup.toggled.connect(
            lambda v: self._cfg_mgr.update(capture={"skip_duplicates": v})
        )
        clay.addWidget(self.chk_skip_dup)

        self.chk_auto_stop = QCheckBox("最終ページで自動停止")
        self.chk_auto_stop.setChecked(self._cfg_mgr.config.capture.auto_stop_on_last_page)
        self.chk_auto_stop.toggled.connect(
            lambda v: self._cfg_mgr.update(capture={"auto_stop_on_last_page": v})
        )
        clay.addWidget(self.chk_auto_stop)

        self.chk_minimize = QCheckBox("実行中はアプリを最小化（推奨）")
        self.chk_minimize.setToolTip(
            "ページ送りが Kindle に届かない最大の原因はフォーカス取り合いです。\n"
            "オンにすると、開始時に自動でこのアプリを最小化し、Kindle にフォーカスを譲ります。\n"
            "停止は F9（開始/停止）/ Esc（緊急停止）でも可能です。"
        )
        self.chk_minimize.setChecked(
            getattr(self._cfg_mgr.config.capture, "minimize_during_capture", True)
        )
        self.chk_minimize.toggled.connect(
            lambda v: self._cfg_mgr.update(capture={"minimize_during_capture": v})
        )
        clay.addWidget(self.chk_minimize)

        self.chk_invert_page = QCheckBox("ページ送り方向を反転")
        self.chk_invert_page.setToolTip(
            "進みと逆にめくれるときオンにすると、Page↔PageUp・左右矢印・ホイール符号を入れ替えます。"
        )
        self.chk_invert_page.setChecked(
            bool(getattr(self._cfg_mgr.config.navigation, "invert_page_turn", True))
        )
        self.chk_invert_page.toggled.connect(
            lambda v: self._cfg_mgr.update(navigation={"invert_page_turn": v})
        )
        clay.addWidget(self.chk_invert_page)

        clay.addStretch(1)
        root.addWidget(cap_box)

        # --- ROI
        roi_box = QGroupBox("キャプチャ範囲")
        rlay = QHBoxLayout(roi_box)
        self.lbl_roi = QLabel("ウィンドウ全体（クライアント領域）")
        self.btn_select_roi = QPushButton("範囲を選択…")
        self.btn_select_roi.clicked.connect(self._on_select_roi)
        self.btn_clear_roi = QPushButton("ウィンドウ全体に戻す")
        self.btn_clear_roi.clicked.connect(self._on_clear_roi)
        rlay.addWidget(self.lbl_roi, 1)
        rlay.addWidget(self.btn_select_roi)
        rlay.addWidget(self.btn_clear_roi)
        root.addWidget(roi_box)

        # --- 実行状態
        st_box = QGroupBox("実行状態")
        slay = QVBoxLayout(st_box)
        info_row = QHBoxLayout()
        self.lbl_state = QLabel("状態: 停止中")
        self.lbl_state.setStyleSheet("font-weight: bold;")
        self.lbl_count = QLabel("取得: 0 枚")
        self.lbl_elapsed = QLabel("経過: 00:00:00")
        info_row.addWidget(self.lbl_state)
        info_row.addStretch(1)
        info_row.addWidget(self.lbl_count)
        info_row.addSpacing(20)
        info_row.addWidget(self.lbl_elapsed)
        slay.addLayout(info_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # busy
        self.progress.setVisible(False)
        slay.addWidget(self.progress)

        ctrl_row = QHBoxLayout()
        self.btn_start = QPushButton("▶ 開始 (F9)")
        self.btn_pause = QPushButton("⏸ 一時停止 (F10)")
        self.btn_stop = QPushButton("⏹ 停止 (Esc)")
        self.btn_start.setMinimumHeight(36)
        self.btn_pause.setMinimumHeight(36)
        self.btn_stop.setMinimumHeight(36)
        self.btn_start.clicked.connect(self._on_start)
        self.btn_pause.clicked.connect(self._on_pause_resume)
        self.btn_stop.clicked.connect(self._on_stop)
        self.btn_pause.setEnabled(False)
        self.btn_stop.setEnabled(False)
        ctrl_row.addWidget(self.btn_start)
        ctrl_row.addWidget(self.btn_pause)
        ctrl_row.addWidget(self.btn_stop)
        slay.addLayout(ctrl_row)
        root.addWidget(st_box)

        # --- ログ
        log_box = QGroupBox("ログ")
        llay = QVBoxLayout(log_box)
        self.log_widget = LogWidget()
        llay.addWidget(self.log_widget)
        log_buttons = QHBoxLayout()
        btn_clear_log = QPushButton("クリア")
        btn_clear_log.clicked.connect(self.log_widget.clear_log)
        log_buttons.addStretch(1)
        log_buttons.addWidget(btn_clear_log)
        llay.addLayout(log_buttons)
        root.addWidget(log_box, 1)

        # --- ツールバー & ステータスバー
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        act_settings = QAction("設定", self)
        act_settings.triggered.connect(self._on_open_settings)
        toolbar.addAction(act_settings)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("準備完了")

    # ---------------- worker ----------------
    def _init_worker(self) -> None:
        self._thread = QThread(self)
        self._controller = CaptureController()
        self._controller.moveToThread(self._thread)

        self._controller.stateChanged.connect(self._on_state_changed)
        self._controller.progressUpdated.connect(self._on_progress)
        self._controller.pageCaptured.connect(self._on_page_captured)
        self._controller.logMessage.connect(self.log_widget.append_log)
        self._controller.finished.connect(self._on_finished)
        self._controller.errorOccurred.connect(self._on_error)
        self._controller.sessionStarted.connect(self._on_session_started)

        self.requestStart.connect(self._controller.start)

        self._thread.start()

    def _init_hotkeys(self) -> None:
        cfg = self._cfg_mgr.config.hotkeys
        self._hotkeys = GlobalHotkeys()
        self._hotkeys.startStopPressed.connect(self._on_hotkey_start_stop)
        self._hotkeys.pausePressed.connect(self._on_hotkey_pause)
        self._hotkeys.emergencyPressed.connect(self._on_hotkey_emergency)
        ok = self._hotkeys.install(cfg.start_stop, cfg.pause, cfg.emergency_stop)
        if not ok:
            self.log_widget.append_log(
                "WARNING",
                "グローバルホットキーが無効です（管理者権限で再起動すると有効化される場合があります）",
            )

    def _init_timer(self) -> None:
        self._elapsed_timer = QTimer(self)
        self._elapsed_timer.setInterval(1000)
        self._elapsed_timer.timeout.connect(self._tick_elapsed)

    # ---------------- handlers ----------------
    def refresh_window_list(self) -> None:
        self._windows_cache = self._finder.find_kindle_windows()
        self.cmb_window.clear()
        if not self._windows_cache:
            self.cmb_window.addItem("(Kindle ウィンドウが見つかりません)", None)
            return
        last_hwnd = self._cfg_mgr.config.window.last_selected_hwnd
        select_idx = 0
        for i, w in enumerate(self._windows_cache):
            label = f"{w.title}  [{w.width}x{w.height}]"
            self.cmb_window.addItem(label, w.hwnd)
            if last_hwnd and w.hwnd == last_hwnd:
                select_idx = i
        self.cmb_window.setCurrentIndex(select_idx)

    def _selected_window(self) -> Optional[WindowInfo]:
        hwnd = self.cmb_window.currentData()
        if hwnd is None:
            return None
        return self._finder.get_window_info(int(hwnd))

    def _on_browse_output(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "保存先フォルダを選択", self.edit_output.text())
        if d:
            self.edit_output.setText(d)
            self._cfg_mgr.update(storage={"output_dir": d})

    def _on_open_folder(self) -> None:
        path = self.edit_output.text().strip()
        if not path or not Path(path).exists():
            QMessageBox.warning(self, APP_TITLE, "保存先フォルダが存在しません")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            QMessageBox.warning(self, APP_TITLE, f"フォルダを開けませんでした: {e}")

    def _on_select_roi(self) -> None:
        self._range_selector = RangeSelector()
        self._range_selector.selectionMade.connect(self._on_roi_selected)
        self._range_selector.cancelled.connect(self._on_roi_cancelled)
        self.hide()
        QTimer.singleShot(200, self._range_selector.show_and_select)

    @pyqtSlot(int, int, int, int)
    def _on_roi_selected(self, x: int, y: int, w: int, h: int) -> None:
        self.show()
        self._selected_roi = CaptureRegion(left=x, top=y, width=w, height=h)
        self.lbl_roi.setText(f"範囲: ({x}, {y})  {w} x {h}")
        self._cfg_mgr.update(roi={"use_window": False, "x": x, "y": y, "width": w, "height": h})

    @pyqtSlot()
    def _on_roi_cancelled(self) -> None:
        self.show()

    def _on_clear_roi(self) -> None:
        self._selected_roi = None
        self.lbl_roi.setText("ウィンドウ全体（クライアント領域）")
        self._cfg_mgr.update(roi={"use_window": True, "x": 0, "y": 0, "width": 0, "height": 0})

    def _on_open_settings(self) -> None:
        dlg = SettingsDialog(self)
        if dlg.exec():
            self.spin_delay.setValue(self._cfg_mgr.config.capture.page_delay_sec)
            self.chk_skip_dup.setChecked(self._cfg_mgr.config.capture.skip_duplicates)
            self.chk_auto_stop.setChecked(self._cfg_mgr.config.capture.auto_stop_on_last_page)
            self.chk_minimize.setChecked(
                bool(getattr(self._cfg_mgr.config.capture, "minimize_during_capture", True))
            )
            self.chk_invert_page.setChecked(
                bool(getattr(self._cfg_mgr.config.navigation, "invert_page_turn", True))
            )
            cfg = self._cfg_mgr.config.hotkeys
            self._hotkeys.install(cfg.start_stop, cfg.pause, cfg.emergency_stop)

    def _on_start(self) -> None:
        if self._state in (CaptureState.RUNNING, CaptureState.PAUSED):
            return
        info = self._selected_window()
        if info is None:
            QMessageBox.warning(self, APP_TITLE, "Kindle ウィンドウを選択してください")
            return
        out_dir = Path(self.edit_output.text().strip())
        if not out_dir.exists():
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, APP_TITLE, f"保存先を作成できません: {e}")
                return

        self._cfg_mgr.update(
            window={
                "last_selected_hwnd": info.hwnd,
                "last_selected_title": info.title,
            },
            storage={"output_dir": str(out_dir)},
        )

        request = CaptureRequest(
            hwnd=info.hwnd,
            output_dir=out_dir,
            max_pages=self.spin_max.value(),
            region=self._selected_roi,
        )

        self._elapsed_seconds = 0
        self._update_elapsed_label()
        self._elapsed_timer.start()

        # フォーカスを Kindle に確実に譲るため、開始直前にアプリを最小化
        if getattr(self._cfg_mgr.config.capture, "minimize_during_capture", True):
            try:
                self._finder.bring_to_foreground(info.hwnd)
            except Exception:
                pass
            self.showMinimized()
            QTimer.singleShot(
                250, lambda r=request: (self._finder.bring_to_foreground(r.hwnd), self.requestStart.emit(r))
            )
        else:
            self.requestStart.emit(request)

    def _on_pause_resume(self) -> None:
        # pause/resume/stop は threading.Event のみ操作するためメインスレッドから直接呼んで OK
        # （ワーカースレッドが _run_loop でブロックしているため QueuedConnection は届かない）
        if self._state == CaptureState.RUNNING:
            self._controller.pause()
        elif self._state == CaptureState.PAUSED:
            self._controller.resume()

    def _on_stop(self) -> None:
        if self._state in (CaptureState.RUNNING, CaptureState.PAUSED):
            self._controller.stop()

    # ---- hotkey slots
    def _on_hotkey_start_stop(self) -> None:
        if self._state in (CaptureState.RUNNING, CaptureState.PAUSED):
            self._on_stop()
        else:
            self._on_start()

    def _on_hotkey_pause(self) -> None:
        self._on_pause_resume()

    def _on_hotkey_emergency(self) -> None:
        self._on_stop()

    # ---- controller signals
    @pyqtSlot(object)
    def _on_state_changed(self, state: CaptureState) -> None:
        self._state = state
        self.lbl_state.setText(f"状態: {state.label}")
        running = state == CaptureState.RUNNING
        paused = state == CaptureState.PAUSED
        self.btn_start.setEnabled(state == CaptureState.IDLE)
        self.btn_pause.setEnabled(running or paused)
        self.btn_pause.setText("▶ 再開 (F10)" if paused else "⏸ 一時停止 (F10)")
        self.btn_stop.setEnabled(running or paused)
        self.statusBar().showMessage(state.label)

        if state == CaptureState.IDLE:
            self._elapsed_timer.stop()
            self.progress.setVisible(False)
        else:
            self.progress.setVisible(True)

    @pyqtSlot(int, int)
    def _on_progress(self, captured: int, total: int) -> None:
        if total > 0:
            self.progress.setRange(0, total)
            self.progress.setValue(captured)
            self.lbl_count.setText(f"取得: {captured} / {total} 枚")
        else:
            self.progress.setRange(0, 0)
            self.lbl_count.setText(f"取得: {captured} 枚")

    @pyqtSlot(int, str, bool)
    def _on_page_captured(self, idx: int, path: str, was_dup: bool) -> None:
        # 既に logMessage 経由でログされているのでここはステータスのみ
        self.statusBar().showMessage(f"保存: {Path(path).name}")

    @pyqtSlot(object, str)
    def _on_finished(self, reason: StopReason, summary: str) -> None:
        # 最小化解除して結果を見せる
        if self.isMinimized():
            self.showNormal()
            self.raise_()
            self.activateWindow()
        QMessageBox.information(self, APP_TITLE, f"{reason.label}\n\n{summary}")

    @pyqtSlot(str)
    def _on_error(self, msg: str) -> None:
        self.log_widget.append_log("ERROR", msg)

    @pyqtSlot(str)
    def _on_session_started(self, session_dir: str) -> None:
        self.statusBar().showMessage(f"セッション開始: {session_dir}")

    # ---- timer
    def _tick_elapsed(self) -> None:
        if self._state in (CaptureState.RUNNING, CaptureState.PAUSED):
            self._elapsed_seconds += 1
            self._update_elapsed_label()

    def _update_elapsed_label(self) -> None:
        h, rem = divmod(self._elapsed_seconds, 3600)
        m, s = divmod(rem, 60)
        self.lbl_elapsed.setText(f"経過: {h:02d}:{m:02d}:{s:02d}")

    # ---- close
    def closeEvent(self, event) -> None:
        try:
            if self._state in (CaptureState.RUNNING, CaptureState.PAUSED):
                ans = QMessageBox.question(
                    self,
                    APP_TITLE,
                    "実行中です。停止して終了しますか？",
                )
                if ans != QMessageBox.StandardButton.Yes:
                    event.ignore()
                    return
                self._on_stop()
                # 少し待つ
                self._thread.requestInterruption()

            self._hotkeys.uninstall()
            self._thread.quit()
            self._thread.wait(2000)
        finally:
            super().closeEvent(event)
