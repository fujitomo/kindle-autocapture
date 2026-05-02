"""Capture orchestration controller (QObject worker).

QThread にmoveToThreadして使う。シグナルでGUIに状態を通知する。
"""

from __future__ import annotations

import time
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import QObject, pyqtSignal, pyqtSlot

from capture.duplicate_detector import DuplicateDetector
from capture.screenshotter import CaptureRegion, ScreenCapturer
from capture.window_finder import WindowFinder
from config.config_manager import AppConfig, ConfigManager
from navigation.page_navigator import PageNavigator
from storage.image_saver import ImageSaver
from storage.pdf_assembler import build_session_pdf
from utils.logger import get_logger

from .states import CaptureState, StopReason

log = get_logger("controller")

# 旧 config.json で min_captures=2 のままだと「2枚で最終ページ」誤爆が再発するため下限をかける
_MIN_CAPTURES_BEFORE_LAST_PAGE_FLOOR = 5


@dataclass
class CaptureRequest:
    hwnd: int
    output_dir: Path
    max_pages: int = 0  # 0 = 制限なし
    region: Optional[CaptureRegion] = None  # None ならウィンドウ全体


class CaptureController(QObject):
    """キャプチャの実行制御本体."""

    # ---------------- signals ----------------
    stateChanged = pyqtSignal(object)  # CaptureState
    progressUpdated = pyqtSignal(int, int)  # captured, max(0は不明)
    pageCaptured = pyqtSignal(int, str, bool)  # index, file_path, was_duplicate
    logMessage = pyqtSignal(str, str)  # level, message
    finished = pyqtSignal(object, str)  # StopReason, summary
    errorOccurred = pyqtSignal(str)
    sessionStarted = pyqtSignal(str)  # session_dir

    def __init__(self) -> None:
        super().__init__()
        self._state: CaptureState = CaptureState.IDLE
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._pause_event.set()  # set = 動作可（clear = 一時停止）
        self._captured_count: int = 0
        self._start_time: Optional[float] = None

        self._capturer = ScreenCapturer()
        self._navigator = PageNavigator()
        self._duplicate = DuplicateDetector()
        self._finder = WindowFinder()
        self._saver: Optional[ImageSaver] = None

    # ---------------- public API ----------------
    @property
    def state(self) -> CaptureState:
        return self._state

    @property
    def captured_count(self) -> int:
        return self._captured_count

    @pyqtSlot(object)
    def start(self, request: CaptureRequest) -> None:
        """キャプチャ開始。worker thread から呼ばれる前提（GUIからは invokeMethod 経由）。"""
        if self._state in (CaptureState.RUNNING, CaptureState.STOPPING):
            self._emit_log("WARNING", "既に実行中のため start を無視")
            return
        try:
            self._run_loop(request)
        except Exception as e:
            log.exception("予期せぬ例外: %s", e)
            self.errorOccurred.emit(str(e))
            self._set_state(CaptureState.IDLE)
            self.finished.emit(StopReason.ERROR, f"異常終了: {e}")

    @pyqtSlot()
    def pause(self) -> None:
        if self._state == CaptureState.RUNNING:
            self._pause_event.clear()
            self._set_state(CaptureState.PAUSED)
            self._emit_log("INFO", "一時停止しました")

    @pyqtSlot()
    def resume(self) -> None:
        if self._state == CaptureState.PAUSED:
            self._pause_event.set()
            self._set_state(CaptureState.RUNNING)
            self._emit_log("INFO", "再開しました")

    @pyqtSlot()
    def stop(self) -> None:
        if self._state in (CaptureState.RUNNING, CaptureState.PAUSED):
            self._stop_event.set()
            self._pause_event.set()  # 一時停止中の場合に解除
            self._set_state(CaptureState.STOPPING)
            self._emit_log("INFO", "停止要求を受け付けました")

    # ---------------- core loop ----------------
    def _run_loop(self, request: CaptureRequest) -> None:
        cfg: AppConfig = ConfigManager().config
        session_dir: Optional[Path] = None

        # 古い config.json に "right" / "left" / "pageup" 系の逆方向キーが残っていると
        # ページが逆送りになるため、安全なキーかどうかを検証してから使う
        nav_key = cfg.navigation.key.lower().strip()
        _FORWARD_KEYS = {"right", "pagedown", "page_down", "down", "space"}
        if nav_key not in _FORWARD_KEYS:
            self._emit_log(
                "WARNING",
                f"ページ送りキー '{nav_key}' は次ページ方向ではない可能性があります。"
                " 設定で 'pagedown' または 'right' に変更してください。",
            )
        self._navigator.configure(
            method=cfg.navigation.method,
            key=nav_key,
            key_delivery=cfg.navigation.key_delivery,
            wheel_scroll_notches=max(
                1, min(10, int(getattr(cfg.navigation, "wheel_scroll_notches", 1)))
            ),
            wheel_scroll_sign=int(getattr(cfg.navigation, "wheel_scroll_sign", -1)),
            invert_page_turn=bool(getattr(cfg.navigation, "invert_page_turn", True)),
            click_x_offset=cfg.navigation.click_x_offset,
            click_y_offset=cfg.navigation.click_y_offset,
            refocus_before_action=cfg.navigation.refocus_before_action,
        )
        self._duplicate = DuplicateDetector(
            duplicate_threshold=cfg.capture.duplicate_threshold,
            last_page_max_distance=cfg.capture.last_page_max_hash_distance,
        )

        # 最小化中は GetClientRect が (-32000,…) になり width/height=0 になる。
        # 必ず復元・前面化してから座標を取る（従来はここが逆順だった）。
        settle = max(0.25, float(cfg.capture.pre_capture_delay_sec))
        info = self._finder.prepare_for_capture(request.hwnd, settle_sec=settle)
        if info is None:
            self.errorOccurred.emit("Kindle ウィンドウが見つかりません")
            self.finished.emit(StopReason.ERROR, "ウィンドウ未検出")
            return
        if info.is_likely_minimized_or_offscreen():
            msg = (
                "ウィンドウの座標が無効です（多くの場合、Kindle が最小化されています）。"
                "タスクバーからウィンドウを復元してから、もう一度「開始」してください。"
            )
            self.errorOccurred.emit(msg)
            self._emit_log("ERROR", msg)
            self.finished.emit(StopReason.ERROR, "ウィンドウ未復元")
            return

        if request.region is not None:
            region = request.region
        else:
            region = self._build_region_from_window(info)
        if not region.is_valid:
            self.errorOccurred.emit(f"無効なキャプチャ領域: {region}")
            self.finished.emit(StopReason.ERROR, "無効な領域")
            return

        self._saver = ImageSaver(
            output_root=request.output_dir,
            book_subfolder_template=cfg.storage.book_subfolder_template,
            file_template=cfg.storage.file_template,
            save_format=cfg.capture.save_format,
            jpeg_quality=cfg.capture.jpeg_quality,
            webp_quality=cfg.capture.webp_quality,
            resume_from_existing=cfg.storage.resume_from_existing,
        )
        session_dir = self._saver.start_session()
        self.sessionStarted.emit(str(session_dir))

        self._captured_count = 0
        self._start_time = time.time()
        self._stop_event.clear()
        self._pause_event.set()
        self._set_state(CaptureState.RUNNING)

        self._emit_log("INFO", f"キャプチャ開始: {info.title}")
        self._emit_log("INFO", f"範囲: {region.width}x{region.height} @ ({region.left},{region.top})")
        self._emit_log("INFO", f"保存先: {session_dir}")

        stop_reason: StopReason = StopReason.USER
        consecutive_capture_errors = 0
        duplicate_skips_without_save = 0
        min_cap_last = max(
            _MIN_CAPTURES_BEFORE_LAST_PAGE_FLOOR,
            int(cfg.capture.min_captures_before_last_page_stop),
        )

        try:
            while True:
                if self._stop_event.is_set():
                    stop_reason = StopReason.USER
                    break

                # pause 中はブロック
                if not self._pause_event.is_set():
                    self._pause_event.wait()
                    if self._stop_event.is_set():
                        stop_reason = StopReason.USER
                        break

                # ページ更新があるかもしれないので最新情報で region を取り直す
                # （ROI がウィンドウ依存なら更新、そうでなければ固定）
                if request.region is None:
                    info = self._finder.get_window_info(request.hwnd) or info
                    if info.is_likely_minimized_or_offscreen():
                        self._emit_log("WARNING", "ウィンドウが最小化された可能性。復元を試みます")
                        info = self._finder.prepare_for_capture(request.hwnd, settle_sec=0.35) or info
                    region = self._build_region_from_window(info)
                    if not region.is_valid:
                        self._emit_log("ERROR", "キャプチャ領域が無効になりました（最小化など）。停止します")
                        stop_reason = StopReason.ERROR
                        self.errorOccurred.emit("キャプチャ中にウィンドウ座標が無効化されました")
                        break

                # キャプチャ
                try:
                    img = self._capturer.capture(region)
                except Exception as e:
                    consecutive_capture_errors += 1
                    self._emit_log("ERROR", f"キャプチャ失敗 ({consecutive_capture_errors}/3): {e}")
                    if consecutive_capture_errors >= 3:
                        stop_reason = StopReason.ERROR
                        self.errorOccurred.emit(f"キャプチャ連続失敗: {e}")
                        break
                    time.sleep(0.5)
                    continue
                consecutive_capture_errors = 0

                if ScreenCapturer.is_blank(img):
                    self._emit_log("WARNING", "ブラックスクリーンの可能性。短時間待機して再試行")
                    time.sleep(0.5)
                    continue

                # 重複判定
                dup = self._duplicate.check(img)
                if dup.is_duplicate:
                    self._emit_log(
                        "DEBUG",
                        "重複検出 "
                        f"(distance={dup.distance}, 厳密連続={dup.last_page_streak}/"
                        f"{cfg.capture.last_page_max_hash_distance}以下で加算)",
                    )
                    if (
                        cfg.capture.auto_stop_on_last_page
                        and self._captured_count >= min_cap_last
                        and self._duplicate.is_likely_last_page(
                            cfg.capture.last_page_consecutive_dupes
                        )
                    ):
                        self._emit_log("INFO", "最終ページを検知しました")
                        stop_reason = StopReason.LAST_PAGE
                        break
                    if cfg.capture.skip_duplicates:
                        duplicate_skips_without_save += 1
                        max_skips = max(1, int(cfg.capture.max_duplicate_skips_without_save))
                        if duplicate_skips_without_save >= max_skips:
                            msg = (
                                f"保存なしで重複スキップが {max_skips} 回続きました。"
                                "右キーが Kindle に届いていない可能性があります。"
                                "設定の「キー送り」を postmessage / pyautogui で切り替えるか、"
                                "Kindle を一度前面にしてから再実行してください。"
                            )
                            self._emit_log("ERROR", msg)
                            self.errorOccurred.emit(msg)
                            stop_reason = StopReason.STUCK_DUPLICATES
                            break
                        self._send_next_page(request.hwnd, cfg.capture.page_delay_sec)
                        continue

                # 保存
                try:
                    result = self._saver.save(img)
                except Exception as e:
                    self._emit_log("ERROR", f"保存失敗: {e}")
                    self.errorOccurred.emit(str(e))
                    stop_reason = StopReason.ERROR
                    break

                self._captured_count += 1
                duplicate_skips_without_save = 0
                self.pageCaptured.emit(result.index, str(result.path), bool(dup.is_duplicate))
                self.progressUpdated.emit(self._captured_count, request.max_pages)

                # ディスク空き確認（100枚ごと）
                if self._captured_count % 100 == 0:
                    free = ImageSaver.free_disk_bytes(session_dir)
                    if 0 < free < 200 * 1024 * 1024:  # 200MB 未満
                        self._emit_log("ERROR", "ディスク空き容量不足")
                        stop_reason = StopReason.DISK_FULL
                        break

                # 上限
                if request.max_pages > 0 and self._captured_count >= request.max_pages:
                    self._emit_log("INFO", "指定枚数に到達しました")
                    stop_reason = StopReason.MAX_PAGES
                    break

                # ページ送り直前にフォーカス保険：Kindle が前面でなければ前面化
                if not self._finder.is_foreground(request.hwnd):
                    self._finder.bring_to_foreground(request.hwnd)
                    time.sleep(0.12)

                # ページ送り
                ok = self._send_next_page(request.hwnd, cfg.capture.page_delay_sec)
                if not ok:
                    self._emit_log("WARNING", "ページ送り失敗。リトライします")
                    time.sleep(0.5)
                    ok = self._send_next_page(request.hwnd, cfg.capture.page_delay_sec)
                    if not ok:
                        self._emit_log("ERROR", "ページ送り再試行も失敗")
                        stop_reason = StopReason.ERROR
                        break

        finally:
            cfg_end = ConfigManager().config
            pdf_line = ""
            if (
                session_dir is not None
                and self._captured_count > 0
                and bool(getattr(cfg_end.storage, "auto_pdf", True))
            ):
                pdf_name = str(getattr(cfg_end.storage, "pdf_filename", None) or "book.pdf").strip() or "book.pdf"
                pdf_path = session_dir / pdf_name
                try:
                    pages = build_session_pdf(session_dir, pdf_path)
                    if pages > 0:
                        pdf_line = f"\nPDF: {pdf_path}"
                        self._emit_log("INFO", f"PDF を保存しました ({pages} ページ): {pdf_path}")
                except Exception as e:
                    log.exception("PDF 生成失敗")
                    self._emit_log("ERROR", f"PDF 生成に失敗しました: {e}")

            elapsed = time.time() - (self._start_time or time.time())
            summary = (
                f"取得 {self._captured_count} 枚 / 経過 {elapsed:.1f}秒 / 理由: {stop_reason.label}"
                f"{pdf_line}"
            )
            self._emit_log("INFO", summary)
            self._set_state(CaptureState.IDLE)
            self.finished.emit(stop_reason, summary)
            self._capturer.close()

    # ---------------- helpers ----------------
    def _send_next_page(self, hwnd: int, delay: float) -> bool:
        result = self._navigator.next_page(hwnd)
        if not result.success:
            self._emit_log("WARNING", f"ページ送り失敗: {result.error}")
            return False
        # 描画完了待ち（GUI 応答性のため細切れに sleep）
        end = time.time() + delay
        while time.time() < end:
            if self._stop_event.is_set():
                return True
            if not self._pause_event.is_set():
                self._pause_event.wait()
                if self._stop_event.is_set():
                    return True
            time.sleep(0.05)
        return True

    def _build_region_from_window(self, info) -> CaptureRegion:
        cfg = ConfigManager().config.roi
        if cfg.use_window or cfg.width <= 0 or cfg.height <= 0:
            l, t, r, b = info.client_rect
            return CaptureRegion(left=l, top=t, width=r - l, height=b - t)
        return CaptureRegion(left=cfg.x, top=cfg.y, width=cfg.width, height=cfg.height)

    def _set_state(self, state: CaptureState) -> None:
        if self._state != state:
            self._state = state
            self.stateChanged.emit(state)

    def _emit_log(self, level: str, message: str) -> None:
        log_method = {
            "DEBUG": log.debug,
            "INFO": log.info,
            "WARNING": log.warning,
            "ERROR": log.error,
        }.get(level, log.info)
        log_method(message)
        self.logMessage.emit(level, message)
