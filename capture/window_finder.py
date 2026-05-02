"""Kindle for PC ウィンドウ検出.

タイトルでフィルタしつつ、トップレベルウィンドウを列挙する。
クライアント領域 / ウィンドウ枠 の座標も返す。
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import List, Optional

try:
    import win32gui
    import win32process
    import win32con
except ImportError:  # 開発環境（非Windows）対応
    win32gui = None
    win32process = None
    win32con = None

from utils.logger import get_logger

log = get_logger("window_finder")


@dataclass
class WindowInfo:
    hwnd: int
    title: str
    class_name: str
    pid: int
    rect: tuple  # (left, top, right, bottom) ウィンドウ枠込み
    client_rect: tuple  # (left, top, right, bottom) スクリーン座標系のクライアント領域

    @property
    def width(self) -> int:
        return self.rect[2] - self.rect[0]

    @property
    def height(self) -> int:
        return self.rect[3] - self.rect[1]

    @property
    def client_width(self) -> int:
        return self.client_rect[2] - self.client_rect[0]

    @property
    def client_height(self) -> int:
        return self.client_rect[3] - self.client_rect[1]

    def is_likely_minimized_or_offscreen(self) -> bool:
        """最小化・オフスクリーン時に Win32 が返すダミー座標を検出する。"""
        # 最小化ウィンドウは (-32000, -32000) 付近に配置されることが多い
        if self.rect[0] < -10000 or self.rect[1] < -10000:
            return True
        if self.client_rect[0] < -10000 or self.client_rect[1] < -10000:
            return True
        if self.client_width <= 0 or self.client_height <= 0:
            return True
        if self.width <= 0 or self.height <= 0:
            return True
        return False

    def __str__(self) -> str:
        return f"{self.title} [hwnd={self.hwnd}, {self.width}x{self.height}]"


class WindowFinder:
    """Win32 API を介したウィンドウ列挙・取得."""

    def __init__(self) -> None:
        if win32gui is None:
            log.warning("pywin32 が利用できません。ウィンドウ検出機能は無効化されます。")

    def list_top_level_windows(
        self,
        title_filter: Optional[str] = None,
        require_visible: bool = True,
    ) -> List[WindowInfo]:
        """トップレベルウィンドウを列挙する。

        title_filter: 部分一致（大文字小文字区別なし）。None なら全件。
        require_visible: True なら IsWindowVisible のみ。
        """
        if win32gui is None:
            return []

        results: List[WindowInfo] = []
        needle = title_filter.lower() if title_filter else None

        def _enum_proc(hwnd: int, _) -> bool:
            try:
                if require_visible and not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                if not title:
                    return True
                if needle and needle not in title.lower():
                    return True

                class_name = win32gui.GetClassName(hwnd) or ""
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                client_rect = self._get_client_screen_rect(hwnd)
                results.append(
                    WindowInfo(
                        hwnd=hwnd,
                        title=title,
                        class_name=class_name,
                        pid=pid,
                        rect=rect,
                        client_rect=client_rect,
                    )
                )
            except Exception as e:
                log.debug("ウィンドウ列挙中の例外 hwnd=%s: %s", hwnd, e)
            return True

        win32gui.EnumWindows(_enum_proc, None)
        return results

    def find_kindle_windows(self) -> List[WindowInfo]:
        """タイトルに 'Kindle' を含むウィンドウを返す。"""
        return self.list_top_level_windows(title_filter="Kindle")

    def get_window_info(self, hwnd: int) -> Optional[WindowInfo]:
        """hwnd から最新のウィンドウ情報を取得（座標が変わっている可能性あり）。"""
        if win32gui is None:
            return None
        try:
            if not win32gui.IsWindow(hwnd):
                return None
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd) or ""
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            client_rect = self._get_client_screen_rect(hwnd)
            return WindowInfo(
                hwnd=hwnd,
                title=title,
                class_name=class_name,
                pid=pid,
                rect=rect,
                client_rect=client_rect,
            )
        except Exception as e:
            log.warning("ウィンドウ情報取得失敗 hwnd=%s: %s", hwnd, e)
            return None

    def bring_to_foreground(self, hwnd: int) -> bool:
        """ウィンドウを前面に持ち出す（最小化されていれば復元）。

        AttachThreadInput で自プロセスの入力スレッドと Kindle を一時結合し、
        SetForegroundWindow の成功率を上げる（Qt から別プロセスを前面にする定石）。
        """
        if win32gui is None:
            return False
        try:
            if not win32gui.IsWindow(hwnd):
                return False
            try:
                # 最小化: IsIconic / GetWindowPlacement の両方で判定
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                placement = win32gui.GetWindowPlacement(hwnd)
                if placement and placement[1] == win32con.SW_SHOWMINIMIZED:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            except Exception:
                pass
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            except Exception:
                pass

            fg_hwnd = None
            fg_tid = None
            target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
            try:
                fg_hwnd = win32gui.GetForegroundWindow()
                if fg_hwnd:
                    fg_tid = win32process.GetWindowThreadProcessId(fg_hwnd)[0]
            except Exception:
                pass

            attached = False
            if fg_tid and target_tid and fg_tid != target_tid:
                try:
                    win32process.AttachThreadInput(fg_tid, target_tid, True)
                    attached = True
                except Exception as e:
                    log.debug("AttachThreadInput 失敗（続行）: %s", e)

            try:
                try:
                    import ctypes

                    ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)
                except Exception:
                    pass
                win32gui.SetForegroundWindow(hwnd)
                return True
            except Exception as e:
                log.debug("SetForegroundWindow 失敗、ALTトグル後再試行: %s", e)
                try:
                    import ctypes

                    try:
                        ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)
                    except Exception:
                        pass
                    ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
                    ctypes.windll.user32.keybd_event(0x12, 0, 0x0002, 0)
                    win32gui.SetForegroundWindow(hwnd)
                    return True
                except Exception as e2:
                    log.warning("ウィンドウ前面化失敗 hwnd=%s: %s", hwnd, e2)
                    return False
            finally:
                if attached and fg_tid is not None:
                    try:
                        win32process.AttachThreadInput(fg_tid, target_tid, False)
                    except Exception:
                        pass
        except Exception as e:
            log.warning("bring_to_foreground 失敗: %s", e)
            return False

    def is_foreground(self, hwnd: int) -> bool:
        """前面ウィンドウが Kindle 本体か、その子（読書ペイン）か。"""
        if win32gui is None:
            return False
        try:
            fg = win32gui.GetForegroundWindow()
            if not fg:
                return False
            if fg == hwnd:
                return True
            root = win32gui.GetAncestor(fg, win32con.GA_ROOT)
            return root == hwnd
        except Exception:
            return False

    def find_postmessage_key_target(self, root_hwnd: int) -> int:
        """PostMessage 先として有望な子 HWND を返す（Chromium/Qt の描画ウィンドウ）。

        トップレベルへ送っても Kindle が無視するため、スコア最大の子に送る。
        """
        if win32gui is None:
            return root_hwnd
        best_hwnd = root_hwnd
        best_score = 0

        def score_window(h: int) -> int:
            try:
                if not win32gui.IsWindowVisible(h):
                    return 0
                cn = (win32gui.GetClassName(h) or "").lower()
                r = win32gui.GetWindowRect(h)
                area = max(0, (r[2] - r[0]) * (r[3] - r[1]))
                s = 0
                if "chrome_widget" in cn or "cef" in cn or "webview" in cn:
                    s = 2_000_000 + area
                elif "intermediate d3d" in cn:
                    s = 1_800_000 + area
                elif "qt" in cn and "qwindow" in cn.replace("_", ""):
                    s = 900_000 + area
                elif area > 80_000:
                    s = area // 100
                return s
            except Exception:
                return 0

        def visit(h: int, depth: int) -> None:
            nonlocal best_hwnd, best_score
            if depth > 8:
                return
            sc = score_window(h)
            if sc > best_score:
                best_score = sc
                best_hwnd = h
            try:

                def _cb(ch: int, _) -> bool:
                    visit(ch, depth + 1)
                    return True

                win32gui.EnumChildWindows(h, _cb, None)
            except Exception:
                pass

        try:
            visit(root_hwnd, 0)
        except Exception:
            pass
        if best_hwnd != root_hwnd:
            log.debug("PostMessage 先を子 HWND に変更: %s (score=%s)", best_hwnd, best_score)
        return best_hwnd

    def _get_client_screen_rect(self, hwnd: int) -> tuple:
        """クライアント領域をスクリーン座標で返す。"""
        try:
            cl, ct, cr, cb = win32gui.GetClientRect(hwnd)
            sl, st = win32gui.ClientToScreen(hwnd, (cl, ct))
            sr, sb = win32gui.ClientToScreen(hwnd, (cr, cb))
            return (sl, st, sr, sb)
        except Exception:
            return win32gui.GetWindowRect(hwnd)

    def prepare_for_capture(self, hwnd: int, settle_sec: float = 0.35) -> Optional[WindowInfo]:
        """最小化解除・前面化のあと座標が安定するまで待って WindowInfo を返す。

        キャプチャ領域計算の**前**に必ず呼ぶこと。呼ばないと最小化時に (-32000,0) 等になる。
        """
        if win32gui is None:
            return None
        if not win32gui.IsWindow(hwnd):
            return None

        self.bring_to_foreground(hwnd)
        time.sleep(settle_sec)
        info = self.get_window_info(hwnd)
        if info is None:
            return None
        if info.is_likely_minimized_or_offscreen():
            # レイアウトが遅い環境向けに1回だけ再試行
            self.bring_to_foreground(hwnd)
            time.sleep(max(settle_sec, 0.5))
            info = self.get_window_info(hwnd)
        return info
