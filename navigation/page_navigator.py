"""Page navigation: キーボード or マウスクリックでページ送り."""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Optional

import pyautogui

from capture.window_finder import WindowFinder, WindowInfo
from utils.logger import get_logger

log = get_logger("navigator")

# よく使う仮想キー（PostMessage 用）
try:
    import win32api
    import win32con as _wcon

    _VK_MAP = {
        "right": _wcon.VK_RIGHT,
        "left": _wcon.VK_LEFT,
        "pagedown": _wcon.VK_NEXT,
        "page_down": _wcon.VK_NEXT,
        "pageup": _wcon.VK_PRIOR,
        "page_up": _wcon.VK_PRIOR,
        "space": _wcon.VK_SPACE,
        "down": _wcon.VK_DOWN,
    }
except ImportError:
    win32api = None
    _wcon = None
    _VK_MAP = {}

# pyautogui の安全装置を無効化（左上にカーソル移動で abort されるのを防ぐ）
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0  # 個別の sleep を自前で制御するためゼロにする


@dataclass
class NavigationResult:
    success: bool
    error: Optional[str] = None


_VALID_DELIVERY = frozenset(
    ("wheel_pyautogui", "wheel_postmessage", "sendinput", "postmessage", "pyautogui")
)


class PageNavigator:
    """ページ送りを担当する."""

    def __init__(
        self,
        method: str = "key",
        key: str = "pagedown",
        key_delivery: str = "sendinput",
        wheel_scroll_notches: int = 1,
        wheel_scroll_sign: int = -1,
        invert_page_turn: bool = True,
        click_x_offset: int = 0,
        click_y_offset: int = 0,
        refocus_before_action: bool = True,
    ) -> None:
        self.method = method
        self.key = key
        self.key_delivery = key_delivery
        self.wheel_scroll_notches = max(1, min(10, int(wheel_scroll_notches)))
        self.wheel_scroll_sign = -1 if int(wheel_scroll_sign) < 0 else 1
        self.invert_page_turn = bool(invert_page_turn)
        self.click_x_offset = click_x_offset
        self.click_y_offset = click_y_offset
        self.refocus_before_action = refocus_before_action
        self._finder = WindowFinder()

    def configure(
        self,
        method: Optional[str] = None,
        key: Optional[str] = None,
        key_delivery: Optional[str] = None,
        wheel_scroll_notches: Optional[int] = None,
        wheel_scroll_sign: Optional[int] = None,
        invert_page_turn: Optional[bool] = None,
        click_x_offset: Optional[int] = None,
        click_y_offset: Optional[int] = None,
        refocus_before_action: Optional[bool] = None,
    ) -> None:
        if method is not None:
            self.method = method
        if key is not None:
            self.key = key
        if key_delivery is not None:
            self.key_delivery = key_delivery if key_delivery in _VALID_DELIVERY else "sendinput"
        if wheel_scroll_notches is not None:
            self.wheel_scroll_notches = max(1, min(10, int(wheel_scroll_notches)))
        if wheel_scroll_sign is not None:
            self.wheel_scroll_sign = -1 if int(wheel_scroll_sign) < 0 else 1
        if invert_page_turn is not None:
            self.invert_page_turn = bool(invert_page_turn)
        if click_x_offset is not None:
            self.click_x_offset = click_x_offset
        if click_y_offset is not None:
            self.click_y_offset = click_y_offset
        if refocus_before_action is not None:
            self.refocus_before_action = refocus_before_action

    def _physical_key_name(self) -> str:
        """設定の送りキーを、反転オプションに応じて入れ替える。"""
        k = self.key.lower().strip()
        if not self.invert_page_turn:
            return k
        swap = {
            "pagedown": "pageup",
            "page_down": "pageup",
            "pageup": "pagedown",
            "page_up": "pagedown",
            "right": "left",
            "left": "right",
            "down": "up",
            "up": "down",
        }
        return swap.get(k, k)

    def _effective_wheel_sign(self) -> int:
        s = int(self.wheel_scroll_sign)
        return -s if self.invert_page_turn else s

    def next_page(self, hwnd: int) -> NavigationResult:
        """ページ送り操作を1回実行."""
        info = self._finder.get_window_info(hwnd)
        if info is None:
            return NavigationResult(success=False, error="ウィンドウが見つかりません")

        if self.refocus_before_action:
            ok = self._finder.bring_to_foreground(hwnd)
            if not ok:
                log.debug("フォーカス取得失敗（続行）")
            # ホイールは「その座標の最前面ウィンドウ」に届くため Kindle を前面にする
            time.sleep(
                0.18
                if self.key_delivery in ("wheel_pyautogui", "sendinput")
                else 0.08
            )

        try:
            if self.method == "click":
                self._send_click(info)
            elif self.key_delivery == "wheel_pyautogui":
                self._send_wheel_pyautogui(info)
            elif self.key_delivery == "wheel_postmessage":
                self._send_wheel_postmessage(info.hwnd, info)
            elif self.key_delivery == "sendinput":
                self._send_key_sendinput()
            elif self.key_delivery == "postmessage":
                target = self._finder.find_postmessage_key_target(info.hwnd)
                self._send_key_postmessage(target)
            else:
                self._send_key()
            return NavigationResult(success=True)
        except Exception as e:
            log.exception("ページ送り失敗: %s", e)
            return NavigationResult(success=False, error=str(e))

    def _send_key(self) -> None:
        pk = self._physical_key_name()
        log.debug("キー送信(pyautogui): %s (logical=%s invert=%s)", pk, self.key, self.invert_page_turn)
        pyautogui.press(pk)

    def _send_wheel_pyautogui(self, info: WindowInfo) -> None:
        """本文付近へマウスを移し、ホイールでページ送り。

        重要: ``scroll(3)`` は 3 ノッチ＝最大 3 ページ分一度に動く。Kindle では一気に表紙付近まで戻る原因になる。
        既定は ``notches=1``。進み方向が逆なら設定 ``wheel_scroll_sign`` を反転する。
        """
        cl, ct, cr, cb = info.client_rect
        cw, ch = cr - cl, cb - ct
        cx = cl + max(1, int(cw * 0.62))
        cy = ct + max(1, ch // 2)
        eff = self._effective_wheel_sign()
        clicks = int(eff) * int(self.wheel_scroll_notches)
        log.debug(
            "ホイール送信(pyautogui): (%d,%d) clicks=%s (eff_sign=%s notches=%s invert=%s)",
            cx,
            cy,
            clicks,
            eff,
            self.wheel_scroll_notches,
            self.invert_page_turn,
        )
        pyautogui.moveTo(cx, cy, duration=0.04)
        time.sleep(0.02)
        pyautogui.scroll(clicks)

    def _send_wheel_postmessage(self, hwnd: int, info: WindowInfo) -> None:
        """WM_MOUSEWHEEL を描画子 HWND へ（座標はクライアントの中心付近）。"""
        if win32api is None or _wcon is None:
            self._send_wheel_pyautogui(info)
            return
        target = self._finder.find_postmessage_key_target(hwnd)
        cl, ct, cr, cb = info.client_rect
        cw, ch = cr - cl, cb - ct
        sx = cl + max(1, int(cw * 0.62))
        sy = ct + max(1, ch // 2)
        delta = int(self._effective_wheel_sign()) * 120 * int(self.wheel_scroll_notches)
        wparam = ((delta & 0xFFFF) << 16) & 0xFFFFFFFF
        lparam = ((int(sy) & 0xFFFF) << 16) | (int(sx) & 0xFFFF)
        log.debug("ホイール送信(SendMessage hwnd=%s delta=%s)", target, delta)
        win32api.SendMessage(target, _wcon.WM_MOUSEWHEEL, wparam, lparam)

    def _send_key_sendinput(self) -> None:
        """前面ウィンドウへ keybd_event で VK を送る（Kindle がフォーカスを持っている前提）。"""
        if _wcon is None or not _VK_MAP:
            log.warning("sendinput に pywin32 が必要です。pyautogui にフォールバックします。")
            self._send_key()
            return
        import ctypes

        pk = self._physical_key_name()
        vk = int(_VK_MAP.get(pk, _wcon.VK_NEXT)) & 0xFF
        KEYEVENTF_KEYUP = 0x0002
        KEYEVENTF_EXTENDEDKEY = 0x0001
        # Page Up/Down は拡張キー扱いが必要なことが多い
        ext = KEYEVENTF_EXTENDEDKEY if vk in (_wcon.VK_NEXT, _wcon.VK_PRIOR) else 0
        log.debug("キー送信(sendinput vk=%s ext=%s logical=%s)", vk, ext, self.key)
        ctypes.windll.user32.keybd_event(vk, 0, ext, 0)
        time.sleep(0.05)
        ctypes.windll.user32.keybd_event(vk, 0, ext | KEYEVENTF_KEYUP, 0)

    def _send_key_postmessage(self, hwnd: int) -> None:
        """HWND に WM_KEYDOWN/UP を送る。前面化できないときでも Kindle が反応することがある。"""
        if win32api is None or _wcon is None or not _VK_MAP:
            log.warning("postmessage キー送信は pywin32 が必要です。pyautogui にフォールバックします。")
            self._send_key()
            return
        pk = self._physical_key_name()
        vk = _VK_MAP.get(pk, _wcon.VK_RIGHT)
        log.debug("キー送信(postmessage hwnd=%s vk=%s logical=%s)", hwnd, vk, self.key)
        win32api.PostMessage(hwnd, _wcon.WM_KEYDOWN, vk, 0)
        time.sleep(0.03)
        win32api.PostMessage(hwnd, _wcon.WM_KEYUP, vk, 0)

    def _send_click(self, info: WindowInfo) -> None:
        # 本文右寄り（見開きでも右ページ寄り）＋オフセット
        cl, ct, cr, cb = info.client_rect
        cw, ch = cr - cl, cb - ct
        cx_default = cl + max(1, int(cw * 0.62))
        cy_default = ct + ch // 2
        x = cx_default + self.click_x_offset
        y = cy_default + self.click_y_offset
        log.debug("クリック送信: (%d, %d)", x, y)
        pyautogui.moveTo(x, y, duration=0.04)
        pyautogui.click(x=x, y=y)
