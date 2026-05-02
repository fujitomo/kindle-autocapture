"""Screen capture using mss (GDI BitBlt) or PrintWindow (client-only).

画面矩形の BitBlt は、Kindle の上に他アプリ（IDE など）が重なるとその内容も写る。
PrintWindow はウィンドウ自身の描画バッファを取るため重なりを除外できる（失敗時は mss にフォールバック）。

注意: mss.mss() は使用するスレッドで生成すること。
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Optional, Tuple

from PIL import Image

from utils.logger import get_logger

log = get_logger("screenshotter")


@dataclass
class CaptureRegion:
    left: int
    top: int
    width: int
    height: int

    def to_mss_monitor(self) -> dict:
        return {
            "left": int(self.left),
            "top": int(self.top),
            "width": max(1, int(self.width)),
            "height": max(1, int(self.height)),
        }

    @property
    def is_valid(self) -> bool:
        return self.width > 0 and self.height > 0


def _intersect_screen_rects(
    r: CaptureRegion, client_screen: Tuple[int, int, int, int]
) -> Optional[CaptureRegion]:
    """スクリーン座標の矩形同士の共通部分。無ければ None。"""
    cl, ct, cr, cb = client_screen
    rl, rt = r.left, r.top
    rr, rb = r.left + r.width, r.top + r.height
    il = max(rl, cl)
    it = max(rt, ct)
    ir = min(rr, cr)
    ib = min(rb, cb)
    if il >= ir or it >= ib:
        return None
    return CaptureRegion(left=il, top=it, width=ir - il, height=ib - it)


class ScreenCapturer:
    """mss ベースのキャプチャ。スレッドローカルに mss インスタンスを保持。"""

    def __init__(self) -> None:
        self._sct = None  # type: ignore[assignment]

    def _ensure_sct(self):
        if self._sct is None:
            import mss  # 遅延 import（PyInstaller の hiddenimport 関係でも有効）
            self._sct = mss.mss()
        return self._sct

    def close(self) -> None:
        if self._sct is not None:
            try:
                self._sct.close()
            except Exception:
                pass
            self._sct = None

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass

    def capture(
        self,
        region: CaptureRegion,
        *,
        target_hwnd: Optional[int] = None,
        client_rect_screen: Optional[Tuple[int, int, int, int]] = None,
    ) -> Image.Image:
        """指定領域をキャプチャして PIL Image (RGB) を返す。

        target_hwnd と client_rect_screen が与えられ、Windows で PrintWindow が成功すれば、
        デスクトップ合成ではなくウィンドウ内容のみを取得する（他ウィンドウの重なりを含まない）。
        """
        if not region.is_valid:
            raise ValueError(f"無効なキャプチャ領域: {region}")
        if (
            target_hwnd
            and client_rect_screen
            and sys.platform == "win32"
        ):
            pw = self._capture_via_printwindow(region, int(target_hwnd), client_rect_screen)
            if pw is not None:
                return pw
            log.debug("PrintWindow 失敗または領域外 — 画面キャプチャにフォールバックします")
        return self._capture_mss(region)

    def _capture_mss(self, region: CaptureRegion) -> Image.Image:
        sct = self._ensure_sct()
        monitor = region.to_mss_monitor()
        shot = sct.grab(monitor)
        return Image.frombytes("RGB", shot.size, shot.rgb)

    def _capture_via_printwindow(
        self,
        region: CaptureRegion,
        hwnd: int,
        client_rect_screen: Tuple[int, int, int, int],
    ) -> Optional[Image.Image]:
        """クライアント全体を PrintWindow し、スクリーン座標 region 相当を切り出す。"""
        try:
            import win32gui
            import win32ui
            from ctypes import windll
        except ImportError:
            return None

        if not win32gui.IsWindow(hwnd):
            return None

        cl, ct, cr, cb = client_rect_screen
        cw, ch = cr - cl, cb - ct
        if cw <= 0 or ch <= 0:
            return None

        inter = _intersect_screen_rects(region, client_rect_screen)
        if inter is None:
            return None

        hwnd_dc = win32gui.GetDC(hwnd)
        if not hwnd_dc:
            return None
        try:
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            save_bmp = win32ui.CreateBitmap()
            save_bmp.CreateCompatibleBitmap(mfc_dc, cw, ch)
            save_dc.SelectObject(save_bmp)
            PW_RENDERFULLCONTENT = 0x00000002
            hdc = int(save_dc.GetSafeHdc())
            user32 = windll.user32
            ok = bool(user32.PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT))
            if not ok:
                ok = bool(user32.PrintWindow(hwnd, hdc, 0))
            if not ok:
                return None

            bmpinfo = save_bmp.GetInfo()
            bmpstr = save_bmp.GetBitmapBits(True)
            full = Image.frombuffer(
                "RGB",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpstr,
                "raw",
                "BGRX",
                0,
                1,
            )
        finally:
            try:
                save_dc.DeleteDC()
            except Exception:
                pass
            try:
                mfc_dc.DeleteDC()
            except Exception:
                pass
            win32gui.ReleaseDC(hwnd, hwnd_dc)

        if full.size != (cw, ch):
            full = full.resize((cw, ch), Image.Resampling.LANCZOS)

        rel_l = inter.left - cl
        rel_t = inter.top - ct
        box = (rel_l, rel_t, rel_l + inter.width, rel_t + inter.height)
        try:
            return full.crop(box)
        except Exception as e:
            log.debug("PrintWindow 画像の切り出し失敗: %s", e)
            return None

    @staticmethod
    def is_blank(image: Image.Image, threshold: int = 5) -> bool:
        """ほぼ単色（黒/白）画像かどうかの粗い判定。

        ブラックスクリーン検知用。標準偏差が閾値未満なら blank とみなす。
        """
        try:
            small = image.resize((64, 64))
            stats = small.convert("L").getextrema()
            return (stats[1] - stats[0]) < threshold
        except Exception:
            return False
