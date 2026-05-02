"""Screen capture using mss (DXGI Desktop Duplication backend).

mss は Win32 BitBlt にも自動フォールバックするため、
GPU 合成画面でもまずこちらを推奨する。

注意: mss.mss() は使用するスレッドで生成すること。
"""

from __future__ import annotations

from dataclasses import dataclass

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

    def capture(self, region: CaptureRegion) -> Image.Image:
        """指定領域をキャプチャして PIL Image (RGB) を返す。"""
        if not region.is_valid:
            raise ValueError(f"無効なキャプチャ領域: {region}")
        sct = self._ensure_sct()
        monitor = region.to_mss_monitor()
        shot = sct.grab(monitor)
        # mss の bgra → PIL RGB に変換
        img = Image.frombytes("RGB", shot.size, shot.rgb)
        return img

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
