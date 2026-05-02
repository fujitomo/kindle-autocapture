"""pHash ベースの重複ページ検知.

連続して重複検知された回数も保持し、最終ページ判定に使う。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

import imagehash
from PIL import Image

from utils.logger import get_logger

log = get_logger("duplicate")


@dataclass
class DuplicateResult:
    """is_duplicate: duplicate_threshold 以下（スキップ送り用・緩い）。

    last_page_streak: last_page_max_hash_distance 以下のときだけ加算（最終ページ判定用・厳しい）。
    隣接ページが pHash で 4～8 程度離れることがあり、緩い重複だけで最終ページ判定すると誤爆する。
    """

    is_duplicate: bool
    distance: int
    consecutive_dupes: int  # last_page_streak と同値（ログ互換のため残す）
    last_page_streak: int = 0


class DuplicateDetector:
    """pHash で重複/最終ページを判定する."""

    def __init__(
        self,
        duplicate_threshold: int = 8,
        last_page_max_distance: int = 2,
        hash_size: int = 16,
    ) -> None:
        self.duplicate_threshold = duplicate_threshold
        self.last_page_max_distance = last_page_max_distance
        self.hash_size = hash_size
        self._last_hash: Optional[imagehash.ImageHash] = None
        self._last_page_streak: int = 0
        self._history: List[imagehash.ImageHash] = []

    def reset(self) -> None:
        self._last_hash = None
        self._last_page_streak = 0
        self._history.clear()

    @property
    def consecutive_duplicates(self) -> int:
        return self._last_page_streak

    def check(self, image: Image.Image) -> DuplicateResult:
        current = imagehash.phash(image, hash_size=self.hash_size)
        if self._last_hash is None:
            self._last_hash = current
            self._history.append(current)
            return DuplicateResult(
                is_duplicate=False,
                distance=-1,
                consecutive_dupes=0,
                last_page_streak=0,
            )

        # imagehash の差分は numpy スカラーになることがあり、
        # bool 比較の結果が numpy.bool_ になり PyQt の bool シグナルと非互換になる。
        distance = int(current - self._last_hash)
        is_loose_dup = bool(distance <= int(self.duplicate_threshold))

        if not is_loose_dup:
            self._last_page_streak = 0
            self._last_hash = current
            self._history.append(current)
            return DuplicateResult(
                is_duplicate=False,
                distance=distance,
                consecutive_dupes=0,
                last_page_streak=0,
            )

        # 緩い重複（スキップ送り）だが、最終ページストリークは「ほぼ同一」のときだけ積む
        if distance <= int(self.last_page_max_distance):
            self._last_page_streak += 1
        else:
            self._last_page_streak = 0

        return DuplicateResult(
            is_duplicate=True,
            distance=distance,
            consecutive_dupes=self._last_page_streak,
            last_page_streak=self._last_page_streak,
        )

    def is_likely_last_page(self, required_consecutive: int = 2) -> bool:
        """厳密ストリークが N 以上なら最終ページ候補。"""
        return bool(self._last_page_streak >= int(required_consecutive))
