"""Capture controller state enum."""

from __future__ import annotations

from enum import Enum, auto


class CaptureState(Enum):
    IDLE = auto()
    RUNNING = auto()
    PAUSED = auto()
    STOPPING = auto()

    @property
    def label(self) -> str:
        return {
            CaptureState.IDLE: "停止中",
            CaptureState.RUNNING: "実行中",
            CaptureState.PAUSED: "一時停止中",
            CaptureState.STOPPING: "停止処理中",
        }[self]


class StopReason(Enum):
    USER = auto()
    LAST_PAGE = auto()
    ERROR = auto()
    MAX_PAGES = auto()
    DISK_FULL = auto()
    STUCK_DUPLICATES = auto()

    @property
    def label(self) -> str:
        return {
            StopReason.USER: "ユーザー操作による停止",
            StopReason.LAST_PAGE: "最終ページ検知により停止",
            StopReason.ERROR: "エラーにより停止",
            StopReason.MAX_PAGES: "上限ページ数に到達",
            StopReason.DISK_FULL: "ディスク空き容量不足",
            StopReason.STUCK_DUPLICATES: "同一画面のまま送り続けたため停止（ページ送りを確認）",
        }[self]
