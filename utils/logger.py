"""Centralized logger.

- ファイル: %APPDATA%\\KindleAutoCapture\\logs\\app_YYYYMMDD.log（日付ローテーション）
- コンソール: stderr
- GUI: PyQt6 シグナルへ流すための専用ハンドラを別ファイルに置かず、
  ここで `QtSignalLogHandler` も提供する（importは遅延）。
"""

from __future__ import annotations

import logging
import logging.handlers
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from .paths import logs_dir

_INITIALIZED = False
LOGGER_NAME = "kindle_autocapture"

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """名前付きロガーを返す。初回呼び出し時に全体設定を初期化する。"""
    global _INITIALIZED
    if not _INITIALIZED:
        _setup_root_logger()
        _INITIALIZED = True
    if name:
        return logging.getLogger(f"{LOGGER_NAME}.{name}")
    return logging.getLogger(LOGGER_NAME)


def _setup_root_logger() -> None:
    root = logging.getLogger(LOGGER_NAME)
    root.setLevel(logging.DEBUG)
    root.propagate = False

    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)

    log_file: Path = logs_dir() / f"app_{datetime.now():%Y%m%d}.log"
    file_handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    root.addHandler(file_handler)

    stream_handler = logging.StreamHandler(stream=sys.stderr)
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)
    root.addHandler(stream_handler)


class QtSignalLogHandler(logging.Handler):
    """QObject の pyqtSignal にログをブロードキャストするハンドラ。

    GUIスレッドで安全に表示するため、シグナル経由で送る前提。
    """

    def __init__(self, signal_emit_callable):
        super().__init__()
        self._emit = signal_emit_callable
        self.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            self._emit(record.levelname, msg)
        except Exception:
            self.handleError(record)
