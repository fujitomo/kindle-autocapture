"""Application path resolution.

EXE化された場合と開発時で挙動が変わる箇所をここに集約する。
ユーザーデータは %APPDATA%\KindleAutoCapture を使用する。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "KindleAutoCapture"


def is_frozen() -> bool:
    """PyInstaller 等で凍結された実行ファイルかどうか。"""
    return getattr(sys, "frozen", False)


def app_root() -> Path:
    """ソースコードのルート（凍結時は実行ファイルの隣）。"""
    if is_frozen():
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def user_data_dir() -> Path:
    """ユーザーデータディレクトリ。なければ作成。"""
    appdata = os.environ.get("APPDATA")
    if appdata:
        base = Path(appdata) / APP_NAME
    else:
        base = Path.home() / f".{APP_NAME.lower()}"
    base.mkdir(parents=True, exist_ok=True)
    return base


def config_path() -> Path:
    return user_data_dir() / "config.json"


def logs_dir() -> Path:
    d = user_data_dir() / "logs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def default_output_dir() -> Path:
    """デフォルト保存先（Pictures\\KindleAutoCapture）。"""
    pictures = Path.home() / "Pictures"
    if not pictures.exists():
        pictures = Path.home()
    out = pictures / APP_NAME
    out.mkdir(parents=True, exist_ok=True)
    return out
