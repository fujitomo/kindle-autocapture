"""User configuration management.

JSON で `%APPDATA%\\KindleAutoCapture\\config.json` を保存。
スキーマはコード内のデフォルト値とマージして欠損キーを補完する。
"""

from __future__ import annotations

import json
import threading
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, Optional

from utils.paths import config_path, default_output_dir
from utils.logger import get_logger

log = get_logger("config")


@dataclass
class HotkeyConfig:
    start_stop: str = "f9"
    pause: str = "f10"
    emergency_stop: str = "esc"


@dataclass
class CaptureConfig:
    method: str = "mss"  # "mss" のみサポート（将来 "win32" 等を追加可）
    page_delay_sec: float = 1.5
    pre_capture_delay_sec: float = 0.2
    save_format: str = "png"  # "png" / "jpeg" / "webp"
    jpeg_quality: int = 95
    webp_quality: int = 95
    skip_duplicates: bool = True
    duplicate_threshold: int = 8  # pHash ハミング距離（重複スキップ用・緩め）
    auto_stop_on_last_page: bool = True
    # 最終ページ用ストリークに加算するのは「この距離以下」のときだけ（隣接ページが distance≈4 になり得る）
    last_page_max_hash_distance: int = 2
    # last_page_max_hash_distance 以下がこの回数連続したら最終ページ候補
    last_page_consecutive_dupes: int = 6
    # この枚数未満の保存では「最終ページで自動停止」しない（少数ページで同一画面が続く誤検知防止）
    min_captures_before_last_page_stop: int = 5
    # 保存せずに重複スキップだけが続く上限（ページ送りが効いていないときの無限ループ防止）
    max_duplicate_skips_without_save: int = 60
    auto_trim: bool = False
    # 実行中は本アプリを最小化して Kindle にフォーカスを譲る（ページ送りが効かない時の決定打）
    minimize_during_capture: bool = True


@dataclass
class NavigationConfig:
    method: str = "key"  # "key" / "click"
    # Amazon 公式: 次ページは Page Down または右矢印（Page Down の方が送りやすいことが多い）
    key: str = "pagedown"
    # sendinput: 前面化＋keybd_event（本アプリ最小化と組合せれば最も確実）
    # wheel_pyautogui: マウス座標へ移動＋ホイール
    # wheel_postmessage: WM_MOUSEWHEEL を子 HWND へ
    # postmessage / pyautogui: 補助
    key_delivery: str = "sendinput"
    # wheel_* 用: 1回のページ送りで送るホイールノッチ数（3 だと3ページ分動くので既定は 1）
    wheel_scroll_notches: int = 1
    # pyautogui.scroll(sign * notches): Kindle では環境で前後が逆。戻りすぎる場合は +1 / -1 を切り替え
    wheel_scroll_sign: int = -1
    # True: 送るキーとホイール符号を「進行と逆」に入れ替える（多くの日本語縦書き・版差で必要）
    invert_page_turn: bool = True
    click_x_offset: int = 0  # ウィンドウ右端からのオフセット（負: ウィンドウ内側）
    click_y_offset: int = 0  # ウィンドウ縦中央からのオフセット
    refocus_before_action: bool = True


@dataclass
class WindowConfig:
    title_filter: str = "Kindle"  # 部分一致
    last_selected_hwnd: Optional[int] = None
    last_selected_title: Optional[str] = None


@dataclass
class StorageConfig:
    output_dir: str = ""  # 空なら default_output_dir() を使用
    book_subfolder_template: str = "book_{timestamp}"
    file_template: str = "page_{index:04d}"
    resume_from_existing: bool = True


@dataclass
class RoiConfig:
    """キャプチャ範囲。空（width/height=0）ならウィンドウ全体。"""
    use_window: bool = True
    x: int = 0
    y: int = 0
    width: int = 0
    height: int = 0


@dataclass
class AppConfig:
    capture: CaptureConfig = field(default_factory=CaptureConfig)
    navigation: NavigationConfig = field(default_factory=NavigationConfig)
    window: WindowConfig = field(default_factory=WindowConfig)
    storage: StorageConfig = field(default_factory=StorageConfig)
    roi: RoiConfig = field(default_factory=RoiConfig)
    hotkeys: HotkeyConfig = field(default_factory=HotkeyConfig)


class ConfigManager:
    """シングルトンとして使う想定のスレッドセーフな設定マネージャ。"""

    _instance: Optional["ConfigManager"] = None
    _lock = threading.Lock()

    def __new__(cls) -> "ConfigManager":
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._initialize()
            return cls._instance

    def _initialize(self) -> None:
        self._path: Path = config_path()
        self._rwlock = threading.RLock()
        self._config: AppConfig = AppConfig()
        self.load()
        if not self._config.storage.output_dir:
            self._config.storage.output_dir = str(default_output_dir())

    @property
    def path(self) -> Path:
        return self._path

    @property
    def config(self) -> AppConfig:
        with self._rwlock:
            return self._config

    def load(self) -> AppConfig:
        with self._rwlock:
            if not self._path.exists():
                log.info("config.json が見つからないためデフォルト値で生成: %s", self._path)
                self.save()
                return self._config
            try:
                raw = json.loads(self._path.read_text(encoding="utf-8"))
                self._config = _merge_into_appconfig(raw)
                log.info("設定読み込み完了: %s", self._path)
            except Exception as e:
                log.exception("設定読み込み失敗。デフォルトを使用: %s", e)
                self._config = AppConfig()
            return self._config

    def save(self) -> None:
        with self._rwlock:
            try:
                self._path.parent.mkdir(parents=True, exist_ok=True)
                tmp = self._path.with_suffix(".json.tmp")
                tmp.write_text(
                    json.dumps(asdict(self._config), ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                tmp.replace(self._path)
                log.debug("設定保存完了: %s", self._path)
            except Exception as e:
                log.exception("設定保存失敗: %s", e)

    def update(self, **patches: Any) -> AppConfig:
        """部分的にセクションを更新する。例: update(capture={"page_delay_sec": 2.0})"""
        with self._rwlock:
            current = asdict(self._config)
            for section, value in patches.items():
                if section not in current:
                    log.warning("未知の設定セクション: %s", section)
                    continue
                if isinstance(current[section], dict) and isinstance(value, dict):
                    current[section].update(value)
                else:
                    current[section] = value
            self._config = _merge_into_appconfig(current)
            self.save()
            return self._config


def _merge_into_appconfig(data: Dict[str, Any]) -> AppConfig:
    base = asdict(AppConfig())
    for k, v in (data or {}).items():
        if k in base and isinstance(base[k], dict) and isinstance(v, dict):
            base[k].update(v)
        else:
            base[k] = v
    return AppConfig(
        capture=CaptureConfig(**base["capture"]),
        navigation=NavigationConfig(**base["navigation"]),
        window=WindowConfig(**base["window"]),
        storage=StorageConfig(**base["storage"]),
        roi=RoiConfig(**base["roi"]),
        hotkeys=HotkeyConfig(**base["hotkeys"]),
    )


def get_config() -> AppConfig:
    return ConfigManager().config
