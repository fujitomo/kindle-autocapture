"""連番画像保存.

- セッション開始時にサブフォルダを作成
- 既存ファイルから次の連番を継続（resume_from_existing=True 時）
- 一時ファイル → rename によるアトミック書き込み
"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from PIL import Image

from utils.logger import get_logger

log = get_logger("saver")

_FILENAME_INDEX_RE = re.compile(r"(?P<idx>\d{3,8})")


@dataclass
class SaveResult:
    path: Path
    index: int
    bytes_written: int


class ImageSaver:
    def __init__(
        self,
        output_root: Path,
        book_subfolder_template: str = "book_{timestamp}",
        file_template: str = "page_{index:04d}",
        save_format: str = "png",
        jpeg_quality: int = 95,
        webp_quality: int = 95,
        resume_from_existing: bool = True,
    ) -> None:
        self.output_root = Path(output_root)
        self.book_subfolder_template = book_subfolder_template
        self.file_template = file_template
        self.save_format = save_format.lower()
        self.jpeg_quality = jpeg_quality
        self.webp_quality = webp_quality
        self.resume_from_existing = resume_from_existing

        self._session_dir: Optional[Path] = None
        self._next_index: int = 1

    @property
    def session_dir(self) -> Optional[Path]:
        return self._session_dir

    @property
    def next_index(self) -> int:
        return self._next_index

    def start_session(self, override_dir: Optional[Path] = None) -> Path:
        """セッションを開始しサブフォルダを準備."""
        if override_dir is not None:
            self._session_dir = Path(override_dir)
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            sub = self.book_subfolder_template.format(timestamp=timestamp)
            self._session_dir = self.output_root / sub

        self._session_dir.mkdir(parents=True, exist_ok=True)
        log.info("セッションフォルダ: %s", self._session_dir)

        if self.resume_from_existing:
            self._next_index = self._scan_next_index(self._session_dir) or 1
            if self._next_index > 1:
                log.info("既存ファイル検出: 連番 %d から再開", self._next_index)
        else:
            self._next_index = 1
        return self._session_dir

    def save(self, image: Image.Image) -> SaveResult:
        if self._session_dir is None:
            raise RuntimeError("セッションが開始されていません。start_session() を呼んでください。")

        ext = self._extension()
        filename = self.file_template.format(index=self._next_index) + ext
        target = self._session_dir / filename
        tmp = target.with_suffix(ext + ".tmp")

        try:
            self._save_image(image, tmp)
            tmp.replace(target)
        except Exception:
            if tmp.exists():
                try:
                    tmp.unlink()
                except Exception:
                    pass
            raise

        size = target.stat().st_size
        result = SaveResult(path=target, index=self._next_index, bytes_written=size)
        self._next_index += 1
        return result

    def _extension(self) -> str:
        return {
            "png": ".png",
            "jpeg": ".jpg",
            "jpg": ".jpg",
            "webp": ".webp",
        }.get(self.save_format, ".png")

    def _save_image(self, image: Image.Image, path: Path) -> None:
        fmt = self.save_format
        if fmt in ("jpeg", "jpg"):
            img = image.convert("RGB") if image.mode != "RGB" else image
            img.save(path, format="JPEG", quality=self.jpeg_quality, optimize=True)
        elif fmt == "webp":
            image.save(path, format="WEBP", quality=self.webp_quality, method=4)
        else:
            image.save(path, format="PNG", optimize=False, compress_level=1)

    @staticmethod
    def _scan_next_index(folder: Path) -> Optional[int]:
        max_idx = 0
        try:
            for p in folder.iterdir():
                if not p.is_file():
                    continue
                m = _FILENAME_INDEX_RE.search(p.stem)
                if m:
                    try:
                        idx = int(m.group("idx"))
                        max_idx = max(max_idx, idx)
                    except ValueError:
                        continue
        except FileNotFoundError:
            return None
        return max_idx + 1 if max_idx > 0 else None

    @staticmethod
    def free_disk_bytes(path: Path) -> int:
        try:
            usage = shutil.disk_usage(str(path))
            return usage.free
        except Exception:
            return -1
