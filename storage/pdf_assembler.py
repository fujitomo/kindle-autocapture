"""セッション保存画像から 1 つの PDF を組み立てる."""

from __future__ import annotations

import re
from pathlib import Path
from typing import List

from PIL import Image

from utils.logger import get_logger

log = get_logger("pdf")

_FILENAME_INDEX_RE = re.compile(r"(?P<idx>\d{3,8})")
_IMAGE_EXTS = frozenset({".png", ".jpg", ".jpeg", ".webp"})


def collect_ordered_images(session_dir: Path) -> List[Path]:
    """連番ファイル名の画像のみをインデックス順に返す（tmp・PDF は除外）。"""
    items: list[tuple[int, Path]] = []
    for p in session_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() == ".pdf":
            continue
        if p.suffix.lower() not in _IMAGE_EXTS:
            continue
        if ".tmp" in p.name or p.name.endswith(".tmp"):
            continue
        m = _FILENAME_INDEX_RE.search(p.stem)
        if not m:
            continue
        items.append((int(m.group("idx")), p))
    items.sort(key=lambda x: x[0])
    return [p for _, p in items]


def _to_rgb_detached(im: Image.Image) -> Image.Image:
    """元の `Image.open` ハンドルから切り離した RGB 画像を返す。

    `with Image.open(...) as im:` を抜けるとファイルが閉じられるため、
    必ず新しいオブジェクトを返してメモリ上に展開しておく。
    """
    if im.mode == "RGB":
        return im.copy()
    if im.mode == "RGBA":
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[3])
        return bg
    if im.mode == "LA":
        rgba = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[3])
        return bg
    if im.mode == "P":
        if "transparency" in im.info:
            return _to_rgb_detached(im.convert("RGBA"))
        return im.convert("RGB")
    return im.convert("RGB")


def build_session_pdf(session_dir: Path, pdf_path: Path) -> int:
    """session_dir 内の連番画像から PDF を生成。戻り値はページ数。

    画像が 0 件のときはファイルを作らず 0 を返す。
    """
    paths = collect_ordered_images(session_dir)
    if not paths:
        log.info("PDF スキップ: 画像ファイルがありません (%s)", session_dir)
        return 0

    images_rgb: list[Image.Image] = []
    skipped: list[Path] = []
    try:
        for p in paths:
            try:
                with Image.open(p) as im:
                    im.load()
                    images_rgb.append(_to_rgb_detached(im))
            except Exception as e:
                skipped.append(p)
                log.warning("PDF: 画像を読めずスキップ %s: %s", p, e)

        if not images_rgb:
            raise RuntimeError(
                f"読み込める画像が 1 枚もありません（候補 {len(paths)} 件、すべてスキップ）"
            )

        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        first, *rest = images_rgb
        first.save(
            pdf_path,
            "PDF",
            resolution=100.0,
            save_all=True,
            append_images=rest,
        )
        if skipped:
            log.info(
                "PDF 作成: %s (%d ページ, スキップ %d)",
                pdf_path, len(images_rgb), len(skipped),
            )
        else:
            log.info("PDF 作成: %s (%d ページ)", pdf_path, len(images_rgb))
        return len(images_rgb)
    finally:
        for im in images_rgb:
            try:
                im.close()
            except Exception:
                pass
