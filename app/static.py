from __future__ import annotations

from pathlib import Path


APP_ROOT = Path(__file__).resolve().parent.parent
ASSETS_DIR = APP_ROOT / "app" / "assets"
FONTS_DIR = ASSETS_DIR / "fonts"


def resolve_font_asset(font_name: str) -> Path | None:
    safe_name = Path(font_name).name
    candidate = FONTS_DIR / safe_name
    if candidate.exists() and candidate.is_file():
        return candidate
    return None


def guess_content_type(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".woff2":
        return "font/woff2"
    if suffix == ".woff":
        return "font/woff"
    if suffix == ".ttf":
        return "font/ttf"
    if suffix == ".otf":
        return "font/otf"
    return "application/octet-stream"
