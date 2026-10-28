"""Версия приложения — единственный источник правды.

`build_exe.py` берёт версию отсюда и сверяет с первым заголовком
`docs/CHANGELOG.md`; расхождение — ошибка сборки, а не тихое `_v0`.
"""
from __future__ import annotations

import pathlib
import re

__version__ = "2.3.0"

CHANGELOG_HEADER_PATTERN = re.compile(r"^##\s*\[(\d+(?:\.\d+)*)\]", re.MULTILINE)


def changelog_version(changelog_path: pathlib.Path | str) -> str:
    """Версия из первого заголовка CHANGELOG или ``""``, если её нет."""
    try:
        text = pathlib.Path(changelog_path).read_text(encoding="utf-8")
    except OSError:
        return ""
    match = CHANGELOG_HEADER_PATTERN.search(text)
    return match.group(1) if match else ""


__all__ = ["__version__", "changelog_version"]
