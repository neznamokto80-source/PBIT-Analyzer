"""
Определение типа входного файла: .pbit (шаблон с Layout) или .pbix (модель).

Расширение выбирает предполагаемый формат, содержимое zip его подтверждает:
    .pbit -> member DataModelSchema (+ Layout)
    .pbix -> member DataModel (сжатая модель VertiPaq)
Layout в .pbit лежит либо в корне архива, либо в папке Report
(Report/Layout в новых версиях Power BI Desktop) — поддержаны оба варианта.
Несовпадение — понятная ошибка, а не падение внутри парсера.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

PBIT = "pbit"
PBIX = "pbix"

SUPPORTED_SUFFIXES = {".pbit": PBIT, ".pbix": PBIX}


class FileFormatError(Exception):
    """Файл не является поддерживаемым .pbit/.pbix или повреждён."""


def _zip_members(path: str) -> set[str]:
    try:
        with zipfile.ZipFile(path) as archive:
            return set(archive.namelist())
    except zipfile.BadZipFile as error:
        raise FileFormatError(
            "Файл повреждён или не является архивом Power BI (.pbit/.pbix)."
        ) from error
    except OSError as error:
        raise FileFormatError(f"Не удалось открыть файл: {error}") from error


def _member_basename(member: str) -> str:
    """Имя файла внутри архива без папок: 'Report/Layout' -> 'Layout'."""
    return member.replace("\\", "/").rsplit("/", 1)[-1]


def has_archive_part(members: set[str], name: str) -> bool:
    """Есть ли часть архива с таким именем в корне или во вложенной папке."""
    return any(_member_basename(member) == name for member in members)


def _detect_by_content(members: set[str]) -> str | None:
    if "DataModelSchema" in members:
        return PBIT
    if "DataModel" in members:
        return PBIX
    return None


def detect_file_kind(path: str) -> str:
    """Возвращает PBIT/PBIX или бросает FileFormatError с объяснением."""
    suffix = Path(str(path)).suffix.lower()
    expected = SUPPORTED_SUFFIXES.get(suffix)
    if expected is None:
        raise FileFormatError(
            "Неподдерживаемый формат файла. Поддерживаются .pbit и .pbix."
        )

    members = _zip_members(path)
    actual = _detect_by_content(members)
    if actual is None:
        raise FileFormatError(
            "В архиве нет ни модели PBIX (DataModel), ни шаблона PBIT "
            "(DataModelSchema). Файл не является файлом Power BI."
        )

    if actual != expected:
        hint = ".pbix" if actual == PBIX else ".pbit"
        raise FileFormatError(
            f"Расширение файла не совпадает с содержимым: это {actual.upper()}. "
            f"Переименуйте файл в *{hint} и повторите."
        )

    if actual == PBIT and not has_archive_part(members, "Layout"):
        raise FileFormatError(
            "В PBIT-архиве нет Layout (ни в корне, ни в папке Report): "
            "файл неполный или повреждён."
        )

    return actual


__all__ = [
    "PBIT",
    "PBIX",
    "SUPPORTED_SUFFIXES",
    "FileFormatError",
    "detect_file_kind",
    "has_archive_part",
]
