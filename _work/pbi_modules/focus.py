"""Разбор и нормализация идентификатора lineage-фокуса.

Один набор правил для GUI, lineage и CLI: ``S4`` / ``s04`` / ``A105`` / ``M3`` /
``VIS_ID :abc`` / ``VIS:abc`` / ``Тип | ID | Имя``.

Внутренние ID: ``s1`` (таблица), ``A101`` (столбец), ``M1`` (мера), ``VIS:<id>``.
"""
from __future__ import annotations

import re

VISUAL_ID_PATTERN = re.compile(r"VIS_ID\s*:\s*(.+)$", re.IGNORECASE)
FOCUS_ID_PATTERN = re.compile(r"(?i)^(s\d+|a\d+|m\d+|vis_id\s*:.+|vis:.+)$")

# Иконки, которые GUI добавляет к ID при отображении.
_ICON_PREFIXES = ("🔹", "📊", "📋", "📐", "🧬", "🧭", "💾", "🎨", "📑")


def extract_visual_id(value) -> str:
    """Чистый ID визуала: ``VIS_ID :abc`` → ``abc``, ``VIS:abc`` → ``abc``."""
    raw = str(value or "").strip()
    if not raw or raw == "None":
        return ""
    match = VISUAL_ID_PATTERN.match(raw)
    if match:
        return match.group(1).strip()
    if raw.upper().startswith("VIS:"):
        return raw[4:].strip()
    return raw


def _strip_icons(text: str) -> str:
    result = text
    for icon in _ICON_PREFIXES:
        result = result.replace(icon, "")
    return result.strip()


def normalize_focus_id(value) -> str:
    """Приводит фокус к внутреннему ID; ``""`` — если фокус не распознан."""
    raw = str(value or "").strip()
    if not raw:
        return ""

    # Подпись из выпадающего списка: «Тип | ID | Имя».
    if "|" in raw:
        for part in (item.strip() for item in raw.split("|")):
            if FOCUS_ID_PATTERN.match(part):
                return normalize_focus_id(part)
        return ""

    text = _strip_icons(raw)
    upper = text.upper()

    if upper.startswith("VIS_ID"):
        visual_id = extract_visual_id(text)
        return f"VIS:{visual_id}" if visual_id else ""
    if upper.startswith("VIS:"):
        return f"VIS:{text[4:].strip()}"

    if upper.startswith("T:"):
        return text[2:].strip()
    if upper.startswith("C:"):
        return text[2:].strip()
    if upper.startswith("M:"):
        return f"M{text[2:].strip().lstrip('M')}"

    if upper.startswith("S"):
        return f"s{text[1:].strip()}"
    if upper.startswith("A"):
        return f"A{text[1:].strip()}"
    if upper.startswith("M"):
        return f"M{text[1:].strip()}"

    return text


def focus_node_type(focus_id: str) -> str:
    """Тип узла по нормализованному ID."""
    value = str(focus_id or "").strip()
    if value.startswith("VIS:"):
        return "Визуал"
    if value.startswith("s"):
        return "Таблица"
    if value.startswith("A"):
        return "Столбец"
    if value.startswith("M"):
        return "Мера"
    return ""


def display_focus_id(node_type: str, node_id: str) -> str:
    """ID в том виде, в каком его видит пользователь."""
    if node_type == "Визуал":
        visual_id = node_id[4:] if str(node_id).startswith("VIS:") else node_id
        return f"VIS_ID :{visual_id}"
    return str(node_id).upper()


def parse_focus_label(label: str) -> tuple[str, str, str]:
    """Парсит подпись «Тип | ID | Имя» в (тип, ID, имя).

    Если формат не соблюдён, тип определяется по ID.
    """
    text = str(label or "").strip()
    parts = [part.strip() for part in text.split("|")]
    if len(parts) == 3 and parts[0] and parts[1]:
        return parts[0], normalize_focus_id(parts[1]), parts[2]

    focus_id = normalize_focus_id(text)
    return focus_node_type(focus_id), focus_id, ""


__all__ = [
    "VISUAL_ID_PATTERN",
    "FOCUS_ID_PATTERN",
    "extract_visual_id",
    "normalize_focus_id",
    "focus_node_type",
    "display_focus_id",
    "parse_focus_label",
]
