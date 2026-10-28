"""Извлечение слоя отчёта из Layout (страницы, поля, фильтры, форматирование)."""
from __future__ import annotations

from .conditional import ConditionalMixin
from .fields import FieldsMixin
from .filters import FiltersMixin
from .formatting import FormattingMixin
from .page import PageMixin
from .visuals import VisualsMixin


class LayoutExtractor(
    PageMixin,
    FieldsMixin,
    FiltersMixin,
    FormattingMixin,
    ConditionalMixin,
    VisualsMixin,
):
    """Извлечение листов и визуальных элементов из Layout."""

    def __init__(self, layout_data: dict, extractor):
        self.sections = layout_data.get("sections", [])
        self.ext = extractor


__all__ = ["LayoutExtractor"]
