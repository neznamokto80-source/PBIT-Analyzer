"""Совместимость: реализация переехала в пакет ``extract``.

Модуль оставлен точкой входа для существующих импортов:
``from pbi_modules.extractors import DataModelExtractor``.
"""
from .extract import (
    COLUMN_ICON,
    MEASURE_ICON,
    TABLE_ICON,
    DataModelExtractor,
    DependencyAnalyzer,
    LayoutExtractor,
    load_json_utf16,
)

__all__ = [
    "TABLE_ICON",
    "COLUMN_ICON",
    "MEASURE_ICON",
    "DataModelExtractor",
    "DependencyAnalyzer",
    "LayoutExtractor",
    "load_json_utf16",
]
