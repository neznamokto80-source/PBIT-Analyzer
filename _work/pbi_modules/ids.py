"""Формирование ID объектов модели: ``s3``, ``A201``, ``M12``.

Столбец кодирует позицию таблицы и позицию столбца: ``A`` + номер таблицы
(без ``s``) + двузначный номер столбца, например ``A101``.
"""
from __future__ import annotations


def table_number(index: int) -> str:
    """ID таблицы по её позиции в справочнике (с 1)."""
    return f"s{index}"


def measure_number(index: int) -> str:
    """ID меры по её позиции в модели (с 1)."""
    return f"M{index}"


def column_number(table_number_value: str, column_index: int) -> str:
    """ID столбца: ``A`` + номер таблицы + двузначный номер столбца."""
    table_digits = str(table_number_value or "").strip().lstrip("sS")
    return f"A{table_digits}{int(column_index):02d}"


__all__ = ["table_number", "measure_number", "column_number"]
