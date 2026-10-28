"""Поиск зависимостей в DAX-формуле (таблицы, столбцы, меры)."""
from __future__ import annotations

import re

from .icons import COLUMN_ICON, MEASURE_ICON, TABLE_ICON


class DependencyAnalyzer:
    """Анализ зависимостей в DAX формулах.
    Формат вывода как в эталоне:
    📊 s1
    📊 s12
    📊 s2
    📊 s3

    📋 A101
    📋 A102
    ...
    """
    PAT_QUOTED = re.compile(r"'([^']+)'\[([^\]]+)\]")
    PAT_UNQUOTED = re.compile(r'(?<!\w)([A-Za-zА-Яа-яёЁ0-9_]+)\[([^\]]+)\]')
    PAT_MEASURE = re.compile(r'\[([^\]]+)\]')

    def __init__(self, table_names: set, table_number: dict,
                 column_to_table: dict, measure_to_number: dict,
                 get_column_number_func):
        self.table_names = table_names
        self.table_number = table_number
        self.column_to_table = column_to_table
        self.measure_to_number = measure_to_number
        self.get_column_number = get_column_number_func

    def find_dependencies(self, formula: str, exclude_measure: str = None) -> str:
        if not formula:
            return "—"
        formula = str(formula)
        table_nums = set()
        col_nums = set()
        measure_nums = set()

        # Собираем все таблицы и столбцы из формулы
        for tbl, col in self.PAT_QUOTED.findall(formula):
            # Таблица
            if tbl in self.table_number:
                table_nums.add(self.table_number[tbl])
            # Столбец
            if tbl in self.table_names and col:
                col_num = self.get_column_number(tbl, col)
                if col_num:
                    col_nums.add(col_num)
            elif col in self.column_to_table:
                real_tbl = self.column_to_table[col]
                col_num = self.get_column_number(real_tbl, col)
                if col_num:
                    col_nums.add(col_num)

        for tbl, col in self.PAT_UNQUOTED.findall(formula):
            if tbl in self.table_names:
                # Таблица
                if tbl in self.table_number:
                    table_nums.add(self.table_number[tbl])
                # Столбец
                col_num = self.get_column_number(tbl, col)
                if col_num:
                    col_nums.add(col_num)

        # Меры
        for meas in self.PAT_MEASURE.findall(formula):
            if meas != exclude_measure:
                if meas in self.measure_to_number:
                    measure_nums.add(self.measure_to_number[meas])

        # Формируем вывод
        parts = []
        # Таблицы
        if table_nums:
            sorted_tables = sorted(table_nums)
            for t in sorted_tables:
                parts.append(f"{TABLE_ICON} {str(t).upper()}")
        # Разделитель
        if col_nums or measure_nums:
            if table_nums:
                parts.append("")  # пустая строка
        # Столбцы
        if col_nums:
            sorted_cols = sorted(col_nums)
            for c in sorted_cols:
                parts.append(f"{COLUMN_ICON} {str(c).upper()}")
        # Меры
        if measure_nums:
            if col_nums:
                parts.append("")
            sorted_measures = sorted(measure_nums)
            for m in sorted_measures:
                parts.append(f"{MEASURE_ICON} {str(m).upper()}")

        return "\n".join(parts) if parts else "—"

__all__ = ["DependencyAnalyzer"]
