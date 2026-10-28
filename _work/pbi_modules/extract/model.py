"""Семантическая модель PBIT: таблицы, столбцы, источники, связи, меры."""
from __future__ import annotations

import re

from ..formatters import DAXFormatter, SourceFormatter
from ..ids import column_number, measure_number, table_number
from ..schemas import frame_for
from .dependencies import DependencyAnalyzer

SYSTEM_TABLE_PREFIXES = ("DateTableTemplate", "LocalDateTable")


def table_name_of(table: dict) -> str:
    """Имя таблицы модели; ``''`` для битых записей без имени."""
    name = table.get("name")
    return name.strip() if isinstance(name, str) else ""


def is_system_table(name: str) -> bool:
    """Служебные таблицы дат, которые Power BI создаёт сам."""
    return not name or name.startswith(SYSTEM_TABLE_PREFIXES)


class DataModelExtractor:
    def __init__(self, model_data: dict):
        self.model = model_data.get("model", {})
        self.tables_raw = self.model.get("tables", [])
        # all_tables — видимые и скрытые: меры скрытых таблиц тоже нужны.
        self.all_tables = [
            t for t in self.tables_raw if not is_system_table(table_name_of(t))
        ]
        self.tables = [t for t in self.all_tables if not t.get("isHidden", False)]
        self.table_names = {table_name_of(t) for t in self.tables}
        self.table_number = {}
        for idx, t in enumerate(self.tables, 1):
            self.table_number[table_name_of(t)] = table_number(idx)

        self.column_to_table: dict[str, str] = {}
        self.measure_to_table: dict[str, str] = {}
        self.measure_to_number: dict[str, str] = {}
        self.column_to_number: dict[tuple[str, str], str] = {}
        self.norm_measure_to_name: dict[str, str] = {}
        self.norm_column_to_name: dict[str, str] = {}
        self.norm_table_to_name: dict[str, str] = {}
        self._build_lookups()

        # Инициализируем AFTER _build_lookups
        self.dep_analyzer = DependencyAnalyzer(
            self.table_names,
            self.table_number,
            self.column_to_table,
            self.measure_to_number,
            self.get_column_number,
        )

    @staticmethod
    def _normalize(s: str) -> str:
        return re.sub(r'[\s_\-]+', '', s).lower()

    def _build_lookups(self):
        m_counter = 0
        for t in self.all_tables:
            t_name = table_name_of(t)
            visible = t in self.tables
            if visible:
                self.norm_table_to_name[self._normalize(t_name)] = t_name
                t_num = self.table_number.get(t_name, "")
                for ci, col in enumerate(t.get("columns", []), 1):
                    c_name = col["name"]
                    self.column_to_table[c_name] = t_name
                    self.norm_column_to_name[self._normalize(c_name)] = c_name
                    if c_name and t_num:
                        # O(1)-поиск номера столбца вместо обхода таблиц в цикле.
                        self.column_to_number[(t_name, c_name)] = column_number(t_num, ci)
            # Меры собираются по всем таблицам: у меры скрытой таблицы номер
            # таблицы будет «—», но сама мера и её зависимости не теряются.
            for m in t.get("measures", []):
                m_counter += 1
                m_name = m["name"]
                self.measure_to_table[m_name] = t_name
                self.measure_to_number[m_name] = measure_number(m_counter)
                self.norm_measure_to_name[self._normalize(m_name)] = m_name

    def get_column_number(self, table_name: str, col_name: str) -> str:
        """Номер столбца из индекса, построенного в _build_lookups (O(1))."""
        return self.column_to_number.get((table_name, col_name), "")

    def extract_tables_ref(self):
        rows = []
        for t in self.tables:
            rows.append({
                "№ таблицы": self.table_number[t["name"]],
                "Имя таблицы": t["name"],
                "Описание": t.get("description", "—")
            })
        return frame_for(rows, "tables_ref")

    def extract_columns(self):
        rows = []
        for t in self.tables:
            t_num = self.table_number[t["name"]]
            for ci, col in enumerate(t.get("columns", []), 1):
                col_num = column_number(t_num, ci)
                col_type = col.get("type", "")
                data_type = col.get("dataType", "")
                if col_type == "calculated":
                    value_type = "Вычисляемый столбец"
                elif col_type == "calculatedTableColumn":
                    value_type = "Столбец вычисляемой таблицы"
                else:
                    dt_map = {
                        "int64": "Целое число", "double": "Десятичное число",
                        "string": "Строка", "dateTime": "Дата/время",
                        "boolean": "Логическое", "currency": "Денежный",
                        "binary": "Двоичный",
                    }
                    value_type = dt_map.get(data_type, data_type)
                formula = "—"
                expr = col.get("expression")
                col_name = col["name"]
                if expr:
                    formula = DAXFormatter.format_column(col_name, expr)
                elif col_type == "calculated":
                    formula = "Нет формулы"
                deps = "—"
                if expr:
                    deps = self.dep_analyzer.find_dependencies(expr)
                description = col.get("description", "—")
                format_string = col.get("formatString", "—")
                sort_by_column = col.get("sortByColumn", "—")
                summarize_by = col.get("summarizeBy", "—")
                rows.append({
                    "№ столбца": col_num, "№ таблицы": t_num, "Таблица": t["name"],
                    "Столбец": col["name"], "Описание": description,
                    "Тип данных": data_type, "Тип значения": value_type,
                    "Формат": format_string, "Сортировка по столбцу": sort_by_column,
                    "Агрегация по умолчанию": summarize_by,
                    "Формула (если вычисляемый)": formula, "Зависимости": deps,
                })
        return frame_for(rows, "columns")

    def extract_sources(self):
        rows = []
        for t in self.tables:
            t_num = self.table_number[t["name"]]
            partitions = t.get("partitions", [])
            if not partitions:
                rows.append({
                    "№ таблицы": t_num, "Таблица": t["name"],
                    "Тип источника": "Не определён", "Как создается": "—",
                    "Зависимости": "—",
                })
                continue
            for part in partitions:
                source = part.get("source", part)
                src_type = source.get("type", part.get("sourceType", ""))
                expression = source.get("expression", part.get("expression", ""))
                partition_name = part.get("name", t["name"])
                if isinstance(expression, list):
                    expression = "".join(str(x) for x in expression)
                type_name = {
                    "m": "Power Query (M)", "calculated": "Вычисляемая таблица (DAX)",
                    "entity": "Entity", "structuredData": "Structured Data",
                }.get(src_type, src_type or "Не определён")
                deps = "—"
                if src_type == "calculated" and expression:
                    deps = self.dep_analyzer.find_dependencies(expression)
                if src_type == "calculated":
                    formatted = DAXFormatter.format_table(partition_name, expression)
                else:
                    formatted = SourceFormatter.format_source(src_type, expression)
                rows.append({
                    "№ таблицы": t_num, "Таблица": t["name"],
                    "Тип источника": type_name,                    "Как создается": formatted or "—",
                    "Зависимости": deps,
                })
        return frame_for(rows, "sources")

    def extract_relationships(self):
        rows = []
        for ri, r in enumerate(self.model.get("relationships", []), 1):
            from_tbl = r.get("fromTable", "")
            from_col = r.get("fromColumn", "")
            to_tbl = r.get("toTable", "")
            to_col = r.get("toColumn", "")
            is_active = r.get("isActive", True)
            from_num = self.table_number.get(from_tbl, "?")
            to_num = self.table_number.get(to_tbl, "?")
            rows.append({
                "Название связи": r.get("name", f"Связь {ri}"),
                "Откуда": f"{from_tbl}[{from_col}] ({from_num})",
                "Куда": f"{to_tbl}[{to_col}] ({to_num})",
                "Состояние": "Активна" if is_active else "Не активна",
            })
        return frame_for(rows, "relationships")

    def extract_measures(self):
        rows = []
        for t in self.all_tables:
            # У скрытой таблицы нет номера в справочнике — прочерк, а не выдуманный ID.
            t_num = self.table_number.get(table_name_of(t), "—")
            for m in t.get("measures", []):
                m_num = self.measure_to_number[m["name"]]
                expr = m.get("expression", "")
                m_name = m["name"]
                # Нормализуем expression: если это список строк, объединяем в одну строку
                if isinstance(expr, list):
                    expr = "".join(str(x) for x in expr)
                if expr:
                    formula = DAXFormatter.format_measure(m_name, expr)
                else:
                    formula = "—"
                fmt = m.get("formatString", "")
                if not fmt:
                    fsd = m.get("formatStringDefinition", {})
                    fmt = fsd.get("expression", "") if isinstance(fsd, dict) else ""
                if not fmt:
                    for ann in m.get("annotations", []):
                        if ann.get("name") == "Format":
                            fmt = ann.get("value", "")
                            break
                fmt = fmt or "—"
                folder = m.get("displayFolder", "—")
                description = m.get("description", "—")
                deps = self.dep_analyzer.find_dependencies(expr, exclude_measure=m_name) if expr else "—"
                rows.append({
                    "№ меры": m_num, "№ таблицы": t_num, "Таблица": table_name_of(t),
                    "Мера": m_name, "Описание": description, "Формула (DAX)": formula,
                    "Формат": fmt, "Папка отображения": folder, "Зависимости": deps,
                })
        return frame_for(rows, "measures")

__all__ = ["DataModelExtractor", "table_name_of", "is_system_table"]
