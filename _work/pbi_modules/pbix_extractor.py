"""
Адаптер PBIXRay -> DataFrame-схемы проекта.

PBIXRay читает семантическую модель .pbix (и .abf/.xlsx), но не слой отчёта.
Модуль приводит метаданные модели к тем же таблицам, которые строит
DataModelExtractor для .pbit: это позволяет переиспользовать lineage,
Excel-отчёт и GUI без изменений в существующей логике.

Идентификаторы объектов те же, что в .pbit:
    таблицы  — s1, s2, ...   (в Excel — S1, S2, ...)
    столбцы  — A101, A102, ... (A{номер таблицы}{порядковый номер})
    меры     — M1, M2, ...   (сквозная нумерация по порядку модели)
"""
from __future__ import annotations

import logging
from typing import Any

import pandas as pd

from .extractors import COLUMN_ICON, MEASURE_ICON, TABLE_ICON, DependencyAnalyzer
from .formatters import DAXFormatter, SourceFormatter

logger = logging.getLogger(__name__)

# TOM Column.Type / DataType / SummarizationType -> имена как в DataModelSchema (.pbit)
COLUMN_TYPE_MAP = {
    1: "data",
    2: "calculated",
    3: "calculatedTableColumn",
    4: "rowNumber",
}
DATA_TYPE_MAP = {
    1: "automatic",
    2: "string",
    6: "int64",
    8: "double",
    9: "dateTime",
    10: "currency",
    11: "boolean",
    17: "binary",
}
VALUE_TYPE_MAP = {
    "int64": "Целое число",
    "double": "Десятичное число",
    "string": "Строка",
    "dateTime": "Дата/время",
    "boolean": "Логическое",
    "currency": "Денежный",
    "binary": "Двоичный",
}
SUMMARIZE_MAP = {
    0: "GroupBy",
    1: "Sum",
    2: "Count",
    3: "Min",
    4: "Max",
}

TABLES_REF_COLUMNS = ["№ таблицы", "Имя таблицы", "Описание"]
COLUMNS_SCHEMA = [
    "№ столбца", "№ таблицы", "Таблица", "Столбец", "Описание", "Тип данных",
    "Тип значения", "Формат", "Сортировка по столбцу", "Агрегация по умолчанию",
    "Формула (если вычисляемый)", "Зависимости",
]
MEASURES_SCHEMA = [
    "№ меры", "№ таблицы", "Таблица", "Мера", "Описание", "Формула (DAX)",
    "Формат", "Папка отображения", "Зависимости",
]
SOURCES_SCHEMA = ["№ таблицы", "Таблица", "Тип источника", "Как создается", "Зависимости"]
RELATIONSHIPS_SCHEMA = ["Название связи", "Откуда", "Куда", "Состояние"]
DAX_SCHEMA = ["Тип объекта", "№ таблицы", "Таблица", "Имя", "Выражение"]
POWER_QUERY_SCHEMA = ["Таблица", "Запрос (M)"]
M_PARAMETERS_SCHEMA = ["Параметр", "Описание", "Выражение", "Изменён"]
RLS_SCHEMA = ["Роль", "Таблица", "Фильтр", "Состояние", "Описание"]
OLS_SCHEMA = ["Роль", "Таблица", "Столбец", "Область", "Права"]
PERSPECTIVES_SCHEMA = ["Перспектива", "Тип объекта", "Таблица", "Объект", "Включить все"]
STATISTICS_SCHEMA = ["Таблица", "Столбец", "Кардинальность", "Словарь", "HashIndex", "DataSize"]


def _clean(value: Any) -> str:
    """Приводит значение из DataFrame PBIXRay к строке без NaN/None."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value)
    return "" if text in {"nan", "NaT"} else text


def _as_frame(value: Any, columns: list[str]) -> pd.DataFrame:
    """Безопасно превращает свойство PBIXRay в DataFrame с ожидаемыми колонками."""
    if isinstance(value, pd.DataFrame):
        frame = value.copy()
        for column in columns:
            if column not in frame.columns:
                frame[column] = ""
        return frame
    return pd.DataFrame(columns=columns)


def _find_col(frame: pd.DataFrame, *names: str) -> str:
    """Возвращает первое существующее имя колонки (PBIXRay иногда меняет регистр)."""
    lowered = {str(column).lower(): column for column in frame.columns}
    for name in names:
        if name in frame.columns:
            return name
        if name.lower() in lowered:
            return lowered[name.lower()]
    return ""


class PBIXModelExtractor:
    """Извлекает метаданные модели PBIX в DataFrame-схемы проекта.

    Объект модели передаётся снаружи (duck typing): в тестах достаточно
    заменить его на фейк с нужными атрибутами, реальный файл не требуется.
    """

    def __init__(self, model: Any):
        self.model = model
        self.warnings: list[str] = []
        self.tables = self._read_tables()
        self.table_names = {name for name, _ in self.tables}
        self.table_number = {name: f"s{idx}" for idx, (name, _) in enumerate(self.tables, 1)}
        self._table_columns: dict[str, list[str]] = {}
        self.column_to_table: dict[str, str] = {}
        self.measure_to_number: dict[str, str] = {}
        self._column_rows: list[dict] = []
        self._measure_rows: list[dict] = []
        self._column_meta: list[tuple[str, str, str, str, Any]] = []
        self._measure_meta: list[tuple[str, str, str, str, Any]] = []
        self._collect_objects()
        self.dep_analyzer = DependencyAnalyzer(
            self.table_names,
            self.table_number,
            self.column_to_table,
            self.measure_to_number,
            self.get_column_number,
        )
        self._build_rows()

    # ------------------------------------------------------------
    # Чтение исходных данных
    # ------------------------------------------------------------

    def _get(self, attr: str, columns: list[str] | None = None) -> pd.DataFrame:
        try:
            value = getattr(self.model, attr)
        except Exception as error:  # noqa: BLE001 — блок необязателен, отдаём частичный результат
            self.warnings.append(f"{attr}: {type(error).__name__}: {error}")
            logger.warning("Не удалось прочитать %s: %s", attr, error)
            return pd.DataFrame(columns=columns or [])
        return _as_frame(value, columns or [])

    def _read_tables(self) -> list[tuple[str, dict]]:
        frame = self._get("tmschema_tables", ["Name", "Description", "IsHidden"])
        table_order: list[str] = []
        try:
            table_order = [str(name) for name in list(self.model.tables)]
        except Exception as error:  # noqa: BLE001
            self.warnings.append(f"tables: {type(error).__name__}: {error}")

        meta: dict[str, dict] = {}
        for _, row in frame.iterrows():
            name = _clean(row.get("Name"))
            if name:
                meta[name] = row.to_dict()

        # Полный порядок таблиц модели: видимые нужны для «Справочника таблиц»,
        # а скрытые — чтобы не потерять меры, закреплённые за скрытой таблицей.
        names = list(table_order) or []
        for name in meta:
            if name not in names:
                names.append(name)
        self.all_table_names = names

        result: list[tuple[str, dict]] = []
        for name in names:
            info = meta.get(name, {})
            hidden = _clean(info.get("IsHidden")) in {"1", "True", "true"}
            system_table = name.startswith("DateTableTemplate") or name.startswith("LocalDateTable")
            if hidden or system_table:
                continue
            # tmschema_tables содержит только пользовательские таблицы;
            # всё остальное в model.tables — служебные объекты движка.
            if meta and name not in meta:
                continue
            result.append((name, info))
        return result

    def _collect_objects(self) -> None:
        columns_frame = self._get(
            "tmschema_columns",
            ["TableName", "Name", "Type", "DataType", "Description", "Expression",
             "FormatString", "SummarizeBy", "DisplayOrdinal"],
        )
        for table_name, _ in self.tables:
            rows = [
                row for _, row in columns_frame.iterrows()
                if _clean(row.get("TableName")) == table_name
            ]
            if "DisplayOrdinal" in columns_frame.columns:
                rows.sort(key=lambda row: _sort_ordinal(row.get("DisplayOrdinal")))
            table_cols: list[str] = []
            t_num = self.table_number[table_name]
            for ci, row in enumerate(rows, 1):
                col_name = _clean(row.get("Name"))
                if not col_name:
                    continue
                table_cols.append(col_name)
                self.column_to_table.setdefault(col_name, table_name)
                col_num = f"A{t_num[1:]}{ci:02d}"
                self._column_meta.append((table_name, t_num, col_num, col_name, row))
            self._table_columns[table_name] = table_cols

        measures_frame = self._get(
            "dax_measures",
            ["TableName", "Name", "Expression", "DisplayFolder", "Description", "FormatString"],
        )
        # Меры собираются по всем таблицам модели, включая скрытые: мера — это
        # объект модели, и её не должно скрывать только место хранения.
        measure_tables: list[str] = list(self.all_table_names)
        for _, row in measures_frame.iterrows():
            table_name = _clean(row.get("TableName"))
            if table_name and table_name not in measure_tables:
                measure_tables.append(table_name)

        m_counter = 0
        for table_name in measure_tables:
            rows = [
                row for _, row in measures_frame.iterrows()
                if _clean(row.get("TableName")) == table_name
            ]
            t_num = self.table_number.get(table_name, "")
            for row in rows:
                measure_name = _clean(row.get("Name"))
                if not measure_name:
                    continue
                m_counter += 1
                m_num = f"M{m_counter}"
                self.measure_to_number[measure_name] = m_num
                self._measure_meta.append(
                    (table_name, t_num, m_num, measure_name, row)
                )

    def _build_rows(self) -> None:
        """Строит строки после создания DependencyAnalyzer (нужен для зависимостей)."""
        for table_name, t_num, col_num, col_name, row in self._column_meta:
            self._column_rows.append(self._build_column_row(table_name, t_num, col_num, col_name, row))
        for table_name, t_num, m_num, m_name, row in self._measure_meta:
            self._measure_rows.append(self._build_measure_row(table_name, t_num, m_num, m_name, row))

    # ------------------------------------------------------------
    # Строки таблиц
    # ------------------------------------------------------------

    def get_column_number(self, table_name: str, col_name: str) -> str:
        t_num = self.table_number.get(table_name)
        if not t_num:
            return ""
        for ci, name in enumerate(self._table_columns.get(table_name, []), 1):
            if name == col_name:
                return f"A{t_num[1:]}{ci:02d}"
        return ""

    def _build_column_row(self, table_name: str, t_num: str, col_num: str, col_name: str, row: Any) -> dict:
        col_type = COLUMN_TYPE_MAP.get(_to_int(row.get("Type")), "")
        data_type = DATA_TYPE_MAP.get(_to_int(row.get("DataType")), _clean(row.get("DataType")))
        if col_type == "calculated":
            value_type = "Вычисляемый столбец"
        elif col_type == "calculatedTableColumn":
            value_type = "Столбец вычисляемой таблицы"
        else:
            value_type = VALUE_TYPE_MAP.get(data_type, data_type)

        expression = row.get("Expression")
        formula = "—"
        deps = "—"
        if expression is not None and _clean(expression):
            formula = DAXFormatter.format_column(col_name, expression)
            deps = self.dep_analyzer.find_dependencies(_clean(expression))
        elif col_type == "calculated":
            formula = "Нет формулы"

        return {
            "№ столбца": col_num,
            "№ таблицы": t_num,
            "Таблица": table_name,
            "Столбец": col_name,
            "Описание": _clean(row.get("Description")) or "—",
            "Тип данных": data_type or "—",
            "Тип значения": value_type or "—",
            "Формат": _clean(row.get("FormatString")) or "—",
            # В TMSCHEMA-выгрузке нет sortByColumn — оставляем прочерк.
            "Сортировка по столбцу": "—",
            "Агрегация по умолчанию": SUMMARIZE_MAP.get(_to_int(row.get("SummarizeBy")), _clean(row.get("SummarizeBy")) or "—"),
            "Формула (если вычисляемый)": formula,
            "Зависимости": deps,
        }

    def _build_measure_row(self, table_name: str, t_num: str, m_num: str, m_name: str, row: Any) -> dict:
        expr = _clean(row.get("Expression"))
        formula = DAXFormatter.format_measure(m_name, expr) if expr else "—"
        deps = self.dep_analyzer.find_dependencies(expr, exclude_measure=m_name) if expr else "—"
        return {
            "№ меры": m_num,
            # У скрытой таблицы нет номера в «Справочнике таблиц» — будет прочерк.
            "№ таблицы": t_num or "—",
            "Таблица": table_name,
            "Мера": m_name,
            "Описание": _clean(row.get("Description")) or "—",
            "Формула (DAX)": formula,
            "Формат": _clean(row.get("FormatString")) or "—",
            "Папка отображения": _clean(row.get("DisplayFolder")) or "—",
            "Зависимости": deps,
        }

    # ------------------------------------------------------------
    # Публичный API
    # ------------------------------------------------------------

    def extract_tables_ref(self) -> pd.DataFrame:
        rows = [
            {
                "№ таблицы": self.table_number[name],
                "Имя таблицы": name,
                "Описание": _clean(info.get("Description")) or "—",
            }
            for name, info in self.tables
        ]
        return pd.DataFrame(rows, columns=TABLES_REF_COLUMNS)

    def extract_columns(self) -> pd.DataFrame:
        return pd.DataFrame(self._column_rows, columns=COLUMNS_SCHEMA)

    def extract_measures(self) -> pd.DataFrame:
        return pd.DataFrame(self._measure_rows, columns=MEASURES_SCHEMA)

    def extract_sources(self) -> pd.DataFrame:
        dax_tables = self._get("dax_tables", ["TableName", "Expression"])
        power_query = self._get("power_query", ["TableName", "Expression"])
        dax_map = {
            _clean(row.get("TableName")): _clean(row.get("Expression"))
            for _, row in dax_tables.iterrows()
        }
        m_map = {
            _clean(row.get("TableName")): _clean(row.get("Expression"))
            for _, row in power_query.iterrows()
        }
        rows = []
        for table_name, _ in self.tables:
            t_num = self.table_number[table_name]
            if table_name in dax_map:
                expression = dax_map[table_name]
                rows.append({
                    "№ таблицы": t_num,
                    "Таблица": table_name,
                    "Тип источника": "Вычисляемая таблица (DAX)",
                    "Как создается": DAXFormatter.format_table(table_name, expression) or "—",
                    "Зависимости": self.dep_analyzer.find_dependencies(expression) if expression else "—",
                })
            elif table_name in m_map:
                expression = m_map[table_name]
                rows.append({
                    "№ таблицы": t_num,
                    "Таблица": table_name,
                    "Тип источника": "Power Query (M)",
                    "Как создается": SourceFormatter.format_source("m", expression) or "—",
                    "Зависимости": "—",
                })
            else:
                rows.append({
                    "№ таблицы": t_num,
                    "Таблица": table_name,
                    "Тип источника": "Не определён",
                    "Как создается": "—",
                    "Зависимости": "—",
                })
        return pd.DataFrame(rows, columns=SOURCES_SCHEMA)

    def extract_relationships(self) -> pd.DataFrame:
        frame = self._get(
            "relationships",
            ["FromTableName", "FromColumnName", "ToTableName", "ToColumnName", "IsActive"],
        )
        rows = []
        for idx, row in frame.iterrows():
            from_table = _clean(row.get("FromTableName"))
            to_table = _clean(row.get("ToTableName"))
            if not from_table or not to_table:
                continue
            from_num = self.table_number.get(from_table, "?")
            to_num = self.table_number.get(to_table, "?")
            is_active = _clean(row.get("IsActive")).lower() not in {"0", "false", ""}
            rows.append({
                "Название связи": f"Связь {idx + 1}",
                "Откуда": f"{from_table}[{_clean(row.get('FromColumnName'))}] ({from_num})",
                "Куда": f"{to_table}[{_clean(row.get('ToColumnName'))}] ({to_num})",
                "Состояние": "Активна" if is_active else "Не активна",
            })
        return pd.DataFrame(rows, columns=RELATIONSHIPS_SCHEMA)

    def extract_dax(self) -> pd.DataFrame:
        rows: list[dict] = []
        for _, row in self._get("dax_tables", ["TableName", "Expression"]).iterrows():
            table_name = _clean(row.get("TableName"))
            if table_name not in self.table_names:
                continue
            rows.append({
                "Тип объекта": "Вычисляемая таблица",
                "№ таблицы": self.table_number.get(table_name, "—"),
                "Таблица": table_name,
                "Имя": table_name,
                "Выражение": _clean(row.get("Expression")) or "—",
            })
        for _, row in self._get("dax_columns", ["TableName", "ColumnName", "Expression"]).iterrows():
            table_name = _clean(row.get("TableName"))
            if table_name not in self.table_names:
                continue
            rows.append({
                "Тип объекта": "Вычисляемый столбец",
                "№ таблицы": self.table_number.get(table_name, "—"),
                "Таблица": table_name,
                "Имя": _clean(row.get("ColumnName")),
                "Выражение": _clean(row.get("Expression")) or "—",
            })
        for measure in self._measure_rows:
            rows.append({
                "Тип объекта": "Мера",
                "№ таблицы": measure["№ таблицы"],
                "Таблица": measure["Таблица"],
                "Имя": measure["Мера"],
                "Выражение": measure["Формула (DAX)"],
            })
        return pd.DataFrame(rows, columns=DAX_SCHEMA)

    def extract_power_query(self) -> pd.DataFrame:
        frame = self._get("power_query", ["TableName", "Expression"])
        rows = []
        for _, row in frame.iterrows():
            table_name = _clean(row.get("TableName"))
            if table_name not in self.table_names:
                continue
            rows.append({"Таблица": table_name, "Запрос (M)": _clean(row.get("Expression")) or "—"})
        return pd.DataFrame(rows, columns=POWER_QUERY_SCHEMA)

    def extract_m_parameters(self) -> pd.DataFrame:
        frame = self._get("m_parameters", ["ParameterName", "Description", "Expression", "ModifiedTime"])
        rows = []
        for _, row in frame.iterrows():
            rows.append({
                "Параметр": _clean(row.get("ParameterName")),
                "Описание": _clean(row.get("Description")) or "—",
                "Выражение": _clean(row.get("Expression")) or "—",
                "Изменён": _clean(row.get("ModifiedTime")) or "—",
            })
        return pd.DataFrame(rows, columns=M_PARAMETERS_SCHEMA)

    def extract_rls(self) -> pd.DataFrame:
        frame = self._get("rls", ["RoleName", "TableName", "FilterExpression", "State", "RoleDescription"])
        rows = []
        for _, row in frame.iterrows():
            rows.append({
                "Роль": _clean(row.get("RoleName")),
                "Таблица": _clean(row.get("TableName")) or "—",
                "Фильтр": _clean(row.get("FilterExpression")) or "—",
                "Состояние": _clean(row.get("State")) or "—",
                "Описание": _clean(row.get("RoleDescription")) or "—",
            })
        return pd.DataFrame(rows, columns=RLS_SCHEMA)

    def extract_ols(self) -> pd.DataFrame:
        frame = self._get("ols", ["RoleName", "TableName", "ColumnName", "Scope", "Permission"])
        rows = []
        for _, row in frame.iterrows():
            rows.append({
                "Роль": _clean(row.get("RoleName")),
                "Таблица": _clean(row.get("TableName")) or "—",
                "Столбец": _clean(row.get("ColumnName")) or "—",
                "Область": _clean(row.get("Scope")) or "—",
                "Права": _clean(row.get("Permission")) or "—",
            })
        return pd.DataFrame(rows, columns=OLS_SCHEMA)

    def extract_perspectives(self) -> pd.DataFrame:
        frame = self._get(
            "perspectives",
            ["PerspectiveName", "ObjectType", "TableName", "ObjectName", "IncludeAll"],
        )
        rows = []
        for _, row in frame.iterrows():
            rows.append({
                "Перспектива": _clean(row.get("PerspectiveName")),
                "Тип объекта": _clean(row.get("ObjectType")) or "—",
                "Таблица": _clean(row.get("TableName")) or "—",
                "Объект": _clean(row.get("ObjectName")) or "—",
                "Включить все": _clean(row.get("IncludeAll")) or "—",
            })
        return pd.DataFrame(rows, columns=PERSPECTIVES_SCHEMA)

    def extract_statistics(self) -> pd.DataFrame:
        frame = self._get(
            "statistics",
            ["TableName", "ColumnName", "Cardinality", "Dictionary", "HashIndex", "DataSize"],
        )
        rows = []
        for _, row in frame.iterrows():
            table_name = _clean(row.get("TableName"))
            if table_name not in self.table_names:
                continue
            rows.append({
                "Таблица": table_name,
                "Столбец": _clean(row.get("ColumnName")),
                "Кардинальность": row.get("Cardinality"),
                "Словарь": row.get("Dictionary"),
                "HashIndex": row.get("HashIndex"),
                "DataSize": row.get("DataSize"),
            })
        return pd.DataFrame(rows, columns=STATISTICS_SCHEMA)

    def extract_all(self) -> dict[str, Any]:
        """Возвращает все наборы данных в ключах, ожидаемых lineage и Excel-отчётом."""
        dfs: dict[str, Any] = {
            "tables_ref": self.extract_tables_ref(),
            "columns": self.extract_columns(),
            "measures": self.extract_measures(),
            "sources": self.extract_sources(),
            "relationships": self.extract_relationships(),
            "dax": self.extract_dax(),
            "power_query": self.extract_power_query(),
            "m_parameters": self.extract_m_parameters(),
            "rls": self.extract_rls(),
            "ols": self.extract_ols(),
            "perspectives": self.extract_perspectives(),
            "statistics": self.extract_statistics(),
        }
        dfs["warnings"] = list(self.warnings)
        return dfs


def _to_int(value: Any) -> int | None:
    try:
        if value is None or pd.isna(value):
            return None
        return int(value)
    except (TypeError, ValueError):
        return None


def _sort_ordinal(value: Any) -> int:
    """Ключ сортировки по DisplayOrdinal: 0 — валидное значение, не пустое."""
    ordinal = _to_int(value)
    return 9999 if ordinal is None else ordinal


__all__ = ["PBIXModelExtractor", "COLUMN_ICON", "MEASURE_ICON", "TABLE_ICON"]
