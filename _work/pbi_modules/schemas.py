"""Единый источник правды о структуре отчёта.

Здесь описаны колонки каждого набора данных, наборы вкладок GUI и листов Excel
для обоих форматов. Раньше эти списки жили в трёх местах (extractors,
gui, excel_report) и расходились при любой правке.
"""
from __future__ import annotations

import pandas as pd

from .file_kind import PBIT, PBIX

# ---------------------------------------------------------------- колонки
SCHEMAS: dict[str, tuple[str, ...]] = {
    "tables_ref": ("№ таблицы", "Имя таблицы", "Описание"),
    "columns": (
        "№ столбца", "№ таблицы", "Таблица", "Столбец", "Описание",
        "Тип данных", "Тип значения", "Формат", "Сортировка по столбцу",
        "Агрегация по умолчанию", "Формула (если вычисляемый)", "Зависимости",
    ),
    "sources": ("№ таблицы", "Таблица", "Тип источника", "Как создается", "Зависимости"),
    "relationships": ("Название связи", "Откуда", "Куда", "Состояние"),
    "measures": (
        "№ меры", "№ таблицы", "Таблица", "Мера", "Описание", "Формула (DAX)",
        "Формат", "Папка отображения", "Зависимости",
    ),
    "sheets": ("№ листа", "Имя листа", "ID листа", "Видимость", "Количество визуалов", "Размеры"),
    "visuals": (
        "Имя листа", "ID визуала", "Тип визуала", "Роль", "Порядок", "Поле",
        "Отображаемое имя", "№ таблицы", "№ столбца/меры", "Расположение",
        "FilterTargetEntity", "FilterTargetProperty", "FilterExpressionShort",
    ),
    "conditional": ("Имя листа", "ID визуала", "Поле", "Элемент", "Правило", "Цвет"),
    "lineage": (
        "Тип источника", "ID источника", "Имя источника", "Тип приемника",
        "ID приемника", "Имя приемника", "Тип связи", "Контекст",
    ),
    "lineage_upstream": (
        "ID фокуса", "Тип фокуса", "Имя фокуса", "Тип источника", "ID источника",
        "Имя источника", "Тип приемника", "ID приемника", "Имя приемника",
        "Тип связи", "Контекст",
    ),
    "lineage_downstream": (
        "ID фокуса", "Тип фокуса", "Имя фокуса", "Тип источника", "ID источника",
        "Имя источника", "Тип приемника", "ID приемника", "Имя приемника",
        "Тип связи", "Контекст",
    ),
    "stats": ("Параметр", "Значение"),
    "dax": ("Тип объекта", "№ таблицы", "Таблица", "Имя", "Выражение"),
    "power_query": ("Таблица", "Запрос (M)"),
    "m_parameters": ("Параметр", "Описание", "Выражение", "Изменён"),
    "rls": ("Роль", "Таблица", "Фильтр", "Состояние", "Описание"),
    "ols": ("Роль", "Таблица", "Столбец", "Область", "Права"),
    "statistics": ("Таблица", "Столбец", "Кардинальность", "Словарь", "HashIndex", "DataSize"),
    # Метаданные модели PBIX, которые не выводятся в GUI/Excel, но доступны в API.
    "perspectives": ("Перспектива", "Тип объекта", "Таблица", "Объект", "Включить все"),
}

# ---------------------------------------------------------------- вкладки GUI
# PBIX не содержит слоя отчёта: визуальных вкладок нет, зато есть
# метаданные модели (M, RLS/OLS, статистика хранения).
TABS_BY_KIND: dict[str, tuple[tuple[str, str], ...]] = {
    PBIT: (
        ("tables_ref", "📋 Справочник таблиц"),
        ("columns", "📊 Столбцы"),
        ("sources", "🔌 Источники"),
        ("relationships", "🔗 Связи"),
        ("measures", "📐 Меры"),
        ("sheets", "📑 Листы"),
        ("visuals", "🎨 Визуальные элементы"),
        ("lineage", "🧬 Линейдж"),
        ("lineage_upstream", "🧭 Lineage Upstream"),
        ("lineage_downstream", "🧭 Lineage Downstream"),
        ("conditional", "🎨 Условное форматирование"),
        ("stats", "📈 Статистика"),
    ),
    PBIX: (
        ("tables_ref", "📋 Справочник таблиц"),
        ("columns", "📊 Столбцы"),
        ("sources", "🔌 Источники"),
        ("relationships", "🔗 Связи"),
        ("measures", "📐 Меры"),
        ("m_parameters", "🎛 Параметры M"),
        ("rls", "🔐 Безопасность (RLS)"),
        ("ols", "🔐 Безопасность (OLS)"),
        ("lineage", "🧬 Линейдж"),
        ("lineage_upstream", "🧭 Lineage Upstream"),
        ("lineage_downstream", "🧭 Lineage Downstream"),
        ("statistics", "💾 Статистика хранения"),
        ("stats", "📈 Статистика"),
    ),
}

# ---------------------------------------------------------------- листы Excel
SHEETS_BY_KIND: dict[str, tuple[tuple[str, str], ...]] = {
    PBIT: (
        ("Справочник таблиц", "tables_ref"),
        ("Столбцы", "columns"),
        ("Источники", "sources"),
        ("Связи", "relationships"),
        ("Меры", "measures"),
        ("Листы", "sheets"),
        ("Визуальные элементы", "visuals"),
        ("Линейдж", "lineage"),
        ("Lineage_Upstream", "lineage_upstream"),
        ("Lineage_Downstream", "lineage_downstream"),
        ("Условное форматирование", "conditional"),
        ("Статистика", "stats"),
    ),
    PBIX: (
        ("Справочник таблиц", "tables_ref"),
        ("Столбцы", "columns"),
        ("Источники", "sources"),
        ("Связи", "relationships"),
        ("Меры", "measures"),
        ("Параметры M", "m_parameters"),
        ("Безопасность (RLS)", "rls"),
        ("Безопасность (OLS)", "ols"),
        ("Линейдж", "lineage"),
        ("Lineage_Upstream", "lineage_upstream"),
        ("Lineage_Downstream", "lineage_downstream"),
        ("Статистика хранения", "statistics"),
        ("Статистика", "stats"),
    ),
}


def _kind(file_kind: str | None) -> str:
    value = str(file_kind or "").lower()
    return PBIX if value == PBIX else PBIT


def tabs_for(file_kind: str | None) -> tuple[tuple[str, str], ...]:
    """Набор вкладок GUI по типу файла (по умолчанию — PBIT)."""
    return TABS_BY_KIND[_kind(file_kind)]


def sheets_for(file_kind: str | None) -> tuple[tuple[str, str], ...]:
    """Набор листов Excel по типу файла (по умолчанию — PBIT)."""
    return SHEETS_BY_KIND[_kind(file_kind)]


def columns_for(df_key: str) -> tuple[str, ...]:
    """Колонки набора данных или пустой кортеж для неизвестного ключа."""
    return SCHEMAS.get(df_key, ())


def frame_for(rows: list[dict], df_key: str) -> pd.DataFrame:
    """DataFrame со строгим порядком колонок из реестра.

    Пустой набор получает колонки схемы, поэтому лист Excel не теряет заголовки.
    """
    frame = pd.DataFrame(rows)
    columns = columns_for(df_key)
    if not columns:
        return frame
    return frame.reindex(columns=list(columns))


__all__ = [
    "SCHEMAS",
    "TABS_BY_KIND",
    "SHEETS_BY_KIND",
    "tabs_for",
    "sheets_for",
    "columns_for",
    "frame_for",
]
