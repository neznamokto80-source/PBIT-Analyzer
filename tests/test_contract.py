"""Golden-тесты контракта: то, что нельзя менять при рефакторинге.

Фиксируют внешнее поведение «как есть»: наборы вкладок GUI, листы Excel,
заголовки колонок и ключи результата анализа. Любое расхождение — изменение
формата отчёта, а не рефакторинг.
"""
import json
import zipfile

import pytest
from pbi_modules.excel_report import ExcelReportGenerator
from pbi_modules.file_kind import PBIT, PBIX
from pbi_modules.gui import tab_mapping_for

EXPECTED_TABS_PBIT = [
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
]

EXPECTED_TABS_PBIX = [
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
]

EXPECTED_SHEETS_PBIT = [
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
]

EXPECTED_SHEETS_PBIX = [
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
]

EXPECTED_COLUMNS = {
    "tables_ref": ["№ таблицы", "Имя таблицы", "Описание"],
    "columns": [
        "№ столбца", "№ таблицы", "Таблица", "Столбец", "Описание",
        "Тип данных", "Тип значения", "Формат", "Сортировка по столбцу",
        "Агрегация по умолчанию", "Формула (если вычисляемый)", "Зависимости",
    ],
    "sources": ["№ таблицы", "Таблица", "Тип источника", "Как создается", "Зависимости"],
    "relationships": ["Название связи", "Откуда", "Куда", "Состояние"],
    "measures": [
        "№ меры", "№ таблицы", "Таблица", "Мера", "Описание", "Формула (DAX)",
        "Формат", "Папка отображения", "Зависимости",
    ],
    "sheets": ["№ листа", "Имя листа", "ID листа", "Видимость", "Количество визуалов", "Размеры"],
    "visuals": [
        "Имя листа", "ID визуала", "Тип визуала", "Роль", "Порядок", "Поле",
        "Отображаемое имя", "№ таблицы", "№ столбца/меры", "Расположение",
        "FilterTargetEntity", "FilterTargetProperty", "FilterExpressionShort",
    ],
    "conditional": ["Имя листа", "ID визуала", "Поле", "Элемент", "Правило", "Цвет"],
    "lineage": [
        "Тип источника", "ID источника", "Имя источника", "Тип приемника",
        "ID приемника", "Имя приемника", "Тип связи", "Контекст",
    ],
    "lineage_upstream": [
        "ID фокуса", "Тип фокуса", "Имя фокуса", "Тип источника", "ID источника",
        "Имя источника", "Тип приемника", "ID приемника", "Имя приемника",
        "Тип связи", "Контекст",
    ],
    "lineage_downstream": [
        "ID фокуса", "Тип фокуса", "Имя фокуса", "Тип источника", "ID источника",
        "Имя источника", "Тип приемника", "ID приемника", "Имя приемника",
        "Тип связи", "Контекст",
    ],
    "stats": ["Параметр", "Значение"],
    "dax": ["Тип объекта", "№ таблицы", "Таблица", "Имя", "Выражение"],
    "power_query": ["Таблица", "Запрос (M)"],
    "m_parameters": ["Параметр", "Описание", "Выражение", "Изменён"],
    "rls": ["Роль", "Таблица", "Фильтр", "Состояние", "Описание"],
    "ols": ["Роль", "Таблица", "Столбец", "Область", "Права"],
    "statistics": ["Таблица", "Столбец", "Кардинальность", "Словарь", "HashIndex", "DataSize"],
}

EXPECTED_RESULT_KEYS = {
    "tables_ref", "columns", "sources", "relationships", "measures", "sheets",
    "visuals", "conditional", "lineage", "lineage_upstream", "lineage_downstream",
    "stats", "excel_path", "file_kind",
}


def test_pbit_tabs_contract():
    assert [(key, name) for key, name in tab_mapping_for(PBIT)] == EXPECTED_TABS_PBIT


def test_pbix_tabs_contract():
    assert [(key, name) for key, name in tab_mapping_for(PBIX)] == EXPECTED_TABS_PBIX


def test_pbit_excel_sheets_contract():
    assert ExcelReportGenerator.sheets_for(PBIT) == EXPECTED_SHEETS_PBIT


def test_pbix_excel_sheets_contract():
    assert ExcelReportGenerator.sheets_for(PBIX) == EXPECTED_SHEETS_PBIX


@pytest.mark.parametrize("df_key", sorted(EXPECTED_COLUMNS))
def test_excel_columns_contract(df_key):
    generator = ExcelReportGenerator("out.xlsx", {}, "")
    assert list(generator._get_columns(df_key)) == EXPECTED_COLUMNS[df_key]


@pytest.mark.parametrize("file_kind", [PBIT, PBIX])
def test_every_tab_and_sheet_has_columns(file_kind):
    """Каждый лист/вкладка обязаны иметь описанные колонки — иначе пустой лист."""
    generator = ExcelReportGenerator("out.xlsx", {}, "")
    for sheet_name, key in ExcelReportGenerator.sheets_for(file_kind):
        assert generator._get_columns(key), f"нет схемы колонок для листа «{sheet_name}» ({key})"


def _write_pbit(path):
    model = {
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "columns": [{"name": "Amount", "dataType": "double"}],
                    "measures": [{"name": "Total Sales", "expression": "SUM(Sales[Amount])"}],
                    "partitions": [{"source": {"type": "m", "expression": "let x = 1 in x"}}],
                }
            ],
            "relationships": [],
        }
    }
    layout = {"sections": [{"displayName": "Page 1", "visualContainers": []}]}
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("DataModelSchema", json.dumps(model).encode("utf-16-le"))
        archive.writestr("Layout", json.dumps(layout).encode("utf-16-le"))
    return path


def test_analysis_result_keys_contract(tmp_path):
    from pbi_modules.dispatcher import analyze_file

    pbit = _write_pbit(tmp_path / "contract.pbit")
    result = analyze_file(str(pbit), str(tmp_path), create_excel=False)

    assert set(result) == EXPECTED_RESULT_KEYS
