import pandas as pd
import pytest
from pbi_modules.lineage import build_lineage_edges
from pbi_modules.pbix_extractor import PBIXModelExtractor


class FakePBIXModel:
    """Минимальная замена PBIXRay: только атрибуты, которые читает адаптер."""

    def __init__(self):
        self.tables = ["Sales", "Products", "DateTableTemplate_1", "LocalDateTable_1"]
        self.tmschema_tables = pd.DataFrame([
            {"Name": "Sales", "Description": "", "IsHidden": 0},
            {"Name": "Products", "Description": "Справочник", "IsHidden": 0},
            {"Name": "HiddenTable", "Description": "", "IsHidden": 1},
        ])
        self.tmschema_columns = pd.DataFrame([
            {"TableName": "Sales", "Name": "Amount", "Type": 1, "DataType": 8,
             "Description": "", "Expression": None, "FormatString": "#,0.00",
             "SummarizeBy": 1, "DisplayOrdinal": 0},
            {"TableName": "Sales", "Name": "Total", "Type": 2, "DataType": 10,
             "Description": "Итог", "Expression": "SUM(Sales[Amount]) * 2",
             "FormatString": None, "SummarizeBy": 0, "DisplayOrdinal": 1},
            {"TableName": "Products", "Name": "ProdKey", "Type": 1, "DataType": 6,
             "Description": "", "Expression": None, "FormatString": "0",
             "SummarizeBy": 0, "DisplayOrdinal": 0},
            {"TableName": "Products", "Name": "Name", "Type": 1, "DataType": 2,
             "Description": "", "Expression": None, "FormatString": None,
             "SummarizeBy": 0, "DisplayOrdinal": 1},
            {"TableName": "HiddenTable", "Name": "Secret", "Type": 1, "DataType": 2,
             "Description": "", "Expression": None, "FormatString": None,
             "SummarizeBy": 0, "DisplayOrdinal": 0},
        ])
        self.dax_measures = pd.DataFrame([
            {"TableName": "Sales", "Name": "Total Sales",
             "Expression": "SUM(Sales[Amount]) + [Total Cost]",
             "DisplayFolder": "Продажи", "Description": "Сумма"},
            {"TableName": "Products", "Name": "Total Cost",
             "Expression": "SUMX(Products, 1)", "DisplayFolder": None, "Description": None},
            {"TableName": "HiddenTable", "Name": "Hidden Sales",
             "Expression": "SUMX(Products, [Total Cost])", "DisplayFolder": None, "Description": None},
        ])
        self.dax_tables = pd.DataFrame([{"TableName": "Products", "Expression": "DISTINCT(Sales[Amount])"}])
        self.dax_columns = pd.DataFrame([
            {"TableName": "Sales", "ColumnName": "Total", "Expression": "SUM(Sales[Amount]) * 2"},
        ])
        self.power_query = pd.DataFrame([{"TableName": "Sales", "Expression": "let Source = 1 in Source"}])
        self.m_parameters = pd.DataFrame([
            {"ParameterName": "Server", "Description": "Сервер", "Expression": '"srv"', "ModifiedTime": ""},
        ])
        self.relationships = pd.DataFrame([
            {"FromTableName": "Sales", "FromColumnName": "Amount",
             "ToTableName": "Products", "ToColumnName": "ProdKey", "IsActive": 1},
        ])
        self.rls = pd.DataFrame([
            {"RoleName": "Читатель", "TableName": "Sales", "FilterExpression": "[Amount] > 0",
             "State": "Enabled", "RoleDescription": ""},
        ])
        self.ols = pd.DataFrame([
            {"RoleName": "Читатель", "TableName": "Products", "ColumnName": "Name",
             "Scope": "Column", "Permission": "None"},
        ])
        self.perspectives = pd.DataFrame([
            {"PerspectiveName": "SalesView", "ObjectType": "Table", "TableName": "Sales",
             "ObjectName": "Sales", "IncludeAll": 1},
        ])
        self.statistics = pd.DataFrame([
            {"TableName": "Sales", "ColumnName": "Amount", "Cardinality": 10,
             "Dictionary": 100, "HashIndex": 10, "DataSize": 80},
            {"TableName": "HiddenTable", "ColumnName": "Secret", "Cardinality": 3,
             "Dictionary": 20, "HashIndex": 5, "DataSize": 30},
        ])
        self.size = 1234

    def close(self):
        pass


@pytest.fixture()
def dfs():
    extractor = PBIXModelExtractor(FakePBIXModel())
    return extractor.extract_all()


def test_tables_skip_hidden_and_system_objects(dfs):
    names = dfs["tables_ref"]["Имя таблицы"].tolist()

    assert names == ["Sales", "Products"]
    assert dfs["tables_ref"]["№ таблицы"].tolist() == ["s1", "s2"]
    assert dfs["tables_ref"].iloc[1]["Описание"] == "Справочник"


def test_columns_get_pbit_style_ids_and_types(dfs):
    columns = dfs["columns"].set_index("Столбец")

    assert columns.loc["Amount", "№ столбца"] == "A101"
    assert columns.loc["Total", "№ столбца"] == "A102"
    assert columns.loc["Total", "Тип значения"] == "Вычисляемый столбец"
    assert columns.loc["ProdKey", "№ столбца"] == "A201"
    assert columns.loc["Amount", "Тип данных"] == "double"
    assert columns.loc["Amount", "Агрегация по умолчанию"] == "Sum"
    assert columns.loc["Total", "Формула (если вычисляемый)"].startswith("Total =")


def test_measure_dependencies_use_ref_icons(dfs):
    measures = dfs["measures"].set_index("Мера")
    deps = measures.loc["Total Sales", "Зависимости"]

    assert measures.loc["Total Sales", "№ меры"] == "M1"
    assert "📊 S1" in deps
    assert "📋 A101" in deps
    assert "📐 M2" in deps


def test_measures_of_hidden_tables_are_kept(dfs):
    measures = dfs["measures"].set_index("Мера")

    assert "Hidden Sales" in measures.index
    assert measures.loc["Hidden Sales", "№ меры"] == "M3"
    assert measures.loc["Hidden Sales", "Таблица"] == "HiddenTable"
    # У скрытой таблицы нет номера в справочнике — прочерк, а не выдуманный ID.
    assert measures.loc["Hidden Sales", "№ таблицы"] == "—"
    assert "📐 M2" in measures.loc["Hidden Sales", "Зависимости"]


def test_hidden_table_measure_has_no_fake_lineage_node(dfs):
    edges = build_lineage_edges(dfs)

    assert (edges["ID приемника"] == "M3").any()
    assert not (edges["ID источника"] == "—").any()
    assert not (edges["ID приемника"] == "—").any()


def test_calculated_column_dependencies(dfs):
    column = dfs["columns"].set_index("Столбец").loc["Total"]

    assert "📊 S1" in column["Зависимости"]
    assert "📋 A101" in column["Зависимости"]


def test_sources_detect_m_and_dax_tables(dfs):
    sources = dfs["sources"].set_index("Таблица")

    assert sources.loc["Sales", "Тип источника"] == "Power Query (M)"
    assert sources.loc["Products", "Тип источника"] == "Вычисляемая таблица (DAX)"


def test_relationships_and_statistics_are_adapted(dfs):
    relationship = dfs["relationships"].iloc[0]
    assert relationship["Откуда"] == "Sales[Amount] (s1)"
    assert relationship["Куда"] == "Products[ProdKey] (s2)"
    assert relationship["Состояние"] == "Активна"

    statistics = dfs["statistics"]
    assert statistics["Таблица"].tolist() == ["Sales"]


def test_model_lineage_links_columns_and_measures(dfs):
    edges = build_lineage_edges(dfs)
    dependencies = edges[edges["Тип связи"] == "Зависит от"]

    assert (dependencies["ID источника"] == "A101").any()
    assert (dependencies["ID приемника"] == "M1").any()
    assert (dependencies["ID источника"] == "M2").any()
    assert "Визуал" not in set(edges["Тип приемника"])


def test_new_frames_have_stable_columns(dfs):
    assert list(dfs["rls"].columns) == ["Роль", "Таблица", "Фильтр", "Состояние", "Описание"]
    assert dfs["rls"].iloc[0]["Роль"] == "Читатель"
    assert dfs["ols"].iloc[0]["Права"] == "None"
    assert dfs["perspectives"].iloc[0]["Перспектива"] == "SalesView"
    assert dfs["m_parameters"].iloc[0]["Параметр"] == "Server"
    assert dfs["power_query"].iloc[0]["Таблица"] == "Sales"
    assert list(dfs["dax"].columns) == ["Тип объекта", "№ таблицы", "Таблица", "Имя", "Выражение"]
    assert dfs["warnings"] == []


class FakePBIXModelWithCalculatedTable(FakePBIXModel):
    """Модель с вычисляемой таблицей.

    PBIXRay не отдаёт её в tmschema_tables/tmschema_columns: строка таблицы
    помечена SystemFlags = 2, а столбцы имеют Column.Type = 4. Список таблиц
    приходит из model.tables (schema), сама таблица — из dax_tables.
    """

    def __init__(self):
        super().__init__()
        self.tables = ["Sales", "Calendar", "H$ Sales (1)$Amount (2)"]
        self.dax_tables = pd.DataFrame([
            {"TableName": "Calendar",
             "Expression": "ADDCOLUMNS(DISTINCT(Sales[Amount]), \"Year\", YEAR([Amount]))"},
        ])
        self.schema = pd.DataFrame([
            {"TableName": "Sales", "ColumnName": "Amount", "PandasDataType": "float64"},
            {"TableName": "Sales", "ColumnName": "Total", "PandasDataType": "Float64"},
            {"TableName": "Calendar", "ColumnName": "Date", "PandasDataType": "datetime64[ns]"},
            {"TableName": "Calendar", "ColumnName": "Year", "PandasDataType": "Int64"},
            {"TableName": "Calendar", "ColumnName": "IsCurrent", "PandasDataType": "boolean"},
            {"TableName": "Calendar", "ColumnName": "Month", "PandasDataType": "string"},
        ])
        self.statistics = pd.DataFrame([
            {"TableName": "Calendar", "ColumnName": "Date", "Cardinality": 366,
             "Dictionary": 400, "HashIndex": 0, "DataSize": 1200},
        ])


def test_calculated_table_is_kept_out_of_tmschema():
    dfs = PBIXModelExtractor(FakePBIXModelWithCalculatedTable()).extract_all()
    tables = dfs["tables_ref"].set_index("Имя таблицы")

    # Служебная таблица движка (H$) отброшена, вычисляемая — сохранена.
    assert "H$ Sales (1)$Amount (2)" not in tables.index
    assert list(tables.index) == ["Sales", "Calendar", "Products"]
    assert tables.loc["Calendar", "№ таблицы"] == "s2"
    assert dfs["warnings"] == []


def test_calculated_table_columns_come_from_schema():
    dfs = PBIXModelExtractor(FakePBIXModelWithCalculatedTable()).extract_all()
    columns = dfs["columns"][dfs["columns"]["Таблица"] == "Calendar"].set_index("Столбец")

    assert list(columns.index) == ["Date", "Year", "IsCurrent", "Month"]
    assert columns.loc["Date", "№ столбца"] == "A201"
    assert columns.loc["Date", "Тип данных"] == "dateTime"
    assert columns.loc["Year", "Тип данных"] == "int64"
    assert columns.loc["IsCurrent", "Тип данных"] == "boolean"
    assert columns.loc["Month", "Тип данных"] == "string"
    assert set(columns["Тип значения"]) == {"Столбец вычисляемой таблицы"}
    assert set(columns["Формула (если вычисляемый)"]) == {"—"}


def test_calculated_table_is_visible_in_sources_dax_and_statistics():
    dfs = PBIXModelExtractor(FakePBIXModelWithCalculatedTable()).extract_all()

    sources = dfs["sources"].set_index("Таблица")
    assert sources.loc["Calendar", "Тип источника"] == "Вычисляемая таблица (DAX)"
    assert "📊 S1" in sources.loc["Calendar", "Зависимости"]

    dax_tables = dfs["dax"][dfs["dax"]["Тип объекта"] == "Вычисляемая таблица"]
    assert dax_tables["Таблица"].tolist() == ["Calendar"]
    assert "ADDCOLUMNS" in dax_tables.iloc[0]["Выражение"]

    assert "Calendar" in dfs["statistics"]["Таблица"].tolist()


def test_calculated_column_type_falls_back_to_schema():
    """У вычисляемого столбца DataType = Automatic (1) — тип берём из schema."""

    class AutomaticTypeModel(FakePBIXModel):
        def __init__(self):
            super().__init__()
            self.tmschema_columns.loc[
                self.tmschema_columns["Name"] == "Total", "DataType"
            ] = 1
            self.schema = pd.DataFrame([
                {"TableName": "Sales", "ColumnName": "Amount", "PandasDataType": "float64"},
                {"TableName": "Sales", "ColumnName": "Total", "PandasDataType": "Int64"},
            ])

    dfs = PBIXModelExtractor(AutomaticTypeModel()).extract_all()
    column = dfs["columns"].set_index("Столбец").loc["Total"]

    assert column["Тип данных"] == "int64"
    assert column["Тип значения"] == "Вычисляемый столбец"


def test_calculated_table_without_storage_uses_statistics_columns():
    class UntouchedModel(FakePBIXModelWithCalculatedTable):
        """Таблица ещё ни разу не рассчитывалась: schema по ней пуста."""

        def __init__(self):
            super().__init__()
            self.schema = pd.DataFrame(columns=["TableName", "ColumnName", "PandasDataType"])
            self.statistics = pd.DataFrame([
                {"TableName": "Calendar", "ColumnName": "Date", "Cardinality": 0,
                 "Dictionary": 0, "HashIndex": 0, "DataSize": 0},
                {"TableName": "Calendar", "ColumnName": "Year", "Cardinality": 0,
                 "Dictionary": 0, "HashIndex": 0, "DataSize": 0},
            ])

    dfs = PBIXModelExtractor(UntouchedModel()).extract_all()
    columns = dfs["columns"][dfs["columns"]["Таблица"] == "Calendar"]

    assert columns["Столбец"].tolist() == ["Date", "Year"]
    assert set(columns["Тип значения"]) == {"Столбец вычисляемой таблицы"}
    assert set(columns["Тип данных"]) == {"—"}


def test_partial_failure_keeps_other_blocks():
    class BrokenModel:
        """Прокси, у которого падает только чтение RLS."""

        def __init__(self):
            self._inner = FakePBIXModel()

        def __getattr__(self, name):
            if name == "rls":
                raise RuntimeError("RLS не читается")
            return getattr(self._inner, name)

    extractor = PBIXModelExtractor(BrokenModel())
    dfs = extractor.extract_all()

    assert not dfs["tables_ref"].empty
    assert dfs["rls"].empty
    assert list(dfs["rls"].columns) == ["Роль", "Таблица", "Фильтр", "Состояние", "Описание"]
    assert any("rls" in warning for warning in dfs["warnings"])
