import json

from pbi_modules.extractors import DataModelExtractor, LayoutExtractor, load_json_utf16


def test_load_json_utf16_fallback_to_utf8(tmp_path):
    file_path = tmp_path / "layout.json"
    payload = {"sections": [{"name": "Page 1"}]}
    file_path.write_text(json.dumps(payload), encoding="utf-8")

    loaded = load_json_utf16(str(file_path))

    assert loaded["sections"][0]["name"] == "Page 1"


def test_extract_tables_and_measures_minimal_model():
    model_data = {
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "columns": [{"name": "Amount", "dataType": "double"}],
                    "measures": [{"name": "Total Sales", "expression": "SUM(Sales[Amount])"}],
                }
            ],
            "relationships": [],
        }
    }

    extractor = DataModelExtractor(model_data)
    tables = extractor.extract_tables_ref()
    measures = extractor.extract_measures()

    assert not tables.empty
    assert tables.iloc[0]["Имя таблицы"] == "Sales"
    assert not measures.empty
    assert measures.iloc[0]["Мера"] == "Total Sales"


def test_extract_visuals_adds_icons_for_table_and_measure_refs():
    model_data = {
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "columns": [{"name": "Amount", "dataType": "double"}],
                    "measures": [{"name": "Total Sales", "expression": "SUM(Sales[Amount])"}],
                }
            ],
            "relationships": [],
        }
    }
    layout_data = {
        "sections": [
            {
                "displayName": "Page 1",
                "visualContainers": [
                    {
                        "x": 1,
                        "y": 2,
                        "width": 300,
                        "height": 200,
                        "config": json.dumps(
                            {
                                "name": "visual_1",
                                "singleVisual": {
                                    "visualType": "card",
                                    "projections": {"Values": [{"queryRef": "Total Sales"}]},
                                    "prototypeQuery": {
                                        "From": [{"Name": "s", "Entity": "Sales"}],
                                        "Select": [
                                            {
                                                "Name": "Total Sales",
                                                "Measure": {
                                                    "Property": "Total Sales",
                                                    "Expression": {"SourceRef": {"Source": "s"}},
                                                },
                                            }
                                        ],
                                    },
                                },
                            }
                        ),
                    }
                ],
            }
        ]
    }

    extractor = DataModelExtractor(model_data)
    visuals = LayoutExtractor(layout_data, extractor).extract_visuals()

    assert visuals.iloc[0]["№ таблицы"] == "📊 S1"
    assert visuals.iloc[0]["№ столбца/меры"] == "📐 M1"
    assert visuals.iloc[0]["ID визуала"] == "VIS_ID :visual_1"


def test_measure_of_hidden_table_is_kept():
    model_data = {
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "columns": [{"name": "Amount", "dataType": "double"}],
                    "measures": [{"name": "Visible M", "expression": "SUM(Sales[Amount])"}],
                },
                {
                    "name": "Hidden",
                    "isHidden": True,
                    "columns": [{"name": "Code", "dataType": "string"}],
                    "measures": [{"name": "Hidden M", "expression": "COUNTROWS(Hidden)"}],
                },
            ],
            "relationships": [],
        }
    }

    extractor = DataModelExtractor(model_data)
    measures = extractor.extract_measures()
    by_name = {row["Мера"]: row for _, row in measures.iterrows()}

    assert set(by_name) == {"Visible M", "Hidden M"}
    # Нумерация видимых мер не изменилась, скрытая получила следующий номер.
    assert by_name["Visible M"]["№ меры"] == "M1"
    assert by_name["Hidden M"]["№ меры"] == "M2"
    # У скрытой таблицы нет номера в справочнике — прочерк, а не выдуманный ID.
    assert by_name["Hidden M"]["№ таблицы"] == "—"
    assert by_name["Visible M"]["№ таблицы"] == "s1"


def test_hidden_table_measure_reaches_lineage():
    from pbi_modules.lineage import build_lineage_edges

    model_data = {
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "columns": [{"name": "Amount", "dataType": "double"}],
                    "measures": [{"name": "Visible M", "expression": "SUM(Sales[Amount])"}],
                },
                {
                    "name": "Hidden",
                    "isHidden": True,
                    "measures": [{"name": "Hidden M", "expression": "[Visible M] * 2"}],
                },
            ],
            "relationships": [],
        }
    }

    extractor = DataModelExtractor(model_data)
    dfs = {
        "tables_ref": extractor.extract_tables_ref(),
        "columns": extractor.extract_columns(),
        "measures": extractor.extract_measures(),
        "sources": extractor.extract_sources(),
    }

    edges = build_lineage_edges(dfs)

    assert "Hidden M" in set(edges["Имя приемника"])
    assert "M2" in set(edges["ID приемника"])


def test_system_date_tables_are_filtered():
    model_data = {
        "model": {
            "tables": [
                {"name": "Sales", "columns": [], "measures": []},
                {"name": "LocalDateTable_abc", "columns": [], "measures": []},
                {"name": "DateTableTemplate_xyz", "columns": [], "measures": []},
            ],
            "relationships": [],
        }
    }

    extractor = DataModelExtractor(model_data)

    assert list(extractor.extract_tables_ref()["Имя таблицы"]) == ["Sales"]


def test_table_without_name_does_not_crash():
    model_data = {
        "model": {
            "tables": [
                {"columns": [{"name": "X"}]},
                {"name": "Sales", "columns": [], "measures": []},
            ],
            "relationships": [],
        }
    }

    extractor = DataModelExtractor(model_data)

    assert list(extractor.extract_tables_ref()["Имя таблицы"]) == ["Sales"]


def test_dependencies_use_ref_type_icons():
    model_data = {
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "columns": [{"name": "Amount", "dataType": "double"}],
                    "measures": [{"name": "Total Sales", "expression": "SUM(Sales[Amount])"}],
                }
            ],
            "relationships": [],
        }
    }
    extractor = DataModelExtractor(model_data)
    measures = extractor.extract_measures()
    deps = measures.iloc[0]["Зависимости"]

    assert "📊" in deps
    assert "📋" in deps
