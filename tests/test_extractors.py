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
