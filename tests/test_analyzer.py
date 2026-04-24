import json
import zipfile

from pbi_modules.analyzer import analyze_pbit_file, build_report_filename


def test_analyze_pbit_file_generates_excel(tmp_path):
    model_payload = {
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
    layout_payload = {"sections": [{"displayName": "Page 1", "visualContainers": []}]}

    pbit_path = tmp_path / "sample.pbit"
    with zipfile.ZipFile(pbit_path, "w") as archive:
        archive.writestr("DataModelSchema", json.dumps(model_payload).encode("utf-16-le"))
        archive.writestr("Layout", json.dumps(layout_payload).encode("utf-16-le"))

    result = analyze_pbit_file(str(pbit_path), str(tmp_path), keep_temp=False)

    assert "excel_path" in result
    assert result["tables_ref"].iloc[0]["Имя таблицы"] == "Sales"
    assert result["excel_path"].endswith(".xlsx")
    assert "sample" in result["excel_path"]
    assert "lineage" in result
    assert not result["lineage"].empty
    assert "lineage_upstream" in result
    assert "lineage_downstream" in result
    assert "Тип источника" in result["lineage"].columns
    assert "ID фокуса" in result["lineage_upstream"].columns
    assert "ID фокуса" in result["lineage_downstream"].columns


def test_build_report_filename_includes_source_name():
    filename = build_report_filename(r"D:\reports\sales model.pbit", "20260416_101010")

    assert filename == "pbi_report_sales model_20260416_101010.xlsx"


def test_analyze_pbit_file_applies_lineage_focus(tmp_path):
    model_payload = {
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
    layout_payload = {"sections": [{"displayName": "Page 1", "visualContainers": []}]}

    pbit_path = tmp_path / "sample_focus.pbit"
    with zipfile.ZipFile(pbit_path, "w") as archive:
        archive.writestr("DataModelSchema", json.dumps(model_payload).encode("utf-16-le"))
        archive.writestr("Layout", json.dumps(layout_payload).encode("utf-16-le"))

    result = analyze_pbit_file(
        str(pbit_path),
        str(tmp_path),
        keep_temp=False,
        lineage_focus="M1",
    )

    assert not result["lineage_upstream"].empty
    assert set(result["lineage_upstream"]["ID фокуса"].unique()) == {"M1"}
    if not result["lineage_downstream"].empty:
        assert set(result["lineage_downstream"]["ID фокуса"].unique()) == {"M1"}
