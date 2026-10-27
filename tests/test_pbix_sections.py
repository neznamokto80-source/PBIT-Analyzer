"""Разделы PBIX: набор вкладок GUI и листов Excel."""
from pbi_modules.excel_report import ExcelReportGenerator
from pbi_modules.gui import TAB_MAPPING_PBIX

PBIT_SHEETS = [
    "tables_ref",
    "columns",
    "sources",
    "relationships",
    "measures",
    "sheets",
    "visuals",
    "lineage",
    "lineage_upstream",
    "lineage_downstream",
    "conditional",
    "stats",
]


def test_pbix_tabs_have_no_perspectives():
    keys = [key for key, _ in TAB_MAPPING_PBIX]

    assert "perspectives" not in keys
    assert "measures" in keys


def test_pbix_excel_sheets_have_no_perspectives():
    sheets = ExcelReportGenerator.sheets_for("pbix")
    keys = [key for _, key in sheets]

    assert "perspectives" not in keys
    assert "measures" in keys
    assert "Перспективы" not in [name for name, _ in sheets]


def test_pbit_excel_sheets_unchanged():
    assert [key for _, key in ExcelReportGenerator.sheets_for("pbit")] == PBIT_SHEETS
