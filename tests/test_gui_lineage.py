"""Сценарии фильтра lineage, перенесённые из черновиков _work/test_*.py.

Проверяют устойчивость фильтра к «смешанному» результату анализа: в словаре
лежат не только DataFrame, но и строки (excel_path) и None (columns).
"""
import os

import pandas as pd
import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pbi_modules.gui import PBITAnalyzerMainWindow  # noqa: E402
from PyQt6.QtWidgets import QApplication  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def window(qapp):
    win = PBITAnalyzerMainWindow()
    yield win
    win.close()


def _dfs_with_strings():
    return {
        "measures": pd.DataFrame(
            {
                "№ меры": ["M1", "M2", "M3"],
                "Мера": ["Мера1", "Мера2", "Мера3"],
                "№ таблицы": ["s1", "s2", "s3"],
            }
        ),
        "tables_ref": pd.DataFrame(
            {
                "№ таблицы": ["s1", "s2", "s3"],
                "Имя таблицы": ["Таблица1", "Таблица2", "Таблица3"],
            }
        ),
        "columns": None,
        "sources": pd.DataFrame(),
        "lineage": pd.DataFrame(
            {
                "Тип источника": ["Мера", "Мера", "Таблица"],
                "ID источника": ["M1", "M2", "s1"],
                "Имя источника": ["Мера1", "Мера2", "Таблица1"],
                "Тип приемника": ["Визуал", "Визуал", "Столбец"],
                "ID приемника": ["VIS:1", "VIS:2", "A1"],
                "Имя приемника": ["Визуал1", "Визуал2", "Столбец1"],
                "Тип связи": ["использует", "использует", "содержит"],
                "Контекст": ["—", "—", "—"],
            }
        ),
        "visuals": pd.DataFrame(
            {
                "ID визуала": ["VIS_ID :1", "VIS_ID :2", "VIS_ID :3"],
                "Тип визуала": ["Карточка", "График", "Таблица"],
                "Имя листа": ["Лист1", "Лист2", "Лист3"],
            }
        ),
        "excel_path": "/path/to/excel.xlsx",
    }


def test_focus_filter_tolerates_non_dataframe_values(window):
    dfs = _dfs_with_strings()
    window.current_dfs = dfs
    window._original_dfs = dfs.copy()

    window.lineage_focus_combo.addItem("Мера | M2 | Мера2", userData="M2")
    window.lineage_focus_combo.setCurrentIndex(0)

    window._apply_lineage_focus()

    assert window.current_dfs["excel_path"] == "/path/to/excel.xlsx"
    assert len(window.current_dfs["measures"]) == 1
    assert window.current_dfs["measures"].iloc[0]["Мера"] == "Мера2"
    # Визуал VIS:2 связан с мерой M2 через lineage — остаётся только он.
    assert list(window.current_dfs["visuals"]["ID визуала"]) == ["VIS_ID :2"]
    # У M2 есть только исходящее ребро M2 → VIS:2, входящих нет.
    assert not window.current_dfs["lineage_downstream"].empty
    assert window.current_dfs["lineage_upstream"].empty


def test_filter_dfs_by_measure_without_gui():
    from pbi_modules.ui.lineage_panel import filter_dfs

    dfs = _dfs_with_strings()
    filtered = filter_dfs(dfs, "M2")

    assert filtered["excel_path"] == "/path/to/excel.xlsx"
    assert list(filtered["measures"]["Мера"]) == ["Мера2"]
    assert list(filtered["visuals"]["ID визуала"]) == ["VIS_ID :2"]
    assert len(filtered["lineage"]) == 1
    assert not filtered["lineage_downstream"].empty
    assert filtered["lineage_upstream"].empty


def test_filter_dfs_by_table():
    from pbi_modules.ui.lineage_panel import filter_dfs

    dfs = _dfs_with_strings()
    filtered = filter_dfs(dfs, "S1")

    assert list(filtered["tables_ref"]["Имя таблицы"]) == ["Таблица1"]
    assert filtered["measures"].empty
    assert len(filtered["lineage"]) == 1


def test_filter_dfs_unknown_focus_returns_empty_frames():
    from pbi_modules.ui.lineage_panel import filter_dfs

    dfs = _dfs_with_strings()
    filtered = filter_dfs(dfs, "M99")

    assert filtered["measures"].empty
    assert filtered["visuals"].empty
    assert filtered["lineage"].empty
    assert filtered["excel_path"] == "/path/to/excel.xlsx"


def test_filter_dfs_by_visual():
    from pbi_modules.ui.lineage_panel import filter_dfs

    dfs = _dfs_with_strings()
    filtered = filter_dfs(dfs, "VIS_ID :2")

    assert list(filtered["visuals"]["ID визуала"]) == ["VIS_ID :2"]
    assert len(filtered["lineage"]) == 1


def test_clear_filter_restores_original_data(window):
    dfs = _dfs_with_strings()
    window.current_dfs = dfs
    window._original_dfs = dfs.copy()

    window.lineage_focus_combo.addItem("Мера | M2 | Мера2", userData="M2")
    window.lineage_focus_combo.setCurrentIndex(0)
    window._apply_lineage_focus()
    window._clear_lineage_filter()

    assert len(window.current_dfs["measures"]) == 3
    assert len(window.current_dfs["visuals"]) == 3
