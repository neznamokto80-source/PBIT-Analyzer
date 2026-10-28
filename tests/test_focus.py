"""Единые правила focus-ID и формирования ID объектов модели."""
import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pbi_modules.focus import (  # noqa: E402
    display_focus_id,
    extract_visual_id,
    normalize_focus_id,
    parse_focus_label,
)
from pbi_modules.gui import PBITAnalyzerMainWindow  # noqa: E402
from pbi_modules.ids import column_number, measure_number, table_number  # noqa: E402
from pbi_modules.lineage import _normalize_focus_id as lineage_normalize  # noqa: E402

import run_cli  # noqa: E402

FOCUS_CASES = [
    ("S4", "s4"),
    ("s04", "s04"),
    ("M3", "M3"),
    ("a105", "A105"),
    ("VIS_ID :abc", "VIS:abc"),
    ("VIS:xyz", "VIS:xyz"),
    ("Тип | S4 | Продажи", "s4"),
    ("Мера | M22 | Total Sales", "M22"),
    ("M:7", "M7"),
    ("C:Amount", "Amount"),
    ("T:123", "123"),
    ("📊 S2", "s2"),
    ("", ""),
    ("мусор", "мусор"),
    ("  S1  ", "s1"),
]


@pytest.mark.parametrize(("raw", "expected"), FOCUS_CASES)
def test_normalize_focus_id(raw, expected):
    assert normalize_focus_id(raw) == expected


@pytest.mark.parametrize(("raw", "expected"), FOCUS_CASES)
def test_all_entry_points_agree(raw, expected):
    """GUI, lineage и CLI используют одни и те же правила."""
    assert lineage_normalize(raw) == expected
    assert run_cli.normalize_focus(raw) == expected
    assert PBITAnalyzerMainWindow._normalize_focus_input(raw) == expected


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("VIS_ID :abc", "abc"),
        ("VIS:abc", "abc"),
        ("plain", "plain"),
        ("", ""),
        (None, ""),
    ],
)
def test_extract_visual_id(value, expected):
    assert extract_visual_id(value) == expected


def test_parse_focus_label():
    assert parse_focus_label("Мера | M22 | Total Sales") == ("Мера", "M22", "Total Sales")
    assert parse_focus_label("S4") == ("Таблица", "s4", "")
    assert parse_focus_label("VIS:abc") == ("Визуал", "VIS:abc", "")
    assert parse_focus_label("без типа") == ("", "без типа", "")


def test_display_focus_id():
    assert display_focus_id("Визуал", "VIS:abc") == "VIS_ID :abc"
    assert display_focus_id("Мера", "M3") == "M3"
    assert display_focus_id("Таблица", "s1") == "S1"


@pytest.mark.parametrize(
    ("table", "column", "expected"),
    [("s1", 1, "A101"), ("s2", 5, "A205"), ("s10", 1, "A1001")],
)
def test_column_number(table, column, expected):
    assert column_number(table, column) == expected


def test_table_and_measure_numbers():
    assert table_number(1) == "s1"
    assert table_number(12) == "s12"
    assert measure_number(1) == "M1"
    assert measure_number(12) == "M12"
