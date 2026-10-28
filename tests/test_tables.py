"""format_cell: безопасное строковое представление значения ячейки."""
import os

import pandas as pd
import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pbi_modules.gui import format_cell  # noqa: E402


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (None, ""),
        (pd.NA, ""),
        (pd.NaT, ""),
        (float("nan"), ""),
        ("", ""),
        ("текст", "текст"),
        (0, "0"),
        (12.5, "12.5"),
        (True, "True"),
        ("строка\nс переносом", "строка\nс переносом"),
    ],
)
def test_format_cell_values(value, expected):
    assert format_cell(value) == expected


def test_format_cell_truncates_long_text():
    text = "x" * 300

    result = format_cell(text)

    assert len(result) == 200 + len("...")
    assert result.endswith("...")


def test_format_cell_keeps_short_text_intact():
    text = "y" * 200

    assert format_cell(text) == text


def test_format_cell_handles_containers():
    assert format_cell([1, 2]) == "[1, 2]"
    assert format_cell({"a": 1}) == "{'a': 1}"
