"""Форматирование DAX: структура вывода и стабильность."""
from pbi_modules.formatters import DAXFormatter, SourceFormatter, SQLFormatter


def test_format_measure_adds_name_and_indent():
    result = DAXFormatter.format_measure("Total", "SUM(Sales[Amount])")

    assert result.splitlines()[0] == "Total ="
    # DAXFormatter ставит пробел перед скобкой: «SUM (Sales[Amount])».
    assert "SUM (Sales[Amount])" in result
    assert result.splitlines()[1].startswith("    ")


def test_format_measure_handles_var_return():
    result = DAXFormatter.format_measure("Calc", "VAR x = SUM(T[C]) RETURN x * 2")

    assert "VAR x =" in result
    assert "RETURN" in result
    assert "x * 2" in result


def test_format_column_and_table_use_same_render():
    expression = "SUM(Sales[Amount])"

    assert DAXFormatter.format_column("C", expression) == DAXFormatter.format_measure("C", expression)
    assert DAXFormatter.format_table("T", expression) == DAXFormatter.format_measure("T", expression)


def test_format_empty_expression_returns_dash():
    assert DAXFormatter.format_measure("Total", "") == "—"
    assert DAXFormatter.format_measure("Total", None) == "—"


def test_format_dax_formula_extracts_measure_name():
    result = DAXFormatter.format_dax_formula("Total = SUM(Sales[Amount])")

    assert result.splitlines()[0] == "Total ="
    assert result.count("Total =") == 1


def test_format_dax_formula_without_name():
    result = DAXFormatter.format_dax_formula("SUM(Sales[Amount])")

    assert "=" not in result.splitlines()[0]
    assert "SUM (Sales[Amount])" in result


def test_source_formatter_dispatches_by_type():
    m_code = 'let Source = Sql.Database("srv", "db", [Query="SELECT 1"]) in Source'

    formatted_m = SourceFormatter.format_source("m", m_code)
    assert "\n" in formatted_m or "let" in formatted_m

    calculated = SourceFormatter.format_source("calculated", "SUM(A[B])")
    assert "SUM" in calculated and "A[B]" in calculated

    raw = SourceFormatter.format_source("entity", "unchanged")
    assert raw == "unchanged"


def test_sql_formatter_keeps_text_without_sql():
    # M-код переносится на строки, но не теряет содержимого.
    result = SQLFormatter.format_sql_in_m("let x = 1 in x")
    assert result.startswith("let")
    assert "x = 1" in result
    assert SQLFormatter.format_sql_in_m(None) is None
