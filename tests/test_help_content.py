"""Справка: markdown-документация должна находиться и внутри бандла EXE."""
import sys
from pathlib import Path

from pbi_modules import help_content


def _make_bundle(root: Path) -> Path:
    """Повторяет раскладку onefile-сборки: документы рядом с папкой pbi_modules."""
    (root / "pbi_modules").mkdir(parents=True)
    (root / "docs").mkdir()
    (root / "docs" / "schema_documentation.md").write_text("# Схема PBIT\n\nТекст", encoding="utf-8")
    (root / "docs" / "CHANGELOG.md").write_text("# Changelog\n\n- пункт", encoding="utf-8")
    return root


def test_help_docs_are_found_in_frozen_bundle(monkeypatch, tmp_path):
    bundle = _make_bundle(tmp_path / "_MEI12345")
    monkeypatch.setattr(sys, "_MEIPASS", str(bundle), raising=False)
    monkeypatch.setattr(help_content, "__file__", str(bundle / "pbi_modules" / "help_content.py"))

    assert "Схема PBIT" in help_content.load_schema_doc_html()
    assert "Changelog" in help_content.load_changelog_html()


def test_help_docs_report_missing_files(monkeypatch, tmp_path):
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    monkeypatch.setattr(help_content, "__file__", str(tmp_path / "pbi_modules" / "help_content.py"))

    assert "не найден" in help_content.load_schema_doc_html()
    assert "не найден" in help_content.load_changelog_html()
