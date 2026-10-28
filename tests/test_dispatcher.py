import json
import zipfile

import pytest
from pbi_modules import dispatcher
from pbi_modules.file_kind import PBIT, PBIX, FileFormatError, detect_file_kind


def _write_pbit(path):
    model = {"model": {"tables": [], "relationships": []}}
    layout = {"sections": []}
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("DataModelSchema", json.dumps(model).encode("utf-16-le"))
        archive.writestr("Layout", json.dumps(layout).encode("utf-16-le"))
    return path


def _write_pbix(path):
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("DataModel", b"\x00\x01")
        archive.writestr("Report/Layout", b"{}")
    return path


def test_detect_file_kind_by_content(tmp_path):
    assert detect_file_kind(str(_write_pbit(tmp_path / "a.pbit"))) == PBIT
    assert detect_file_kind(str(_write_pbix(tmp_path / "a.pbix"))) == PBIX


def test_detect_pbit_with_layout_in_report_folder(tmp_path):
    path = tmp_path / "nested.pbit"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("DataModelSchema", b"{}")
        archive.writestr("Report/Layout", b"{}")

    assert detect_file_kind(str(path)) == PBIT


def test_detect_rejects_mismatched_extension(tmp_path):
    pbit_named_pbix = _write_pbit(tmp_path / "wrong.pbix")

    with pytest.raises(FileFormatError, match="PBIT"):
        detect_file_kind(str(pbit_named_pbix))


def test_detect_rejects_unsupported_extension(tmp_path):
    path = tmp_path / "file.txt"
    path.write_text("data", encoding="utf-8")

    with pytest.raises(FileFormatError, match="Неподдерживаемый"):
        detect_file_kind(str(path))


def test_detect_rejects_corrupted_archive(tmp_path):
    path = tmp_path / "broken.pbix"
    path.write_bytes(b"not a zip")

    with pytest.raises(FileFormatError, match="повреждён"):
        detect_file_kind(str(path))


def test_detect_rejects_zip_without_powerbi_parts(tmp_path):
    path = tmp_path / "other.pbix"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("readme.txt", "hello")

    with pytest.raises(FileFormatError, match="DataModel"):
        detect_file_kind(str(path))


def test_dispatcher_routes_pbit(tmp_path, monkeypatch):
    calls = {}

    def fake_analyze_pbit(file_path, output_dir, **kwargs):
        calls["kind"] = "pbit"
        return {"excel_path": ""}

    monkeypatch.setattr("pbi_modules.analyzer.analyze_pbit_file", fake_analyze_pbit)
    result = dispatcher.analyze_file(str(_write_pbit(tmp_path / "a.pbit")), str(tmp_path))

    assert calls["kind"] == "pbit"
    assert result["file_kind"] == PBIT


def test_dispatcher_routes_pbix(tmp_path, monkeypatch):
    calls = {}

    def fake_analyze_pbix(file_path, output_dir, **kwargs):
        calls["kind"] = "pbix"
        calls["on_disk"] = kwargs.get("on_disk")
        return {"excel_path": ""}

    monkeypatch.setattr("pbi_modules.pbix_analyzer.analyze_pbix_file", fake_analyze_pbix)
    result = dispatcher.analyze_file(str(_write_pbix(tmp_path / "a.pbix")), str(tmp_path), on_disk=True)

    assert calls == {"kind": "pbix", "on_disk": True}
    assert result["file_kind"] == PBIX
