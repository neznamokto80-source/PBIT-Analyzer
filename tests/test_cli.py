import json
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import run_cli  # noqa: E402

PBIX_DIR = Path(__import__("os").environ.get("PBIT_TEST_PBIX_DIR", r"D:\pbix"))


def _write_pbit(path):
    model = {
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "columns": [{"name": "Amount", "dataType": "double"}],
                    "measures": [{"name": "Total Sales", "expression": "SUM(Sales[Amount])"}],
                    "partitions": [{"source": {"type": "m", "expression": "let x=1 in x"}}],
                }
            ],
            "relationships": [],
        }
    }
    layout = {"sections": [{"displayName": "Page 1", "visualContainers": []}]}
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("DataModelSchema", json.dumps(model).encode("utf-16-le"))
        archive.writestr("Layout", json.dumps(layout).encode("utf-16-le"))
    return path


def test_cli_analyzes_pbit_and_returns_zero(tmp_path):
    pbit = _write_pbit(tmp_path / "sample.pbit")

    assert run_cli.main([str(pbit), "-o", str(tmp_path), "-q"]) == 0
    assert list(tmp_path.glob("pbi_report_sample_*.xlsx"))


def test_cli_rejects_unsupported_extension(tmp_path):
    path = tmp_path / "data.txt"
    path.write_text("x", encoding="utf-8")

    assert run_cli.main([str(path)]) == 2


def test_cli_rejects_mismatched_extension(tmp_path):
    renamed = tmp_path / "renamed.pbix"
    _write_pbit(renamed)

    assert run_cli.main([str(renamed), "-o", str(tmp_path), "-q"]) == 3


def test_cli_returns_three_for_broken_pbix(tmp_path):
    broken = tmp_path / "broken.pbix"
    broken.write_bytes(b"broken")

    assert run_cli.main([str(broken), "-o", str(tmp_path), "-q"]) == 3


def test_cli_returns_two_for_visual_focus_on_pbix(tmp_path):
    files = sorted(PBIX_DIR.glob("*.pbix")) if PBIX_DIR.is_dir() else []
    if not files:
        pytest.skip("Нет .pbix для проверки фокуса")

    assert run_cli.main([str(files[0]), "--focus", "VIS_ID :x", "-q"]) == 2
