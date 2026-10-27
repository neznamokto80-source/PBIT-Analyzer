import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import build_exe  # noqa: E402


def test_cleanup_removes_build_dir_and_specs(monkeypatch, tmp_path):
    build_dir = tmp_path / "build"
    dist_dir = tmp_path / "dist"
    build_dir.mkdir()
    dist_dir.mkdir()
    exe_path = dist_dir / "PBIT_Analyzer_GUI_v2.2.0.exe"
    exe_path.write_bytes(b"MZ")
    spec_path = tmp_path / "PBIT_Analyzer_GUI_v2.2.0.spec"
    spec_path.write_text("x", encoding="utf-8")

    monkeypatch.setattr(build_exe, "PROJECT_ROOT", str(tmp_path))
    monkeypatch.setattr(build_exe, "BUILD_DIR", str(build_dir))
    monkeypatch.setattr(build_exe, "DIST_DIR", str(dist_dir))

    removed = build_exe.cleanup_build_artifacts(keep_dist=True)

    assert set(removed) == {"build", "PBIT_Analyzer_GUI_v2.2.0.spec"}
    assert not build_dir.exists()
    assert not spec_path.exists()
    assert exe_path.is_file()


def test_cleanup_removes_dist_when_exe_copied(monkeypatch, tmp_path):
    build_dir = tmp_path / "build"
    dist_dir = tmp_path / "dist"
    build_dir.mkdir()
    dist_dir.mkdir()
    (dist_dir / "PBIT_Analyzer_GUI_v2.2.0.exe").write_bytes(b"MZ")
    exe_path = tmp_path / "PBIT_Analyzer_GUI_v2.2.0.exe"
    exe_path.write_bytes(b"MZ")

    monkeypatch.setattr(build_exe, "PROJECT_ROOT", str(tmp_path))
    monkeypatch.setattr(build_exe, "BUILD_DIR", str(build_dir))
    monkeypatch.setattr(build_exe, "DIST_DIR", str(dist_dir))

    removed = build_exe.cleanup_build_artifacts()

    assert {"build", "dist"} <= set(removed)
    assert not dist_dir.exists()
    assert exe_path.is_file()
