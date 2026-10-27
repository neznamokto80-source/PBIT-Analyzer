"""
Сборка EXE файлов PBIT Analyzer через PyInstaller.

Использование:
    python build_exe.py              # Собрать оба варианта (GUI + CLI)
    python build_exe.py --gui        # Только GUI
    python build_exe.py --cli        # Только CLI
    python build_exe.py --onefile    # Один файл (медленнее запуск, по умолчанию)
    python build_exe.py --dir        # Папка с файлами (быстрее запуск)
    python build_exe.py --clean      # Только очистка build/ dist/ и .spec
    python build_exe.py --no-copy    # Не копировать EXE в корень проекта

После сборки промежуточные файлы (build/, dist/ и *.spec) удаляются автоматически:
остаётся только готовый EXE в корне проекта. С `--no-copy` и `--dir` папка dist/
сохраняется — именно оттуда EXE забирает CI.

Требования:
    pip install pyinstaller
"""
import argparse
import glob
import importlib
import os
import re
import shutil
import subprocess
import sys


def use_utf8_output() -> None:
    """Переводит stdout/stderr в UTF-8.

    На Windows без UTF-8 режима кириллица в print() падает с
    UnicodeEncodeError (например, на CI-раннерах с cp1252).
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass


use_utf8_output()

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
WORK_DIR = os.path.join(PROJECT_ROOT, "_work")
DIST_DIR = os.path.join(PROJECT_ROOT, "dist")
BUILD_DIR = os.path.join(PROJECT_ROOT, "build")

GUI_SCRIPT = os.path.join(PROJECT_ROOT, "run_gui.py")
CLI_SCRIPT = os.path.join(PROJECT_ROOT, "run_cli.py")
GUI_NAME = "PBIT_Analyzer_GUI"
CLI_NAME = "PBIT_Analyzer_CLI"

# pbi_modules лежит в _work — нужен и для импорта, и для анализа hidden-import'ов.
sys.path.insert(0, WORK_DIR)

APP_VERSION = "0"


def detect_app_version() -> str:
    """Версия приложения из plan/CHANGELOG.md (первый заголовок вида ## [2.1.0])."""
    changelog = os.path.join(PROJECT_ROOT, "plan", "CHANGELOG.md")
    try:
        with open(changelog, "r", encoding="utf-8") as file:
            match = re.search(r"^##\s*\[(\d+(?:\.\d+)*)\]", file.read(), re.MULTILINE)
        if match:
            return match.group(1)
    except OSError:
        pass
    print("  [версия не найдена] plan/CHANGELOG.md — использую 0")
    return "0"


APP_VERSION = detect_app_version()
VERSION_SUFFIX = f"_v{APP_VERSION}"

# Модули, которые PyInstaller может не увидеть при статическом анализе.
COMMON_HIDDEN_IMPORTS = [
    "pandas",
    "pandas.core",
    "pandas.io.excel",
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
    "openpyxl.workbook",
    "pbi_modules",
    "pbi_modules.analyzer",
    "pbi_modules.extractors",
    "pbi_modules.excel_report",
    "pbi_modules.formatters",
    "pbi_modules.lineage",
    "pbi_modules.logging_config",
    "pbi_modules.file_kind",
    "pbi_modules.dispatcher",
    "pbi_modules.pbix_extractor",
    "pbi_modules.pbix_analyzer",
    # PBIX: pbixray + нативные декодеры VertiPaq и SQLite-движок для метаданных.
    "pbixray",
    "apsw",
    "xpress8",
    "xpress9",
    "xmhuffman",
]

# xpress8/xpress9/xmhuffman — одиночные .pyd-модули: им достаточно
# --hidden-import, --collect-all для них не нужен.
COLLECT_ALL_PACKAGES: list[str] = []

GUI_HIDDEN_IMPORTS = COMMON_HIDDEN_IMPORTS + [
    "PyQt6",
    "PyQt6.sip",
    "PyQt6.QtCore",
    "PyQt6.QtGui",
    "PyQt6.QtWidgets",
    "pbi_modules.gui",
    "pbi_modules.help_content",
    "pbi_modules.help_dialog",
]

# CLI собирается без PyQt6: headless-версии Qt не нужен.
CLI_HIDDEN_IMPORTS = COMMON_HIDDEN_IMPORTS

CLI_EXCLUDES = [
    "PyQt6",
    "tkinter",
]

# В окружении разработчика могут стоять другие Qt-биндинги (например, PyQt5).
# PyInstaller не собирает два биндинга сразу, поэтому лишние явно исключаем.
GUI_EXCLUDES = [
    "tkinter",
    "PyQt5",
    "PySide2",
    "PySide6",
]

# Данные для копирования в EXE: (исходник, папка внутри бандла).
# styles.qss читается из pbi_modules/gui.py, help_content.py ищет корень проекта
# по pyproject.toml и берёт оттуда schema_documentation.md и plan/CHANGELOG.md.
GUI_DATAS = [
    (os.path.join(WORK_DIR, "pbi_modules", "styles.qss"), "pbi_modules"),
    (os.path.join(PROJECT_ROOT, "pyproject.toml"), "."),
    (os.path.join(PROJECT_ROOT, "README.md"), "."),
    (os.path.join(PROJECT_ROOT, "schema_documentation.md"), "."),
    (os.path.join(PROJECT_ROOT, "plan", "CHANGELOG.md"), os.path.join("plan")),
]

CLI_DATAS = [
    (os.path.join(PROJECT_ROOT, "README.md"), "."),
]


def clean() -> None:
    """Очищает директории сборки и сгенерированные .spec файлы."""
    for directory in [DIST_DIR, BUILD_DIR]:
        if os.path.exists(directory):
            shutil.rmtree(directory, ignore_errors=True)
    for name in (GUI_NAME, CLI_NAME):
        for spec in (f"{name}.spec", f"{name}{VERSION_SUFFIX}.spec"):
            spec_path = os.path.join(PROJECT_ROOT, spec)
            if os.path.exists(spec_path):
                os.remove(spec_path)
    print("Очистка завершена")


def ensure_pyinstaller() -> None:
    """Гарантирует наличие PyInstaller перед сборкой."""
    try:
        import PyInstaller

        print(f"PyInstaller: {PyInstaller.__version__}")
    except ImportError:
        print("PyInstaller не установлен. Устанавливаю...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def ensure_runtime_packages(with_gui: bool) -> None:
    """Устанавливает недостающие библиотеки, без которых EXE не запустится."""
    packages = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "pbixray": "pbixray",
    }
    if with_gui:
        packages["PyQt6"] = "PyQt6"

    missing = []
    for module, pip_name in packages.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        print("Зависимости: установлены")
        return

    print(f"Установка недостающих пакетов: {', '.join(missing)}")
    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])


def filter_hidden_imports(hidden_imports: list[str]) -> list[str]:
    """Оставляет только реально импортируемые hidden-import'ы.

    PyInstaller падает на --hidden-import для неустановленных модулей,
    поэтому недоступные имена отбрасываются, а не ломают сборку.
    """
    kept = []
    for name in hidden_imports:
        try:
            importlib.import_module(name)
            kept.append(name)
        except (ImportError, ModuleNotFoundError):
            print(f"  [пропущен hidden-import] {name} — не установлен/не найден")
    return kept


def _existing_datas(datas: list[tuple[str, str]]) -> list[tuple[str, str]]:
    result = []
    for src, dst in datas:
        if os.path.exists(src):
            result.append((src, dst))
        else:
            print(f"  [пропущен --add-data] {src} — файл не найден")
    return result


def cleanup_build_artifacts(keep_dist: bool = False) -> list[str]:
    """Удаляет промежуточные файлы PyInstaller после сборки.

    keep_dist=True оставляет dist/ (режимы `--no-copy` и `--dir`, а также CI);
    иначе удаляется и dist/ — готовый EXE уже лежит в корне проекта.
    Возвращает имена удалённых файлов и папок для отчёта в консоли.
    """
    removed: list[str] = []
    for directory in (BUILD_DIR, None if keep_dist else DIST_DIR):
        if not directory or not os.path.exists(directory):
            continue
        shutil.rmtree(directory, ignore_errors=True)
        if not os.path.exists(directory):
            removed.append(os.path.basename(directory))

    for spec_path in sorted(glob.glob(os.path.join(PROJECT_ROOT, "*.spec"))):
        try:
            os.remove(spec_path)
        except OSError as error:
            print(f"  [не удалось удалить] {spec_path}: {error}")
            continue
        removed.append(os.path.basename(spec_path))

    return removed


def _copy_exe(name: str, onefile: bool) -> None:
    """Копирует готовый EXE в корень проекта (как create_exe.bat)."""
    if not onefile:
        print(f"Результат (папка): {os.path.join(DIST_DIR, name)}")
        return

    exe_path = os.path.join(DIST_DIR, f"{name}.exe")
    if not os.path.exists(exe_path):
        print(f"  [нет файла для копирования] {exe_path}")
        return

    target = os.path.join(PROJECT_ROOT, f"{name}.exe")
    try:
        shutil.copy2(exe_path, target)
        size_mb = os.path.getsize(target) / 1048576
        print(f"Скопировано в корень: {target} ({size_mb:.1f} МБ)")
    except OSError as error:
        print(f"  [не удалось скопировать] {error}")


def build(
    *,
    name: str,
    script: str,
    windowed: bool,
    onefile: bool,
    hidden_imports: list[str],
    excludes: list[str],
    datas: list[tuple[str, str]],
    copy_to_root: bool = True,
) -> bool:
    """Собирает один EXE-файл. Возвращает True при успехе."""
    full_name = f"{name}{VERSION_SUFFIX}"
    mode = "GUI" if windowed else "CLI"
    print(f"\n=== Сборка {mode}: {full_name} ===")

    if not os.path.exists(script):
        print(f"Ошибка: точка входа не найдена — {script}")
        return False

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", full_name,
        "--windowed" if windowed else "--console",
        "--noconfirm",
        "--clean",
        "--onefile" if onefile else "--onedir",
        "--paths", WORK_DIR,
    ]

    for hidden_import in filter_hidden_imports(hidden_imports):
        cmd.extend(["--hidden-import", hidden_import])

    for module in excludes:
        cmd.extend(["--exclude-module", module])

    for package in COLLECT_ALL_PACKAGES:
        if importlib.util.find_spec(package) is not None:
            cmd.extend(["--collect-all", package])
        else:
            print(f"  [пропущен --collect-all] {package} — не установлен")

    for src, dst in _existing_datas(datas):
        cmd.extend(["--add-data", f"{src}{os.pathsep}{dst}"])

    cmd.append(script)

    print(f"Команда: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=PROJECT_ROOT)
    if result.returncode != 0:
        print(f"Ошибка сборки {mode} (код {result.returncode})")
        return False

    if copy_to_root:
        _copy_exe(full_name, onefile)
    else:
        print(f"Результат: {os.path.join(DIST_DIR, full_name)}")
    return True


def build_gui(onefile: bool = True, copy_to_root: bool = True) -> bool:
    """Собирает GUI версию."""
    return build(
        name=GUI_NAME,
        script=GUI_SCRIPT,
        windowed=True,
        onefile=onefile,
        hidden_imports=GUI_HIDDEN_IMPORTS,
        excludes=GUI_EXCLUDES,
        datas=GUI_DATAS,
        copy_to_root=copy_to_root,
    )


def build_cli(onefile: bool = True, copy_to_root: bool = True) -> bool:
    """Собирает CLI версию (headless, без PyQt6)."""
    return build(
        name=CLI_NAME,
        script=CLI_SCRIPT,
        windowed=False,
        onefile=onefile,
        hidden_imports=CLI_HIDDEN_IMPORTS,
        excludes=CLI_EXCLUDES,
        datas=CLI_DATAS,
        copy_to_root=copy_to_root,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Сборка EXE файлов PBIT Analyzer")
    parser.add_argument("--gui", action="store_true", help="Только GUI")
    parser.add_argument("--cli", action="store_true", help="Только CLI")
    parser.add_argument("--onefile", action="store_true", help="Один файл (по умолчанию)")
    parser.add_argument("--dir", action="store_true", help="Папка с файлами")
    parser.add_argument("--clean", action="store_true", help="Только очистка")
    parser.add_argument("--no-copy", action="store_true", help="Не копировать EXE в корень проекта")
    args = parser.parse_args(argv)

    onefile = not args.dir

    if args.clean:
        clean()
        return 0

    ensure_pyinstaller()
    print(f"Версия приложения: {APP_VERSION}")
    print(f"Режим: {'onefile' if onefile else 'onedir'}")

    results = []
    if not args.cli:
        ensure_runtime_packages(with_gui=True)
        results.append(("GUI", build_gui(onefile, not args.no_copy)))

    if not args.gui:
        ensure_runtime_packages(with_gui=False)
        results.append(("CLI", build_cli(onefile, not args.no_copy)))

    print("\n=== Итоги ===")
    success = True
    for name, ok in results:
        print(f"  {name}: {'OK' if ok else 'ОШИБКА'}")
        success = success and ok

    if args.no_copy:
        print("  (копирование в корень отключено: --no-copy)")

    copied_exes = glob.glob(os.path.join(PROJECT_ROOT, f"PBIT_Analyzer_*{VERSION_SUFFIX}.exe"))
    keep_dist = args.no_copy or not onefile or not success or len(copied_exes) < len(results)

    if keep_dist and os.path.exists(DIST_DIR):
        print(f"\nРезультаты в: {DIST_DIR}")
        for item in sorted(os.listdir(DIST_DIR)):
            print(f"  {item}")

    removed = cleanup_build_artifacts(keep_dist=keep_dist)
    if keep_dist and os.path.exists(DIST_DIR):
        print("\nПапка dist/ сохранена (--no-copy, --dir или сборка без копирования EXE)")
    if removed:
        print(f"\nУдалены временные файлы сборки: {', '.join(removed)}")
    else:
        print("\nВременных файлов сборки не осталось")

    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
