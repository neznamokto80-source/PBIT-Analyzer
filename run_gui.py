"""
Запуск PBIT Analyzer с графическим интерфейсом.

Точка входа для исходников и для сборки GUI-EXE (build_exe.py).
Модули приложения лежат в _work, поэтому путь добавляется до импорта.

Использование:
    python run_gui.py
    python run_gui.py --self-test   # проверка сборки: окно создаётся и закрывается
"""
import os
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
WORK_DIR = PROJECT_ROOT / "_work"

if str(WORK_DIR) not in sys.path:
    sys.path.insert(0, str(WORK_DIR))


def _say(message: str) -> None:
    """Вывод для windowed-сборки: sys.stdout там может быть None."""
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.write(message + "\n")
            stream.flush()
        except (OSError, ValueError):
            pass
        return


def self_test() -> int:
    """Smoke-тест бандла: импорты, документация справки и окно без показа."""
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    try:
        import pbixray  # noqa: F401 — проверяем, что бинарные зависимости попали в EXE
        from PyQt6.QtWidgets import QApplication

        from pbi_modules.file_kind import detect_file_kind  # noqa: F401
        from pbi_modules.gui import PBITAnalyzerMainWindow
        from pbi_modules.help_content import load_changelog_html, load_schema_doc_html
        from pbi_modules.pbix_analyzer import analyze_pbix_file  # noqa: F401
        from pbi_modules.pbix_extractor import PBIXModelExtractor  # noqa: F401
    except ImportError as error:
        _say(f"Не удалось загрузить зависимости ({error}).")
        return 1

    # Вкладки «Схема PBIT» и «Changelog» берут markdown из бандла: если файлы
    # не попали в EXE, справка показывала «файл не найден».
    missing_docs = [
        name
        for name, html in (
            ("schema_documentation.md", load_schema_doc_html()),
            ("plan/CHANGELOG.md", load_changelog_html()),
        )
        if "не найден" in html
    ]
    if missing_docs:
        _say(f"Справка: не найдены {', '.join(missing_docs)}")
        return 3

    try:
        app = QApplication([])
        PBITAnalyzerMainWindow()
        app.quit()
    except Exception as error:  # noqa: BLE001 — smoke-тест обязан вернуть код, а не упасть
        _say(f"Self-test failed: {type(error).__name__}: {error}")
        return 2
    _say("Self-test OK")
    return 0


def main() -> int:
    if "--self-test" in sys.argv:
        return self_test()

    try:
        from PyQt6.QtWidgets import QApplication

        from pbi_modules.gui import PBITAnalyzerMainWindow
        from pbi_modules.logging_config import setup_logging
    except ImportError as error:
        print(
            f"Не удалось загрузить зависимости ({error}).\n"
            "Установите их командой:\n"
            "    python -m pip install -r requirements.txt"
        )
        return 1

    setup_logging()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = PBITAnalyzerMainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
