"""
Запуск PBIT Analyzer с графическим интерфейсом.

Точка входа для исходников и для сборки GUI-EXE (build_exe.py).
Модули приложения лежат в _work, поэтому путь добавляется до импорта.

Использование:
    python run_gui.py
"""
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
WORK_DIR = PROJECT_ROOT / "_work"

if str(WORK_DIR) not in sys.path:
    sys.path.insert(0, str(WORK_DIR))


def main() -> int:
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
