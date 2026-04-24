"""
PBIT Analyzer — GUI приложение для анализа Power BI Template файлов.
Точка входа (модульная версия).
"""

import subprocess
import sys

REQUIRED_PACKAGES = {
    "PyQt6": "PyQt6",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
}


def check_and_install_packages() -> bool:
    missing = []
    for package, pip_name in REQUIRED_PACKAGES.items():
        try:
            __import__(package)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        print("Все необходимые пакеты уже установлены.")
        return True

    print(f"Установка недостающих пакетов: {', '.join(missing)}")
    try:
        for package in missing:
            print(f"Установка {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print("Все пакеты установлены успешно.")
        return True
    except subprocess.CalledProcessError as error:
        print(f"Ошибка установки пакетов: {error}")
        return False


if not check_and_install_packages():
    print("Не удалось установить необходимые библиотеки. Завершение.")
    sys.exit(1)

from PyQt6.QtWidgets import QApplication
from pbi_modules.gui import PBITAnalyzerMainWindow
from pbi_modules.logging_config import setup_logging


def main():
    setup_logging()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = PBITAnalyzerMainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
