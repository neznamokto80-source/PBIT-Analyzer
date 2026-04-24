#!/usr/bin/env python3
"""Тест фильтра lineage с проверкой filterMode."""
import sys
sys.path.insert(0, '.')

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from pbi_modules.gui import PBITAnalyzerMainWindow

app = QApplication(sys.argv)
window = PBITAnalyzerMainWindow()
print("Окно создано успешно")
completer = window.lineage_focus_combo.completer()
if completer:
    print("Completer установлен")
    print("Filter mode:", completer.filterMode())
    print("MatchContains?", completer.filterMode() == Qt.MatchFlag.MatchContains)
    print("Case sensitivity:", completer.caseSensitivity())
else:
    print("Completer отсутствует")
sys.exit(0)