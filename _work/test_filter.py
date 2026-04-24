#!/usr/bin/env python3
"""Тест фильтра lineage."""
import sys
sys.path.insert(0, '.')

from PyQt6.QtWidgets import QApplication
from pbi_modules.gui import PBITAnalyzerMainWindow

app = QApplication(sys.argv)
window = PBITAnalyzerMainWindow()
print("Окно создано успешно")
print("Проверка атрибутов:", hasattr(window, 'lineage_focus_combo'))
print("Completer установлен:", window.lineage_focus_combo.completer() is not None)
if window.lineage_focus_combo.completer():
    print("Filter mode:", window.lineage_focus_combo.completer().filterMode())
sys.exit(0)