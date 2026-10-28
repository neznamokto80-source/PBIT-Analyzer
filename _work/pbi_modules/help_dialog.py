"""
Модуль диалога справки с вкладками.
Данные загружаются автоматически из файлов проекта.
"""
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
)

from .help_content import HELP_TEXT, load_changelog_html, load_schema_doc_html


class HelpDialog(QDialog):
    """Диалоговое окно справки с вкладками."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Справка")
        self.setMinimumWidth(1000)
        self.setMinimumHeight(700)
        self.resize(1000, 700)

        layout = QVBoxLayout(self)

        tabs = QTabWidget()

        # Вкладка "Справка"
        self.help_edit = QTextEdit()
        self.help_edit.setReadOnly(True)
        self.help_edit.setAcceptRichText(True)
        self.help_edit.setHtml(HELP_TEXT)
        self.help_edit.setStyleSheet("QTextEdit { font-size: 14px; background-color: white; }")
        self.help_edit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.help_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        tabs.addTab(self.help_edit, "Справка")

        # Вкладка "Схема PBIT" — загружается из schema_documentation.md
        self.schema_edit = QTextEdit()
        self.schema_edit.setReadOnly(True)
        self.schema_edit.setAcceptRichText(True)
        self.schema_edit.setHtml(load_schema_doc_html())
        self.schema_edit.setStyleSheet("QTextEdit { font-size: 14px; background-color: white; }")
        self.schema_edit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.schema_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        tabs.addTab(self.schema_edit, "Схема PBIT")

        # Вкладка "Changelog" — загружается из docs/CHANGELOG.md
        self.changelog_edit = QTextEdit()
        self.changelog_edit.setReadOnly(True)
        self.changelog_edit.setAcceptRichText(True)
        self.changelog_edit.setHtml(load_changelog_html())
        self.changelog_edit.setStyleSheet("QTextEdit { font-size: 14px; background-color: white; }")
        self.changelog_edit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.changelog_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        tabs.addTab(self.changelog_edit, "Changelog")

        layout.addWidget(tabs)

        # Кнопка закрытия
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.close_button = QPushButton("Закрыть")
        self.close_button.clicked.connect(self.accept)
        button_layout.addWidget(self.close_button)
        layout.addLayout(button_layout)

        self.close_button.setFocus()

    def set_help_text(self, text: str):
        """Обновить текст справки (опционально)."""
        self.help_edit.setHtml(text)
