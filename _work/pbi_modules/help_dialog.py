"""
Модуль диалога справки с вертикальным скроллом.
"""
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTextEdit,
)
from .help_content import HELP_TEXT


class HelpDialog(QDialog):
    """Диалоговое окно справки с вертикальным скроллом."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Справка")
        self.setMinimumWidth(1000)
        self.setMinimumHeight(700)
        self.resize(1000, 700)
        
        layout = QVBoxLayout(self)
        
        # Текстовое поле с поддержкой HTML и вертикальным скроллом
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setAcceptRichText(True)
        self.text_edit.setHtml(HELP_TEXT)
        self.text_edit.setStyleSheet("""
            QTextEdit {
                font-size: 14px;
                background-color: white;
            }
        """)
        # Включение вертикального скролла (по умолчанию включен)
        self.text_edit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.text_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        layout.addWidget(self.text_edit)
        
        # Кнопка закрытия
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.close_button = QPushButton("Закрыть")
        self.close_button.clicked.connect(self.accept)
        button_layout.addWidget(self.close_button)
        layout.addLayout(button_layout)
        
        # Настройка фокуса на кнопке закрытия для удобства
        self.close_button.setFocus()
    
    def set_help_text(self, text: str):
        """Обновить текст справки (опционально)."""
        self.text_edit.setHtml(text)