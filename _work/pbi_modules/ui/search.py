"""Поиск по таблице: панель с полем ввода и фильтрация строк."""
from PyQt6.QtWidgets import QHBoxLayout, QLineEdit, QPushButton, QTableWidget, QWidget


def filter_table(table: QTableWidget, query: str) -> None:
    """Скрывает строки, не содержащие подстроку (регистр не важен)."""
    query = query.strip().lower()
    for row in range(table.rowCount()):
        match = False
        if not query:
            match = True
        else:
            for col in range(table.columnCount()):
                item = table.item(row, col)
                if item and query in item.text().lower():
                    match = True
                    break
        table.setRowHidden(row, not match)


def clear_table_filter(table: QTableWidget, search_input: QLineEdit) -> None:
    search_input.clear()
    for row in range(table.rowCount()):
        table.setRowHidden(row, False)


class SearchBar(QWidget):
    """Поле поиска и кнопки «Найти»/«Сброс» над таблицей."""

    def __init__(self, table: QTableWidget, parent=None):
        super().__init__(parent)
        self.table = table
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        self.input = QLineEdit()
        self.input.setObjectName("searchInput")
        self.input.setPlaceholderText("Поиск по таблице...")
        self.input.setClearButtonEnabled(True)
        layout.addWidget(self.input, 1)

        btn_search = QPushButton("Найти")
        btn_search.setObjectName("btnSearch")
        btn_search.clicked.connect(self.apply)
        layout.addWidget(btn_search)

        btn_clear = QPushButton("Сброс")
        btn_clear.setObjectName("btnClearSearch")
        btn_clear.clicked.connect(self.clear)
        layout.addWidget(btn_clear)

        self.input.returnPressed.connect(self.apply)

    def apply(self) -> None:
        filter_table(self.table, self.input.text())

    def clear(self) -> None:
        clear_table_filter(self.table, self.input)
