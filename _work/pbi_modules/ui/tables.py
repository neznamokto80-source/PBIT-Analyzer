"""Таблица результатов: контекстное меню, копирование, безопасные ячейки."""
import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QHeaderView,
    QMenu,
    QTableWidget,
    QTableWidgetItem,
)

from ..constants import LINEAGE_HEAVY_TABLE_ROWS_THRESHOLD, MAX_CELL_CHARS


def format_cell(value, max_chars: int = MAX_CELL_CHARS) -> str:
    """Безопасное строковое представление значения ячейки.

    Отличается от ``str(value)`` только для отсутствующих значений: None, NaN,
    NaT и pd.NA дают пустую строку, а не «None»/«nan». Длинный текст обрезается.
    """
    if value is None:
        return ""
    try:
        if pd.api.types.is_scalar(value) and pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = value if isinstance(value, str) else str(value)
    if len(text) > max_chars:
        text = text[:max_chars] + "..."
    return text


class CustomTableWidget(QTableWidget):
    """Таблица с возможностью копирования и контекстным меню."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)

    def _show_context_menu(self, position):
        menu = QMenu(self)
        copy_cell = QAction("Копировать ячейку", self)
        copy_cell.triggered.connect(self._copy_cell)
        menu.addAction(copy_cell)
        copy_row = QAction("Копировать строку", self)
        copy_row.triggered.connect(self._copy_row)
        menu.addAction(copy_row)
        copy_selection = QAction("Копировать выделенное", self)
        copy_selection.triggered.connect(self._copy_selection)
        menu.addAction(copy_selection)
        menu.addSeparator()
        copy_all = QAction("Копировать всю таблицу", self)
        copy_all.triggered.connect(self._copy_all)
        menu.addAction(copy_all)
        menu.exec(self.mapToGlobal(position))

    def _copy_cell(self):
        indexes = self.selectedIndexes()
        if indexes:
            item = indexes[0]
            QApplication.clipboard().setText(str(item.data(Qt.ItemDataRole.DisplayRole) or ""))

    def _copy_row(self):
        indexes = self.selectedIndexes()
        if not indexes:
            return
        row = indexes[0].row()
        row_data = []
        for col in range(self.columnCount()):
            item = self.item(row, col)
            row_data.append(str(item.text()) if item else "")
        QApplication.clipboard().setText("\t".join(row_data))

    def _copy_selection(self):
        indexes = self.selectedIndexes()
        if not indexes:
            return
        data_dict = {}
        for idx in indexes:
            row = idx.row()
            col = idx.column()
            if row not in data_dict:
                data_dict[row] = {}
            data_dict[row][col] = str(idx.data(Qt.ItemDataRole.DisplayRole) or "")
        rows_text = []
        for row in sorted(data_dict.keys()):
            cols = data_dict[row]
            row_values = []
            for col in range(self.columnCount()):
                row_values.append(cols.get(col, ""))
            rows_text.append("\t".join(row_values))
        QApplication.clipboard().setText("\n".join(rows_text))

    def _copy_all(self):
        all_data = []
        headers = []
        for col in range(self.columnCount()):
            headers.append(self.horizontalHeaderItem(col).text() if self.horizontalHeaderItem(col) else "")
        all_data.append("\t".join(headers))
        for row in range(self.rowCount()):
            row_data = []
            for col in range(self.columnCount()):
                item = self.item(row, col)
                row_data.append(str(item.text()) if item else "")
            all_data.append("\t".join(row_data))
        QApplication.clipboard().setText("\n".join(all_data))


def _log(callback, message: str) -> None:
    if callback:
        callback(message)


def create_table_from_df(df, df_key: str = "", log=None) -> QTableWidget:
    table = CustomTableWidget()
    table.setRowCount(len(df))
    table.setColumnCount(len(df.columns))
    table.setHorizontalHeaderLabels(df.columns.tolist())

    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            value = df.iloc[row_idx, col_idx]
            item = QTableWidgetItem(format_cell(value))
            raw = str(value)
            item.setToolTip(raw if len(raw) > 100 else "")
            table.setItem(row_idx, col_idx, item)

    header = table.horizontalHeader()
    header.setStretchLastSection(True)
    is_heavy_lineage_table = (
        df_key in {"lineage", "lineage_upstream", "lineage_downstream"}
        and len(df) >= LINEAGE_HEAVY_TABLE_ROWS_THRESHOLD
    )
    if is_heavy_lineage_table:
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        table.verticalHeader().setDefaultSectionSize(24)
        _log(
            log,
            f"Таблица {df_key} большая ({len(df)} строк), "
            "использован облегченный режим ширины колонок.",
        )
    else:
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    return table
