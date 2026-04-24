import os
import re
import subprocess
import sys
import traceback
from datetime import datetime

import pandas as pd
from PyQt6.QtCore import QThread, Qt, pyqtSignal
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .analyzer import analyze_pbit_file
from .excel_report import ExcelReportGenerator
from .help_content import HELP_TEXT
from .help_dialog import HelpDialog
from .lineage import build_lineage_directions

TAB_MAPPING = [
    ('tables_ref', '📋 Справочник таблиц'),
    ('columns', '📊 Столбцы'),
    ('sources', '🔌 Источники'),
    ('relationships', '🔗 Связи'),
    ('measures', '📐 Меры'),
    ('sheets', '📑 Листы'),
    ('visuals', '🎨 Визуальные элементы'),
    ('lineage', '🧬 Линейдж'),
    ('lineage_upstream', '🧭 Lineage Upstream'),
    ('lineage_downstream', '🧭 Lineage Downstream'),
    ('conditional', '🎨 Условное форматирование'),
    ('stats', '📈 Статистика'),
]

LINEAGE_HEAVY_TABLE_ROWS_THRESHOLD = 1200


def open_path_in_os(path: str) -> None:
    if sys.platform == 'win32':
        os.startfile(path)
        return
    opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
    subprocess.run([opener, path], check=False)


class AnalysisWorker(QThread):
    """Фоновый поток для анализа PBIT файла."""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    log = pyqtSignal(str)

    def __init__(self, pbit_path, output_dir, keep_temp=False, create_excel=False):
        super().__init__()
        self.pbit_path = pbit_path
        self.output_dir = output_dir
        self.keep_temp = keep_temp
        self.create_excel = create_excel

    def run(self):
        try:
            self.log.emit(f"Начало анализа: {self.pbit_path}")
            dfs = analyze_pbit_file(
                pbit_path=self.pbit_path,
                output_dir=self.output_dir,
                keep_temp=self.keep_temp,
                create_excel=self.create_excel,
                progress_callback=lambda value, message: self.progress.emit(value, message),
                log_callback=lambda message: self.log.emit(message),
            )
            self.finished.emit(dfs)

        except Exception as e:
            self.error.emit(str(e))
            self.log.emit(f"ОШИБКА: {traceback.format_exc()}")


# ============================================================
# Custom Table Widget с контекстным меню
# ============================================================

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


# ============================================================
# Главное окно приложения
# ============================================================

class PBITAnalyzerMainWindow(QMainWindow):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()
        self.current_dfs = {}
        self.lineage_focus_map = {}
        self.output_dir = ""
        self.pbit_path = ""
        self.worker = None
        self.init_ui()

    def _set_analysis_ui_state(self, is_running):
        self.btn_analyze.setEnabled(not is_running)
        self.progress_bar.setVisible(is_running)
        if is_running:
            self.progress_bar.setValue(0)

    def _ensure_output_dir(self):
        if not self.output_dir and self.pbit_path:
            self.output_dir = os.path.dirname(self.pbit_path)
            self.output_dir_edit.setText(self.output_dir)
        if self.output_dir and not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

    def init_ui(self):
        self.setWindowTitle("PBIT Analyzer — Анализатор Power BI Template файлов")
        self.setGeometry(100, 100, 1400, 900)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        header_panel = self._create_header_panel()
        main_layout.addWidget(header_panel)

        top_panel = self._create_top_panel()
        main_layout.addWidget(top_panel)

        splitter = QSplitter(Qt.Orientation.Vertical)

        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        left_panel = self._create_left_panel()
        content_layout.addWidget(left_panel, 0)
        right_panel = self._create_right_panel()
        content_layout.addWidget(right_panel, 1)
        splitter.addWidget(content_widget)

        log_widget = self._create_log_panel()
        splitter.addWidget(log_widget)
        splitter.setSizes([700, 200])
        main_layout.addWidget(splitter, 1)

        self.statusBar().showMessage("Готово к работе. Выберите PBIT файл.")
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(300)
        self.progress_bar.setVisible(False)
        self.statusBar().addPermanentWidget(self.progress_bar)

    def _create_header_panel(self):
        panel = QWidget()
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        title_label = QLabel("🔍 PBIT Analyzer")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title_label)
        layout.addStretch()
        help_button = QPushButton("❓ Справка")
        help_button.clicked.connect(self._show_help)
        layout.addWidget(help_button)
        return panel

    def _create_top_panel(self):
        panel = QGroupBox("Файл и настройки")
        layout = QGridLayout(panel)
        layout.setSpacing(5)

        layout.addWidget(QLabel("PBIT файл:"), 0, 0)
        self.pbit_path_edit = QLabel("Файл не выбран")
        self.pbit_path_edit.setStyleSheet("color: gray; padding: 5px;")
        layout.addWidget(self.pbit_path_edit, 0, 1)
        btn_select_file = QPushButton("📁 Выбрать PBIT")
        btn_select_file.clicked.connect(self._select_pbit_file)
        layout.addWidget(btn_select_file, 0, 2)

        layout.addWidget(QLabel("Папка вывода:"), 1, 0)
        self.output_dir_edit = QLabel(self.output_dir)
        self.output_dir_edit.setStyleSheet("padding: 5px;")
        layout.addWidget(self.output_dir_edit, 1, 1)
        btn_select_dir = QPushButton("📂 Выбрать папку")
        btn_select_dir.clicked.connect(self._select_output_dir)
        layout.addWidget(btn_select_dir, 1, 2)

        settings_layout = QHBoxLayout()
        self.chk_create_excel = QCheckBox("Создать Excel")
        settings_layout.addWidget(self.chk_create_excel)
        self.chk_keep_temp = QCheckBox("Сохранить временные файлы")
        settings_layout.addWidget(self.chk_keep_temp)
        self.chk_open_folder = QCheckBox("Открыть папку результатов")
        settings_layout.addWidget(self.chk_open_folder)
        self.chk_open_excel = QCheckBox("Открыть Excel файл")
        settings_layout.addWidget(self.chk_open_excel)
        layout.addLayout(settings_layout, 2, 0, 1, 3)

        lineage_layout = QHBoxLayout()
        lineage_layout.addWidget(QLabel("Фильтр:"))
        self.lineage_focus_combo = QComboBox()
        self.lineage_focus_combo.setEditable(True)
        self.lineage_focus_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.lineage_focus_combo.setPlaceholderText("После анализа: S4 / A105 / M3 / VIS_ID :...")
        self.lineage_focus_combo.setToolTip("Выберите объект из списка или введите вручную")
        self.lineage_focus_combo.setEnabled(False)
        lineage_layout.addWidget(self.lineage_focus_combo, 1)
        self.btn_find_lineage = QPushButton("🔎 Найти")
        self.btn_find_lineage.setEnabled(False)
        self.btn_find_lineage.clicked.connect(self._apply_lineage_focus)
        lineage_layout.addWidget(self.btn_find_lineage)
        self.btn_clear_filter = QPushButton("🗑️ Очистить фильтр")
        self.btn_clear_filter.setEnabled(False)
        self.btn_clear_filter.clicked.connect(self._clear_lineage_filter)
        lineage_layout.addWidget(self.btn_clear_filter)
        layout.addLayout(lineage_layout, 3, 0, 1, 3)

        return panel

    def _create_left_panel(self):
        panel = QGroupBox("Действия")
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)

        self.btn_analyze = QPushButton("🚀 Начать анализ")
        self.btn_analyze.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50; color: white; font-weight: bold;
                font-size: 14px; padding: 15px; border-radius: 5px;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.btn_analyze.clicked.connect(self._start_analysis)
        layout.addWidget(self.btn_analyze)

        self.btn_export_excel = QPushButton("📥 Выгрузить в Excel")
        self.btn_export_excel.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0; color: white; font-size: 12px;
                padding: 10px; border-radius: 5px;
            }
            QPushButton:hover { background-color: #7b1fa2; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.btn_export_excel.clicked.connect(self._export_to_excel)
        self.btn_export_excel.setEnabled(False)
        layout.addWidget(self.btn_export_excel)

        self.btn_open_excel = QPushButton("📊 Открыть Excel")
        self.btn_open_excel.setStyleSheet("""
            QPushButton {
                background-color: #2196F3; color: white; font-size: 12px;
                padding: 10px; border-radius: 5px;
            }
            QPushButton:hover { background-color: #0b7dda; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.btn_open_excel.clicked.connect(self._open_excel)
        self.btn_open_excel.setEnabled(False)
        layout.addWidget(self.btn_open_excel)

        self.btn_open_folder = QPushButton("📁 Открыть папку")
        self.btn_open_folder.setStyleSheet("""
            QPushButton {
                background-color: #FF9800; color: white; font-size: 12px;
                padding: 10px; border-radius: 5px;
            }
            QPushButton:hover { background-color: #e68900; }
            QPushButton:disabled { background-color: #cccccc; }
        """)
        self.btn_open_folder.clicked.connect(self._open_output_folder)
        self.btn_open_folder.setEnabled(False)
        layout.addWidget(self.btn_open_folder)

        layout.addStretch()

        btn_help = QPushButton("❓ Справка")
        btn_help.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0; color: white; font-size: 12px;
                padding: 10px; border-radius: 5px;
            }
            QPushButton:hover { background-color: #7b1fa2; }
        """)
        btn_help.clicked.connect(self._show_help)
        layout.addWidget(btn_help)

        panel.setMaximumWidth(200)
        return panel

    def _create_right_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        self.result_tabs = QTabWidget()

        # Вкладка с таблицами результатов
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(False)
        self.tab_widget.setMovable(True)
        self.result_tabs.addTab(self.tab_widget, "Результаты парсинга")

        # Вкладка текстового вывода
        self.text_output = QTextEdit()
        self.text_output.setReadOnly(True)
        self.text_output.setFont(QFont("Consolas", 10))
        self.result_tabs.addTab(self.text_output, "Текстовый вывод")

        # Вкладка статистики
        self.stats_output = QTextEdit()
        self.stats_output.setReadOnly(True)
        self.stats_output.setFont(QFont("Consolas", 10))
        self.result_tabs.addTab(self.stats_output, "Статистика")

        layout.addWidget(self.result_tabs, 1)

        self.info_label = QLabel("Готово к работе. Выберите PBIT-файл и нажмите 'Начать анализ'.")
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet("padding: 6px; color: #333;")
        layout.addWidget(self.info_label)
        return panel

    def _create_log_panel(self):
        panel = QGroupBox("Лог работы")
        layout = QVBoxLayout(panel)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 10))
        self.log_text.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4;")
        layout.addWidget(self.log_text)
        btn_clear_log = QPushButton("Очистить лог")
        btn_clear_log.clicked.connect(lambda: self.log_text.clear())
        layout.addWidget(btn_clear_log)
        return panel

    def _select_pbit_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите PBIT файл", "", "PBIT файлы (*.pbit);;Все файлы (*.*)"
        )
        if file_path:
            self.pbit_path = file_path
            self.pbit_path_edit.setText(file_path)
            self.pbit_path_edit.setStyleSheet("color: black; padding: 5px; font-weight: bold;")
            if not self.output_dir:
                self.output_dir = os.path.dirname(file_path)
                self.output_dir_edit.setText(self.output_dir)
            self.log(f"Выбран файл: {file_path}")
            self.statusBar().showMessage(f"Файл загружен: {os.path.basename(file_path)}")

    def _select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "Выберите папку для сохранения результатов", self.output_dir
        )
        if dir_path:
            self.output_dir = dir_path
            self.output_dir_edit.setText(dir_path)
            self.log(f"Папка вывода: {dir_path}")

    def _start_analysis(self):
        if not self.pbit_path:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите PBIT файл!")
            return

        self._ensure_output_dir()
        self._set_analysis_ui_state(is_running=True)
        self.lineage_focus_combo.setEnabled(False)
        self.btn_find_lineage.setEnabled(False)
        self.btn_clear_filter.setEnabled(False)
        self.lineage_focus_combo.clear()
        self.lineage_focus_map = {}
        if hasattr(self, '_original_dfs'):
            self._original_dfs = None

        keep_temp = self.chk_keep_temp.isChecked()
        create_excel = self.chk_create_excel.isChecked()
        self.worker = AnalysisWorker(
            self.pbit_path,
            self.output_dir,
            keep_temp=keep_temp,
            create_excel=create_excel,
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_analysis_finished)
        self.worker.error.connect(self._on_analysis_error)
        self.worker.log.connect(self.log)
        self.worker.start()
        self.log("=" * 50)
        self.log("ЗАПУСК АНАЛИЗА")

    def _on_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.statusBar().showMessage(message)

    def _on_analysis_finished(self, dfs):
        self.current_dfs = dfs
        self._original_dfs = dfs.copy()  # сохраняем оригинальные данные для фильтрации
        self._set_analysis_ui_state(is_running=False)
        has_excel = bool(dfs.get("excel_path"))
        self.btn_open_excel.setEnabled(has_excel)
        self.btn_open_folder.setEnabled(True)
        self.btn_export_excel.setEnabled(True)
        self._populate_tables(dfs)
        self._update_text_report(dfs)
        self._update_stats_report(dfs)
        self._populate_lineage_focus_options()

        excel_path = dfs.get('excel_path')
        if excel_path:
            self.log(f"АНАЛИЗ ЗАВЕРШЕН! Файл: {excel_path}")
            finish_text = f"Анализ завершен успешно!\n\nОтчет сохранен:\n{excel_path}"
        else:
            self.log("АНАЛИЗ ЗАВЕРШЕН! Excel не создавался (по настройке).")
            finish_text = "Анализ завершен успешно!\n\nExcel не создавался (галка 'Создать Excel' выключена)."
        self.statusBar().showMessage("Анализ завершен успешно!")
        self.info_label.setText("Анализ завершен успешно. Просмотрите вкладки 'Результаты парсинга', 'Текстовый вывод' и 'Статистика'.")
        QMessageBox.information(self, "Готово!", finish_text)

        if self.chk_open_folder.isChecked():
            self._open_output_folder()
        if self.chk_open_excel.isChecked() and has_excel:
            self._open_excel()

    @staticmethod
    def _display_focus_id(node_type: str, node_id: str) -> str:
        if node_type == "Визуал":
            visual_id = node_id[4:] if node_id.startswith("VIS:") else node_id
            return f"VIS_ID :{visual_id}"
        return str(node_id).upper()

    @staticmethod
    def _normalize_focus_input(value: str) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""

        upper_raw = raw.upper()
        if upper_raw.startswith("VIS_ID"):
            match = re.match(r"VIS_ID\s*:\s*(.+)$", raw, re.IGNORECASE)
            return f"VIS:{match.group(1).strip()}" if match else ""
        if upper_raw.startswith("VIS:"):
            return f"VIS:{raw[4:].strip()}"
        if "|" in raw:
            parts = [part.strip() for part in raw.split("|")]
            for part in parts:
                normalized_part = PBITAnalyzerMainWindow._normalize_focus_input(part)
                if normalized_part:
                    return normalized_part
            return ""
        if upper_raw.startswith("S"):
            return f"s{raw[1:]}"
        if upper_raw.startswith("A"):
            return f"A{raw[1:]}"
        if upper_raw.startswith("M"):
            return f"M{raw[1:]}"
        return raw

    @staticmethod
    def _parse_focus_label(label: str) -> tuple[str, str, str]:
        """
        Парсит label формата "Тип | ID | Имя".
        Возвращает (node_type, node_id, display_name).
        Если формат не соответствует, пытается определить тип по ID.
        """
        label = label.strip()
        if " | " in label:
            parts = [p.strip() for p in label.split(" | ", 2)]
            if len(parts) == 3:
                node_type, node_id, display_name = parts
                return node_type, node_id, display_name
        # Попробуем определить тип по ID
        normalized = PBITAnalyzerMainWindow._normalize_focus_input(label)
        if normalized.startswith("VIS:"):
            return "Визуал", normalized, ""
        if normalized.startswith("s"):
            return "Таблица", normalized, ""
        if normalized.startswith("A"):
            return "Столбец", normalized, ""
        if normalized.startswith("M"):
            return "Мера", normalized, ""
        # По умолчанию считаем, что это ID без типа
        return "", normalized, ""

    def _populate_lineage_focus_options(self):
        self.lineage_focus_map = {}
        self.lineage_focus_combo.blockSignals(True)
        self.lineage_focus_combo.clear()
        self.lineage_focus_combo.blockSignals(False)

        edges_df = self.current_dfs.get("lineage")
        if edges_df is None or edges_df.empty:
            self.lineage_focus_combo.setEnabled(False)
            self.btn_find_lineage.setEnabled(False)
            return

        tables_df = self.current_dfs.get("tables_ref")
        columns_df = self.current_dfs.get("columns")
        measures_df = self.current_dfs.get("measures")
        visuals_df = self.current_dfs.get("visuals")

        table_name_by_id = {}
        if tables_df is not None and not tables_df.empty:
            for _, table_row in tables_df.iterrows():
                table_id = str(table_row.get("№ таблицы", "")).strip()
                table_name = str(table_row.get("Имя таблицы", "")).strip()
                if table_id:
                    table_name_by_id[table_id] = table_name or "—"

        column_name_by_id = {}
        if columns_df is not None and not columns_df.empty:
            for _, column_row in columns_df.iterrows():
                column_id = str(column_row.get("№ столбца", "")).strip()
                column_name = str(column_row.get("Столбец", "")).strip()
                if column_id:
                    column_name_by_id[column_id] = column_name or "—"

        measure_name_by_id = {}
        if measures_df is not None and not measures_df.empty:
            for _, measure_row in measures_df.iterrows():
                measure_id = str(measure_row.get("№ меры", "")).strip()
                measure_name = str(measure_row.get("Мера", "")).strip()
                if measure_id:
                    measure_name_by_id[measure_id] = measure_name or "—"

        visual_name_by_id = {}
        visual_type_by_id = {}
        if visuals_df is not None and not visuals_df.empty:
            for _, visual_row in visuals_df.iterrows():
                visual_id_raw = str(visual_row.get("ID визуала", "")).strip()
                visual_type = str(visual_row.get("Тип визуала", "")).strip() or "—"
                page_name = str(visual_row.get("Имя листа", "")).strip()
                if not visual_id_raw:
                    continue
                match = re.match(r"VIS_ID\s*:\s*(.+)$", visual_id_raw, re.IGNORECASE)
                visual_key = f"VIS:{match.group(1).strip()}" if match else visual_id_raw
                visual_type_by_id[visual_key] = visual_type
                if page_name:
                    visual_name_by_id[visual_key] = page_name

        nodes = {}
        for _, row in edges_df.iterrows():
            for type_col, id_col, name_col in (
                ("Тип источника", "ID источника", "Имя источника"),
                ("Тип приемника", "ID приемника", "Имя приемника"),
            ):
                node_type = str(row.get(type_col, "")).strip()
                node_id = str(row.get(id_col, "")).strip()
                if node_type not in {"Таблица", "Столбец", "Мера", "Визуал"} or not node_id:
                    continue
                nodes[node_id] = node_type

        for node_id, node_type in sorted(nodes.items(), key=lambda item: (item[1], item[0])):
            display_id = self._display_focus_id(node_type, node_id)
            if node_type == "Визуал":
                visual_type = visual_type_by_id.get(node_id, "—")
                label = f"{node_type} | {display_id} | {visual_type}"
            else:
                if node_type == "Таблица":
                    display_name = table_name_by_id.get(node_id, "—")
                elif node_type == "Столбец":
                    display_name = column_name_by_id.get(node_id, "—")
                else:
                    display_name = measure_name_by_id.get(node_id, "—")
                label = f"{node_type} | {display_id} | {display_name}"
            self.lineage_focus_combo.addItem(label, userData=node_id)
            self.lineage_focus_map[label] = node_id

        # Настройка completer с фильтром подстроки
        completer = QCompleter()
        completer.setModel(self.lineage_focus_combo.model())
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.lineage_focus_combo.setCompleter(completer)

        has_options = self.lineage_focus_combo.count() > 0
        self.lineage_focus_combo.setEnabled(has_options)
        self.btn_find_lineage.setEnabled(has_options)

    def _apply_lineage_focus(self):
        if not self.current_dfs:
            QMessageBox.warning(self, "Ошибка", "Сначала выполните анализ.")
            return

        raw_value = self.lineage_focus_combo.currentText().strip()
        if not raw_value:
            QMessageBox.warning(self, "Ошибка", "Выберите или введите фокус для lineage.")
            return

        focus_id = self.lineage_focus_combo.currentData()
        if not focus_id:
            focus_id = self.lineage_focus_map.get(raw_value) or self._normalize_focus_input(raw_value)

        # Определяем тип объекта
        node_type, parsed_id, _ = self._parse_focus_label(raw_value)
        if not node_type and parsed_id:
            # Если тип не определился, попробуем определить по focus_id
            if focus_id.startswith("s"):
                node_type = "Таблица"
            elif focus_id.startswith("A"):
                node_type = "Столбец"
            elif focus_id.startswith("M"):
                node_type = "Мера"
            elif focus_id.startswith("VIS:"):
                node_type = "Визуал"
            else:
                node_type = ""

        # Сохраняем оригинальные данные, если ещё не сохранены
        if not hasattr(self, '_original_dfs'):
            self._original_dfs = self.current_dfs.copy()

        # Вычисляем ID визуалов, связанных с focus_id через lineage
        visual_ids = set()
        lineage_df = self._original_dfs.get("lineage")
        if isinstance(lineage_df, pd.DataFrame) and not lineage_df.empty:
            if "ID источника" in lineage_df.columns and "ID приемника" in lineage_df.columns:
                # Находим строки, где focus_id является источником или приемником
                mask = (lineage_df["ID источника"].astype(str).str.upper() == focus_id.upper()) | \
                       (lineage_df["ID приемника"].astype(str).str.upper() == focus_id.upper())
                related_rows = lineage_df[mask]
                for _, row in related_rows.iterrows():
                    src_type = str(row.get("Тип источника", "")).strip()
                    src_id = str(row.get("ID источника", "")).strip()
                    dst_type = str(row.get("Тип приемника", "")).strip()
                    dst_id = str(row.get("ID приемника", "")).strip()
                    if src_type == "Визуал" and src_id:
                        visual_ids.add(src_id)
                    if dst_type == "Визуал" and dst_id:
                        visual_ids.add(dst_id)

        # Фильтрация всех таблиц по focus_id и node_type
        filtered_dfs = {}
        for key, df in self._original_dfs.items():
            # Пропускаем не-DataFrame значения (например, строки)
            if not isinstance(df, pd.DataFrame):
                filtered_dfs[key] = df
                continue
            if df is None or df.empty:
                filtered_dfs[key] = df
                continue
            if key == "measures":
                # Фильтруем меры по № меры
                if "№ меры" in df.columns:
                    filtered = df[df["№ меры"].astype(str).str.upper() == focus_id.upper()]
                    filtered_dfs[key] = filtered
                else:
                    filtered_dfs[key] = df
            elif key == "tables_ref":
                # Находим номер таблицы для меры (если фокус - мера)
                if node_type == "Мера" and "measures" in self._original_dfs:
                    measures_df = self._original_dfs["measures"]
                    if isinstance(measures_df, pd.DataFrame) and not measures_df.empty and "№ таблицы" in measures_df.columns:
                        table_num = measures_df.loc[measures_df["№ меры"].astype(str).str.upper() == focus_id.upper(), "№ таблицы"].iloc[0] if not measures_df.empty else None
                        if table_num is not None:
                            filtered = df[df["№ таблицы"].astype(str) == str(table_num)]
                            filtered_dfs[key] = filtered
                        else:
                            filtered_dfs[key] = pd.DataFrame(columns=df.columns)
                    else:
                        filtered_dfs[key] = pd.DataFrame(columns=df.columns)
                else:
                    # Для других типов оставляем пустым
                    filtered_dfs[key] = pd.DataFrame(columns=df.columns)
            elif key == "visuals":
                # Фильтруем визуалы, которые используют focus_id (через lineage)
                if "ID визуала" in df.columns:
                    # Приводим ID визуала к формату VIS:...
                    def normalize_visual_id(vid):
                        vid_str = str(vid).strip()
                        match = re.match(r"VIS_ID\s*:\s*(.+)$", vid_str, re.IGNORECASE)
                        if match:
                            return f"VIS:{match.group(1).strip()}"
                        return vid_str
                    df_normalized_ids = df["ID визуала"].apply(normalize_visual_id)
                    filtered = df[df_normalized_ids.isin(visual_ids)]
                    filtered_dfs[key] = filtered
                else:
                    filtered_dfs[key] = pd.DataFrame(columns=df.columns)
            elif key == "lineage":
                # Фильтруем lineage где ID источника или приемника равен focus_id
                if "ID источника" in df.columns and "ID приемника" in df.columns:
                    filtered = df[(df["ID источника"].astype(str).str.upper() == focus_id.upper()) |
                                  (df["ID приемника"].astype(str).str.upper() == focus_id.upper())]
                    filtered_dfs[key] = filtered
                else:
                    filtered_dfs[key] = df
            elif key in ("lineage_upstream", "lineage_downstream"):
                # Эти листы будут построены ниже
                continue
            else:
                # Остальные листы оставляем пустыми
                filtered_dfs[key] = pd.DataFrame(columns=df.columns)

        # Строим upstream/downstream на основе отфильтрованного lineage
        edges_df = filtered_dfs.get("lineage")
        upstream_df, downstream_df = build_lineage_directions(edges_df, focus_id=focus_id)
        filtered_dfs["lineage_upstream"] = upstream_df
        filtered_dfs["lineage_downstream"] = downstream_df

        # Обновляем current_dfs отфильтрованными данными
        self.current_dfs = filtered_dfs
        self._populate_tables(self.current_dfs)
        self.statusBar().showMessage(f"Фильтр применен: {raw_value}")
        self.log(f"Фильтр: {raw_value} -> upstream={len(upstream_df)}, downstream={len(downstream_df)}")
        self.btn_clear_filter.setEnabled(True)

    def _clear_lineage_filter(self):
        """Очистить примененный фильтр и показать все данные."""
        if not hasattr(self, '_original_dfs') or self._original_dfs is None:
            QMessageBox.information(self, "Информация", "Фильтр не применялся.")
            return
        
        # Восстанавливаем оригинальные данные
        self.current_dfs = self._original_dfs.copy()
        # Очищаем выбор в комбобоксе
        self.lineage_focus_combo.setCurrentIndex(-1)
        self.lineage_focus_combo.setCurrentText("")
        # Обновляем таблицы
        self._populate_tables(self.current_dfs)
        self.statusBar().showMessage("Фильтр очищен, показаны все данные.")
        self.log("Фильтр lineage очищен.")
        self.btn_clear_filter.setEnabled(False)

    def _on_analysis_error(self, error_msg):
        self._set_analysis_ui_state(is_running=False)
        self.log(f"❌ ОШИБКА: {error_msg}")
        self.statusBar().showMessage("Ошибка анализа!")
        self.info_label.setText(f"Ошибка анализа: {error_msg}")
        QMessageBox.critical(self, "Ошибка", f"Произошла ошибка:\n\n{error_msg}")

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def _populate_tables(self, dfs):
        while self.tab_widget.count() > 0:
            self.tab_widget.removeTab(0)

        for df_key, tab_name in TAB_MAPPING:
            df = dfs.get(df_key)
            if df is None or df.empty:
                continue
            table = self._create_table_from_df(df, df_key=df_key)
            self.tab_widget.addTab(table, tab_name)

        self.log(f"Создано вкладок: {self.tab_widget.count()}")

    def _update_text_report(self, dfs):
        lines = [
            "=== РЕЗУЛЬТАТЫ АНАЛИЗА PBIT ===",
            f"Файл: {self.pbit_path or '—'}",
            f"Папка вывода: {self.output_dir or '—'}",
            f"Дата анализа: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "=== ДОСТУПНЫЕ НАБОРЫ ДАННЫХ ===",
        ]
        for df_key, tab_name in TAB_MAPPING:
            df = dfs.get(df_key)
            if df is not None and not df.empty:
                lines.append(f"- {tab_name}: {len(df)} строк")
        excel_path = dfs.get("excel_path")
        lines.extend(["", f"Excel: {excel_path if excel_path else 'не создавался'}"])
        self.text_output.setPlainText("\n".join(lines))

    def _update_stats_report(self, dfs):
        stats_lines = [
            "=== СТАТИСТИКА ===",
            f"Всего вкладок с данными: {sum(1 for key, _ in TAB_MAPPING if dfs.get(key) is not None and not dfs.get(key).empty)}",
        ]
        total_rows = 0
        for df_key, tab_name in TAB_MAPPING:
            df = dfs.get(df_key)
            if df is None or df.empty:
                continue
            row_count = len(df)
            total_rows += row_count
            stats_lines.append(f"{tab_name}: {row_count}")
        stats_lines.extend(["", f"Всего строк по всем наборам: {total_rows}"])
        self.stats_output.setPlainText("\n".join(stats_lines))

    def _create_table_from_df(self, df, df_key: str = ""):
        table = CustomTableWidget()
        table.setRowCount(len(df))
        table.setColumnCount(len(df.columns))
        table.setHorizontalHeaderLabels(df.columns.tolist())

        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                value = df.iloc[row_idx, col_idx]
                if pd.isna(value):
                    text = ""
                elif isinstance(value, str) and '\n' in value:
                    text = value[:200] + ('...' if len(value) > 200 else '')
                else:
                    text = str(value)
                item = QTableWidgetItem(text)
                item.setToolTip(str(value) if len(str(value)) > 100 else "")
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
            self.log(
                f"Таблица {df_key} большая ({len(df)} строк), "
                "использован облегченный режим ширины колонок."
            )
        else:
            header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
            table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        return table

    def _open_excel(self):
        excel_path = self.current_dfs.get('excel_path')
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "Ошибка", "Excel файл не найден. Сначала выполните анализ.")
            return
        try:
            open_path_in_os(excel_path)
            self.log(f"Открыт Excel файл: {excel_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл:\n{e}")

    def _open_output_folder(self):
        if not os.path.exists(self.output_dir):
            QMessageBox.warning(self, "Ошибка", "Папка результатов не существует.")
            return
        try:
            open_path_in_os(self.output_dir)
            self.log(f"Открыта папка: {self.output_dir}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть папку:\n{e}")

    def _export_to_excel(self):
        """Экспортировать текущие данные в новый Excel файл."""
        if not self.current_dfs:
            QMessageBox.warning(self, "Ошибка", "Нет данных для экспорта. Сначала выполните анализ.")
            return
        
        # Предложить путь для сохранения
        default_name = f"pbit_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        if self.pbit_path:
            source_name = os.path.splitext(os.path.basename(self.pbit_path))[0]
            safe_name = re.sub(r'[<>:"/\\|?*]+', "_", source_name).strip(" .")
            default_name = f"pbi_report_{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        default_dir = self.output_dir if self.output_dir else os.path.expanduser("~")
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel файл",
            os.path.join(default_dir, default_name),
            "Excel файлы (*.xlsx);;Все файлы (*.*)"
        )
        if not file_path:
            return  # пользователь отменил
        
        try:
            self.log(f"Начало экспорта в Excel: {file_path}")
            generator = ExcelReportGenerator(file_path, self.current_dfs, source_file=self.pbit_path or "")
            generator.generate()
            self.log(f"Экспорт завершен: {file_path}")
            QMessageBox.information(self, "Успех", f"Файл успешно сохранен:\n{file_path}")
            # Обновить excel_path в current_dfs, чтобы кнопка "Открыть Excel" работала
            self.current_dfs['excel_path'] = file_path
            self.btn_open_excel.setEnabled(True)
        except Exception as e:
            self.log(f"Ошибка экспорта: {traceback.format_exc()}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{e}")

    def _show_help(self):
        dialog = HelpDialog(self)
        dialog.exec()


# ============================================================
# Главная точка входа
# ============================================================

