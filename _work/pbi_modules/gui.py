import os
import re
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QCompleter,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .excel_report import ExcelReportGenerator
from .file_kind import PBIT, PBIX, FileFormatError, detect_file_kind
from .focus import display_focus_id, normalize_focus_id, parse_focus_label
from .help_dialog import HelpDialog
from .schemas import TABS_BY_KIND, tabs_for
from .ui.lineage_panel import filter_dfs
from .ui.search import SearchBar
from .ui.tables import create_table_from_df, format_cell  # noqa: F401 — реэкспорт
from .ui.worker import AnalysisWorker


def _load_stylesheet() -> str:
    qss_path = Path(__file__).parent / "styles.qss"
    if qss_path.exists():
        return qss_path.read_text(encoding="utf-8")
    return ""

# Совместимость: наборы вкладок берутся из единого реестра schemas.py.
TAB_MAPPING = list(TABS_BY_KIND[PBIT])
TAB_MAPPING_PBIX = list(TABS_BY_KIND[PBIX])


def tab_mapping_for(file_kind: str):
    """Набор вкладок с таблицами по типу файла."""
    return list(tabs_for(file_kind))


def open_path_in_os(path: str) -> None:
    if sys.platform == 'win32':
        os.startfile(path)
        return
    opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
    subprocess.run([opener, path], check=False)


# ============================================================
# Custom Table Widget с контекстным меню
# ============================================================

# ============================================================
# Главное окно приложения
# ============================================================

class PBITAnalyzerMainWindow(QMainWindow):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()
        self.current_dfs = {}
        self._original_dfs = {}
        self.lineage_focus_map = {}
        self.output_dir = ""
        self.pbit_path = ""
        self.file_kind = PBIT
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

    # ============================================================
    # Keyboard shortcuts (Task 2.1)
    # ============================================================

    def _setup_shortcuts(self):
        shortcut_open = QShortcut(QKeySequence("Ctrl+O"), self)
        shortcut_open.activated.connect(self._select_pbit_file)

        shortcut_save = QShortcut(QKeySequence("Ctrl+S"), self)
        shortcut_save.activated.connect(self._export_to_excel)

        shortcut_help = QShortcut(QKeySequence("Ctrl+H"), self)
        shortcut_help.activated.connect(self._show_help)

        shortcut_quit = QShortcut(QKeySequence("Ctrl+Q"), self)
        shortcut_quit.activated.connect(self.close)

    # ============================================================
    # Drag-and-drop (Task 2.2)
    # ============================================================

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith((".pbit", ".pbix")):
                    event.acceptProposedAction()
                    return

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith((".pbit", ".pbix")):
                self._set_selected_file(file_path)
                self.log(f"Файл перетащен: {file_path}")
                self.statusBar().showMessage(f"Файл загружен: {os.path.basename(file_path)}")
                return

    # ============================================================
    # Search / filter in tabs (Task 2.3)
    # ============================================================

    def init_ui(self):
        self.setWindowTitle("PBIT Analyzer — Анализатор Power BI Template (PBIT) и моделей (PBIX)")
        self.setGeometry(100, 100, 1400, 900)

        # Load QSS stylesheet
        stylesheet = _load_stylesheet()
        if stylesheet:
            self.setStyleSheet(stylesheet)

        # Enable drag-and-drop for .pbit files
        self.setAcceptDrops(True)

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

        self.statusBar().showMessage("Готово к работе. Выберите файл PBIT или PBIX.")
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(300)
        self.progress_bar.setVisible(False)
        self.statusBar().addPermanentWidget(self.progress_bar)

        # --- Keyboard shortcuts ---
        self._setup_shortcuts()

    def _create_header_panel(self):
        panel = QWidget()
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        title_label = QLabel("PBIT Analyzer")
        title_label.setObjectName("headerTitle")
        layout.addWidget(title_label)
        layout.addStretch()
        help_button = QPushButton("Справка")
        help_button.setObjectName("btnHelp")
        help_button.clicked.connect(self._show_help)
        layout.addWidget(help_button)
        return panel

    def _create_top_panel(self):
        panel = QGroupBox("Файл и настройки")
        layout = QGridLayout(panel)
        layout.setSpacing(5)

        layout.addWidget(QLabel("PBIT/PBIX файл:"), 0, 0)
        self.pbit_path_edit = QLabel("Файл не выбран")
        self.pbit_path_edit.setObjectName("pbitPathLabel")
        layout.addWidget(self.pbit_path_edit, 0, 1)
        btn_select_file = QPushButton("Выбрать PBIT/PBIX")
        btn_select_file.setObjectName("btnSelectFile")
        btn_select_file.clicked.connect(self._select_pbit_file)
        layout.addWidget(btn_select_file, 0, 2)

        layout.addWidget(QLabel("Папка вывода:"), 1, 0)
        self.output_dir_edit = QLabel(self.output_dir)
        self.output_dir_edit.setObjectName("outputDirLabel")
        layout.addWidget(self.output_dir_edit, 1, 1)
        btn_select_dir = QPushButton("Выбрать папку")
        btn_select_dir.setObjectName("btnSelectDir")
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
        self.lineage_focus_combo.setPlaceholderText("После анализа: S4 / A105 / M3 / VIS_ID :... (визуалы — только PBIT)")
        self.lineage_focus_combo.setToolTip("Выберите объект из списка или введите вручную")
        self.lineage_focus_combo.setEnabled(False)
        lineage_layout.addWidget(self.lineage_focus_combo, 1)
        self.btn_find_lineage = QPushButton("Найти")
        self.btn_find_lineage.setObjectName("btnFindLineage")
        self.btn_find_lineage.setEnabled(False)
        self.btn_find_lineage.clicked.connect(self._apply_lineage_focus)
        lineage_layout.addWidget(self.btn_find_lineage)
        self.btn_clear_filter = QPushButton("Очистить фильтр")
        self.btn_clear_filter.setObjectName("btnClearFilter")
        self.btn_clear_filter.setEnabled(False)
        self.btn_clear_filter.clicked.connect(self._clear_lineage_filter)
        lineage_layout.addWidget(self.btn_clear_filter)
        layout.addLayout(lineage_layout, 3, 0, 1, 3)

        return panel

    def _create_left_panel(self):
        panel = QGroupBox("Действия")
        layout = QVBoxLayout(panel)
        layout.setSpacing(10)

        self.btn_analyze = QPushButton("Начать анализ")
        self.btn_analyze.setObjectName("btnAnalyze")
        self.btn_analyze.clicked.connect(self._start_analysis)
        layout.addWidget(self.btn_analyze)

        self.btn_export_excel = QPushButton("Выгрузить в Excel")
        self.btn_export_excel.setObjectName("btnExportExcel")
        self.btn_export_excel.clicked.connect(self._export_to_excel)
        self.btn_export_excel.setEnabled(False)
        layout.addWidget(self.btn_export_excel)

        self.btn_open_excel = QPushButton("Открыть Excel")
        self.btn_open_excel.setObjectName("btnOpenExcel")
        self.btn_open_excel.clicked.connect(self._open_excel)
        self.btn_open_excel.setEnabled(False)
        layout.addWidget(self.btn_open_excel)

        self.btn_open_folder = QPushButton("Открыть папку")
        self.btn_open_folder.setObjectName("btnOpenFolder")
        self.btn_open_folder.clicked.connect(self._open_output_folder)
        self.btn_open_folder.setEnabled(False)
        layout.addWidget(self.btn_open_folder)

        layout.addStretch()

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

        self.info_label = QLabel("Готово к работе. Выберите файл PBIT или PBIX и нажмите 'Начать анализ'.")
        self.info_label.setObjectName("infoLabel")
        self.info_label.setWordWrap(True)
        layout.addWidget(self.info_label)
        return panel

    def _create_log_panel(self):
        panel = QGroupBox("Лог работы")
        layout = QVBoxLayout(panel)
        self.log_text = QTextEdit()
        self.log_text.setObjectName("logText")
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        btn_clear_log = QPushButton("Очистить лог")
        btn_clear_log.clicked.connect(lambda: self.log_text.clear())
        layout.addWidget(btn_clear_log)
        return panel

    def _set_selected_file(self, file_path: str) -> None:
        """Выбирает файл и определяет его формат по расширению и содержимому."""
        try:
            kind = detect_file_kind(file_path)
        except FileFormatError as error:
            QMessageBox.warning(self, "Неподдерживаемый файл", str(error))
            self.log(f"Файл не принят: {error}")
            return

        self.pbit_path = file_path
        self.file_kind = kind
        self.pbit_path_edit.setText(file_path)
        self.pbit_path_edit.setObjectName("pbitPathLabelActive")
        self.pbit_path_edit.setStyleSheet("color: black; padding: 5px; font-weight: bold;")
        if not self.output_dir:
            self.output_dir = os.path.dirname(file_path)
            self.output_dir_edit.setText(self.output_dir)
        kind_label = "PBIX (модель)" if kind == PBIX else "PBIT (шаблон + отчёт)"
        self.log(f"Выбран файл [{kind_label}]: {file_path}")
        self.statusBar().showMessage(
            f"Файл загружен: {os.path.basename(file_path)} ({kind.upper()})"
        )

    def _select_pbit_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл PBIT или PBIX",
            "",
            "Power BI (*.pbit *.pbix);;PBIT (*.pbit);;PBIX (*.pbix);;Все файлы (*.*)",
        )
        if file_path:
            self._set_selected_file(file_path)

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
            QMessageBox.warning(self, "Ошибка", "Сначала выберите файл PBIT или PBIX!")
            return

        self._ensure_output_dir()
        self._set_analysis_ui_state(is_running=True)
        self.lineage_focus_combo.setEnabled(False)
        self.btn_find_lineage.setEnabled(False)
        self.btn_clear_filter.setEnabled(False)
        self.lineage_focus_combo.clear()
        self.lineage_focus_map = {}
        self._original_dfs = {}

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
        self.file_kind = str(dfs.get("file_kind", self.file_kind)).lower()
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
        self.info_label.setText(
            f"Анализ завершен успешно ({self.file_kind.upper()}). "
            "Просмотрите вкладки 'Результаты парсинга', 'Текстовый вывод' и 'Статистика'."
        )
        self._show_result_message("Готово!", finish_text, excel_path or "")

        if self.chk_open_folder.isChecked():
            self._open_output_folder()
        if self.chk_open_excel.isChecked() and has_excel:
            self._open_excel()

    @staticmethod
    def _display_focus_id(node_type: str, node_id: str) -> str:
        """Совместимость: логика в focus.display_focus_id."""
        return display_focus_id(node_type, node_id)

    @staticmethod
    def _normalize_focus_input(value: str) -> str:
        """Совместимость: логика в focus.normalize_focus_id."""
        return normalize_focus_id(value)

    @staticmethod
    def _parse_focus_label(label: str) -> tuple[str, str, str]:
        """Совместимость: логика в focus.parse_focus_label."""
        return parse_focus_label(label)

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
        """Фильтрует результаты по выбранному объекту через чистую filter_dfs."""
        if not self.current_dfs:
            QMessageBox.warning(self, "Ошибка", "Сначала выполните анализ.")
            return

        raw_value = self.lineage_focus_combo.currentText().strip()
        if not raw_value:
            QMessageBox.warning(self, "Ошибка", "Выберите или введите фокус для lineage.")
            return

        focus_id = (
            self.lineage_focus_combo.currentData()
            or self.lineage_focus_map.get(raw_value)
            or self._normalize_focus_input(raw_value)
        )
        node_type, _, _ = self._parse_focus_label(raw_value)
        if not focus_id:
            QMessageBox.warning(self, "Ошибка", "Не удалось распознать ID объекта.")
            return

        if not self._original_dfs:
            self._original_dfs = self.current_dfs.copy()

        self.current_dfs = filter_dfs(self._original_dfs, focus_id, node_type)
        self._populate_tables(self.current_dfs)
        self.statusBar().showMessage(f"Фильтр применен: {raw_value}")
        self.log(
            f"Фильтр: {raw_value} -> "
            f"upstream={len(self.current_dfs['lineage_upstream'])}, "
            f"downstream={len(self.current_dfs['lineage_downstream'])}"
        )
        self.btn_clear_filter.setEnabled(True)

    def _clear_lineage_filter(self):
        """Очистить примененный фильтр и показать все данные."""
        if not self._original_dfs:
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

        file_kind = dfs.get("file_kind", self.file_kind)
        for df_key, tab_name in tab_mapping_for(file_kind):
            df = dfs.get(df_key)
            if df is None or df.empty:
                # В PBIX показываем все разделы: пустой — с пометкой, а не спрятанный.
                if str(file_kind).lower() == PBIX:
                    self.tab_widget.addTab(self._create_empty_tab(tab_name), tab_name)
                continue
            table = create_table_from_df(df, df_key=df_key, log=self.log)

            # Wrap table with search bar (Task 2.3)
            tab_container = QWidget()
            tab_layout = QVBoxLayout(tab_container)
            tab_layout.setContentsMargins(0, 0, 0, 0)
            tab_layout.setSpacing(2)
            search_bar = SearchBar(table)
            tab_layout.addWidget(search_bar)
            tab_layout.addWidget(table, 1)

            self.tab_widget.addTab(tab_container, tab_name)

        self.log(f"Создано вкладок: {self.tab_widget.count()}")

    @staticmethod
    def _create_empty_tab(tab_name: str):
        """Заглушка для раздела PBIX без данных."""
        container = QWidget()
        layout = QVBoxLayout(container)
        label = QLabel(f"Нет данных: {tab_name}")
        label.setObjectName("emptyTabLabel")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
        return container

    def _update_text_report(self, dfs):
        kind = str(dfs.get("file_kind", self.file_kind)).upper()
        lines = [
            f"=== РЕЗУЛЬТАТЫ АНАЛИЗА {kind} ===",
            f"Файл: {self.pbit_path or '—'}",
            f"Папка вывода: {self.output_dir or '—'}",
            f"Дата анализа: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "=== ДОСТУПНЫЕ НАБОРЫ ДАННЫХ ===",
        ]
        for df_key, tab_name in tab_mapping_for(kind):
            df = dfs.get(df_key)
            if df is not None and not df.empty:
                lines.append(f"- {tab_name}: {len(df)} строк")
        excel_path = dfs.get("excel_path")
        lines.extend(["", f"Excel: {excel_path if excel_path else 'не создавался'}"])
        self.text_output.setPlainText("\n".join(lines))

    def _update_stats_report(self, dfs):
        mapping = tab_mapping_for(dfs.get("file_kind", self.file_kind))
        stats_lines = [
            "=== СТАТИСТИКА ===",
            f"Всего вкладок с данными: {sum(1 for key, _ in mapping if dfs.get(key) is not None and not dfs.get(key).empty)}",
        ]
        total_rows = 0
        for df_key, tab_name in mapping:
            df = dfs.get(df_key)
            if df is None or df.empty:
                continue
            row_count = len(df)
            total_rows += row_count
            stats_lines.append(f"{tab_name}: {row_count}")
        stats_lines.extend(["", f"Всего строк по всем наборам: {total_rows}"])
        self.stats_output.setPlainText("\n".join(stats_lines))

    def _show_result_message(self, title: str, text: str, excel_path: str = "") -> None:
        """Итоговое окно с кнопками «ОК» и «Открыть Excel» (если файл создан)."""
        dialog = QMessageBox(self)
        dialog.setWindowTitle(title)
        dialog.setIcon(QMessageBox.Icon.Information)
        dialog.setText(text)
        ok_button = dialog.addButton("ОК", QMessageBox.ButtonRole.AcceptRole)
        dialog.setDefaultButton(ok_button)
        open_button = None
        if excel_path and os.path.exists(excel_path):
            open_button = dialog.addButton("Открыть Excel", QMessageBox.ButtonRole.ActionRole)
        dialog.exec()
        if open_button is not None and dialog.clickedButton() is open_button:
            self._open_excel_file(excel_path)

    def _open_excel_file(self, excel_path: str) -> None:
        try:
            open_path_in_os(excel_path)
            self.log(f"Открыт Excel файл: {excel_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл:\n{e}")

    def _open_excel(self):
        excel_path = self.current_dfs.get('excel_path')
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "Ошибка", "Excel файл не найден. Сначала выполните анализ.")
            return
        self._open_excel_file(excel_path)

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
            sheets = ExcelReportGenerator.sheets_for(self.current_dfs.get("file_kind", self.file_kind))
            generator = ExcelReportGenerator(
                file_path, self.current_dfs, source_file=self.pbit_path or "", sheets=sheets
            )
            generator.generate()
            self.log(f"Экспорт завершен: {file_path}")
            # Обновить excel_path в current_dfs, чтобы кнопка "Открыть Excel" работала
            self.current_dfs['excel_path'] = file_path
            self.btn_open_excel.setEnabled(True)
            self._show_result_message("Успех", f"Файл успешно сохранен:\n{file_path}", file_path)
        except Exception as e:
            self.log(f"Ошибка экспорта: {traceback.format_exc()}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{e}")

    def _show_help(self):
        dialog = HelpDialog(self)
        dialog.exec()


# ============================================================
# Главная точка входа
# ============================================================

