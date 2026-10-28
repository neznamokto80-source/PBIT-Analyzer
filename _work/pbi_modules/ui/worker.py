"""Фоновый поток анализа (не блокирует интерфейс)."""
import traceback

from PyQt6.QtCore import QThread, pyqtSignal

from ..dispatcher import analyze_file


class AnalysisWorker(QThread):
    """Фоновый поток для анализа PBIT/PBIX файла."""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    log = pyqtSignal(str)

    def __init__(self, file_path, output_dir, keep_temp=False, create_excel=False):
        super().__init__()
        self.file_path = file_path
        self.output_dir = output_dir
        self.keep_temp = keep_temp
        self.create_excel = create_excel

    def run(self):
        try:
            self.log.emit(f"Начало анализа: {self.file_path}")
            dfs = analyze_file(
                self.file_path,
                self.output_dir,
                keep_temp=self.keep_temp,
                create_excel=self.create_excel,
                progress_callback=lambda value, message: self.progress.emit(value, message),
                log_callback=lambda message: self.log.emit(message),
            )
            self.finished.emit(dfs)

        except Exception as e:
            self.error.emit(str(e))
            self.log.emit(f"ОШИБКА: {traceback.format_exc()}")
