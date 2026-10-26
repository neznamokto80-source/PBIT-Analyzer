"""
Анализ .pbix (семантическая модель) без GUI.

Использует PBIXRay: читает только модель (таблицы, столбцы, меры, DAX, M,
связи, RLS/OLS, статистику хранения). Слой отчёта PBIX недоступен — визуалов,
страниц и условного форматирования здесь нет.

Возвращает тот же словарь наборов данных, что и analyze_pbit_file(), плюс:
    file_kind   — "pbix"
    dax, power_query, m_parameters, rls, ols, perspectives, statistics
    warnings    — список проблем, не помешавших остальному анализу
"""
from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime

from .analyzer import build_report_filename
from .excel_report import ExcelReportGenerator
from .lineage import build_lineage_directions, build_lineage_edges
from .pbix_extractor import PBIXModelExtractor

DEFAULT_ON_DISK_THRESHOLD_MB = 500


class PBIXAnalysisError(Exception):
    """Ошибка чтения PBIX, которую CLI/GUI показывают как понятное сообщение."""


def _file_size_mb(path: str) -> float:
    try:
        return os.path.getsize(path) / (1024 * 1024)
    except OSError:
        return 0.0


def _open_model(path: str, on_disk: bool, temp_dir: str):
    """Открывает PBIXRay, переводя его исключения в PBIXAnalysisError."""
    try:
        from pbixray import LiveConnectionError, NoEmbeddedModelError, PBIXRay, PBIXRayError
    except ImportError as error:  # pragma: no cover — зависимость объявлена обязательной
        raise PBIXAnalysisError(
            "Не установлена библиотека pbixray. Установите её: pip install pbixray==0.15.5"
        ) from error

    try:
        return PBIXRay(path, on_disk=on_disk, temp_dir=temp_dir or None)
    except (LiveConnectionError, NoEmbeddedModelError) as error:
        raise PBIXAnalysisError(
            "В файле нет встроенной модели: отчёт подключён к внешней модели "
            "(Live Connection). Анализ таких файлов не поддерживается."
        ) from error
    except PBIXRayError as error:
        raise PBIXAnalysisError(f"Не удалось прочитать PBIX: {error}") from error
    except Exception as error:  # noqa: BLE001 — сообщаем по-русски, а не трейсбеком
        raise PBIXAnalysisError(f"Не удалось прочитать PBIX: {error}") from error


def analyze_pbix_file(
    pbix_path: str,
    output_dir: str,
    keep_temp: bool = False,
    lineage_focus: str = "",
    create_excel: bool = True,
    progress_callback=None,
    log_callback=None,
    on_disk: bool | None = None,
    on_disk_threshold_mb: float = DEFAULT_ON_DISK_THRESHOLD_MB,
) -> dict:
    if not os.path.exists(pbix_path):
        raise FileNotFoundError(f"Файл не найден: {pbix_path}")

    os.makedirs(output_dir, exist_ok=True)

    def _progress(value: int, message: str) -> None:
        if progress_callback:
            progress_callback(value, message)

    def _log(message: str) -> None:
        if log_callback:
            log_callback(message)

    size_mb = _file_size_mb(pbix_path)
    use_on_disk = on_disk if on_disk is not None else size_mb >= on_disk_threshold_mb
    if use_on_disk:
        temp_dir = tempfile.mkdtemp(prefix="_temp_pbix_", dir=output_dir)
    else:
        temp_dir = ""

    _progress(10, "Открытие PBIX...")
    _log(
        f"PBIX: {size_mb:.1f} МБ, режим чтения: "
        f"{'на диске (on_disk)' if use_on_disk else 'в памяти'}"
    )

    model = None
    try:
        model = _open_model(pbix_path, use_on_disk, temp_dir)

        _progress(30, "Извлечение модели данных...")
        extractor = PBIXModelExtractor(model)
        dfs = extractor.extract_all()
        for warning in dfs.get("warnings", []):
            _log(f"Предупреждение: {warning}")

        model_size = 0
        try:
            model_size = int(getattr(model, "size", 0) or 0)
        except (TypeError, ValueError):
            model_size = 0

        _progress(65, "Построение lineage...")
        dfs["lineage"] = build_lineage_edges(dfs)
        lineage_upstream, lineage_downstream = build_lineage_directions(
            dfs["lineage"],
            focus_id=lineage_focus,
        )
        dfs["lineage_upstream"] = lineage_upstream
        dfs["lineage_downstream"] = lineage_downstream

        _progress(80, "Формирование отчета...")
        warnings_text = "\n".join(dfs.get("warnings", [])) or "—"
        extra_stats = {
            "Тип файла": "PBIX (семантическая модель)",
            "Размер файла, МБ": f"{size_mb:.2f}",
            "Размер модели, байт": model_size,
            "Режим чтения": "on_disk" if use_on_disk else "в памяти",
            "Предупреждения": warnings_text,
        }
        dfs["stats"] = ExcelReportGenerator.build_stats(dfs, source_file=pbix_path, extra=extra_stats)

        excel_path = ""
        if create_excel:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_path = os.path.join(output_dir, build_report_filename(pbix_path, timestamp))
            _log("Генерация Excel-файла...")
            ExcelReportGenerator(
                excel_path,
                dfs,
                source_file=pbix_path,
                sheets=ExcelReportGenerator.SHEETS_ORDER_PBIX,
            ).generate()
            _log(f"Отчет сохранен: {excel_path}")
        else:
            _log("Генерация Excel отключена (по настройке).")

        dfs["excel_path"] = excel_path
        dfs["file_kind"] = "pbix"
        dfs["model_size"] = model_size
        _progress(100, "Готово!")
        return dfs
    finally:
        if model is not None:
            try:
                model.close()
            except Exception:  # noqa: BLE001 — закрытие не должно ломать результат
                pass
        if temp_dir:
            if keep_temp:
                _log(f"Временные файлы сохранены: {temp_dir}")
            else:
                shutil.rmtree(temp_dir, ignore_errors=True)


__all__ = ["PBIXAnalysisError", "analyze_pbix_file", "DEFAULT_ON_DISK_THRESHOLD_MB"]
