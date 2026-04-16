import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime

from .excel_report import ExcelReportGenerator
from .extractors import DataModelExtractor, LayoutExtractor, load_json_utf16
from .lineage import build_lineage_directions, build_lineage_edges


def find_model_and_layout_files(temp_dir: str) -> tuple[str | None, str | None]:
    data_model_path = None
    layout_path = None
    for root, _dirs, files in os.walk(temp_dir):
        for filename in files:
            if filename == "DataModelSchema":
                data_model_path = os.path.join(root, filename)
            elif filename == "Layout":
                layout_path = os.path.join(root, filename)
    return data_model_path, layout_path


def build_report_filename(pbit_path: str, timestamp: str) -> str:
    source_name = os.path.splitext(os.path.basename(pbit_path))[0]
    safe_name = re.sub(r'[<>:"/\\|?*]+', "_", source_name).strip(" .")
    safe_name = safe_name or "pbit_source"
    return f"pbi_report_{safe_name}_{timestamp}.xlsx"


def analyze_pbit_file(
    pbit_path: str,
    output_dir: str,
    keep_temp: bool = False,
    lineage_focus: str = "",
    create_excel: bool = True,
    progress_callback=None,
    log_callback=None,
) -> dict:
    if not os.path.exists(pbit_path):
        raise FileNotFoundError(f"Файл не найден: {pbit_path}")

    os.makedirs(output_dir, exist_ok=True)
    temp_dir = tempfile.mkdtemp(prefix="_temp_pbit_", dir=output_dir)

    def _progress(value: int, message: str) -> None:
        if progress_callback:
            progress_callback(value, message)

    def _log(message: str) -> None:
        if log_callback:
            log_callback(message)

    try:
        _progress(10, "Распаковка PBIT файла...")
        with zipfile.ZipFile(pbit_path, "r") as archive:
            archive.extractall(temp_dir)
        _log("PBIT файл распакован")

        data_model_path, layout_path = find_model_and_layout_files(temp_dir)
        if not data_model_path:
            raise FileNotFoundError("Файл DataModelSchema не найден в архиве")
        if not layout_path:
            raise FileNotFoundError("Файл Layout не найден в архиве")

        _progress(30, "Загрузка DataModelSchema...")
        model_data = load_json_utf16(data_model_path)

        _progress(50, "Извлечение модели данных...")
        extractor = DataModelExtractor(model_data)
        df_tables = extractor.extract_tables_ref()
        df_columns = extractor.extract_columns()
        df_sources = extractor.extract_sources()
        df_relationships = extractor.extract_relationships()
        df_measures = extractor.extract_measures()

        _progress(70, "Извлечение Layout...")
        layout_data = load_json_utf16(layout_path)
        layout_ext = LayoutExtractor(layout_data, extractor)
        df_sheets = layout_ext.extract_sheets()
        df_visuals = layout_ext.extract_visuals()
        df_conditional = layout_ext.extract_conditional_formatting()

        _progress(85, "Формирование отчета...")
        dfs = {
            "tables_ref": df_tables,
            "columns": df_columns,
            "sources": df_sources,
            "relationships": df_relationships,
            "measures": df_measures,
            "sheets": df_sheets,
            "visuals": df_visuals,
            "conditional": df_conditional,
        }
        dfs["lineage"] = build_lineage_edges(dfs)
        lineage_upstream, lineage_downstream = build_lineage_directions(
            dfs["lineage"],
            focus_id=lineage_focus,
        )
        dfs["lineage_upstream"] = lineage_upstream
        dfs["lineage_downstream"] = lineage_downstream

        dfs["stats"] = ExcelReportGenerator.build_stats(dfs, source_file=pbit_path)

        excel_path = ""
        if create_excel:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = build_report_filename(pbit_path, timestamp)
            excel_path = os.path.join(output_dir, excel_filename)
            _log("Генерация Excel-файла...")
            ExcelReportGenerator(excel_path, dfs, source_file=pbit_path).generate()
            _log(f"Отчет сохранен: {excel_path}")
        else:
            _log("Генерация Excel отключена (по настройке).")

        if keep_temp:
            _log(f"Временные файлы сохранены: {temp_dir}")
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)
            _log("Временные файлы удалены")

        dfs["excel_path"] = excel_path
        _progress(100, "Готово!")
        return dfs
    except Exception:
        if not keep_temp:
            shutil.rmtree(temp_dir, ignore_errors=True)
        raise
