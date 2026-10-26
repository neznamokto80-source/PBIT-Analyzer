"""
Фасад над анализаторами: единая точка входа для .pbit и .pbix.

Существующие analyze_pbit_file() и analyze_pbix_file() не меняются —
диспетчер только выбирает нужный по типу файла. Импорт PBIX-модулей
ленивый: запуск GUI/CLI и разбор .pbit не тянут pbixray.
"""
from __future__ import annotations

from .file_kind import PBIX, FileFormatError, detect_file_kind


def analyze_file(
    file_path: str,
    output_dir: str,
    keep_temp: bool = False,
    lineage_focus: str = "",
    create_excel: bool = True,
    progress_callback=None,
    log_callback=None,
    on_disk: bool | None = None,
) -> dict:
    """Анализирует .pbit или .pbix и возвращает общий словарь результатов.

    Raises:
        FileNotFoundError: файл не найден.
        FileFormatError: формат не поддерживается или не совпал с расширением.
        PBIXAnalysisError: PBIX не читается (live-connection, повреждение).
    """
    kind = detect_file_kind(file_path)
    if kind == PBIX:
        from .pbix_analyzer import analyze_pbix_file

        dfs = analyze_pbix_file(
            file_path,
            output_dir,
            keep_temp=keep_temp,
            lineage_focus=lineage_focus,
            create_excel=create_excel,
            progress_callback=progress_callback,
            log_callback=log_callback,
            on_disk=on_disk,
        )
    else:
        from .analyzer import analyze_pbit_file

        dfs = analyze_pbit_file(
            file_path,
            output_dir,
            keep_temp=keep_temp,
            lineage_focus=lineage_focus,
            create_excel=create_excel,
            progress_callback=progress_callback,
            log_callback=log_callback,
        )
    dfs.setdefault("file_kind", kind)
    return dfs


__all__ = ["analyze_file", "FileFormatError"]
