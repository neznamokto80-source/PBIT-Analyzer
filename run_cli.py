"""
Запуск PBIT Analyzer без GUI (headless).

Анализирует .pbit (модель + layout) или .pbix (только семантическая модель),
строит lineage и (по умолчанию) сохраняет Excel-отчет. Qt не импортируется,
поэтому скрипт работает на серверах без графической среды.

Использование:
    python run_cli.py report.pbit
    python run_cli.py report.pbix --output-dir out
    python run_cli.py report.pbix --no-excel --focus M22
    python run_cli.py report.pbit --focus "Мера | M22 | Total Sales" --keep-temp
    python run_cli.py report.pbix --on-disk
    python run_cli.py report.pbit --quiet

Коды возврата:
    0 — успех, 1 — ошибка анализа, 2 — неверные аргументы/файл не найден,
    3 — неподдерживаемый/повреждённый файл или PBIX без встроенной модели.
"""
from __future__ import annotations

import argparse
import re
import sys
import traceback
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
WORK_DIR = PROJECT_ROOT / "_work"


def use_utf8_output() -> None:
    """Переводит stdout/stderr в UTF-8 (на Windows cp1252 кириллица падает)."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass


use_utf8_output()

# pbi_modules находится в _work: добавляем путь до импорта пакета.
if str(WORK_DIR) not in sys.path:
    sys.path.insert(0, str(WORK_DIR))

FOCUS_ID_PATTERN = re.compile(r"(?i)^(s\d+|a\d+|m\d+|vis_id\s*:.+|vis:.+)$")
VISUAL_FOCUS_PATTERN = re.compile(r"(?i)^vis(:|_id)")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="run_cli.py",
        description="Анализ .pbit/.pbix без GUI: модель, lineage и Excel-отчет.",
        epilog="Пример: python run_cli.py report.pbix --output-dir out --focus M22",
    )
    parser.add_argument("file", metavar="FILE", nargs="?", help="Путь к файлу .pbit или .pbix")
    parser.add_argument(
        "-o",
        "--output-dir",
        help="Папка для результатов (по умолчанию — папка исходного .pbit)",
    )
    parser.add_argument(
        "--no-excel",
        action="store_true",
        help="Не создавать Excel-отчет (только анализ)",
    )
    parser.add_argument(
        "--focus",
        default="",
        help="Фокус lineage: S4 / A601 / M22 / VIS_ID :<id> или «Тип | ID | Имя»",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Сохранить временные файлы (для PBIX — только при чтении on_disk)",
    )
    disk_group = parser.add_mutually_exclusive_group()
    disk_group.add_argument(
        "--on-disk",
        action="store_true",
        help="PBIX: читать модель через memory-mapped файл (экономно по RAM)",
    )
    disk_group.add_argument(
        "--in-memory",
        action="store_true",
        help="PBIX: читать модель целиком в память, даже если файл большой",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Не печатать прогресс и логи",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Показать traceback при ошибке",
    )
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="Проверить зависимости (pbixray, pbi_modules) и выйти",
    )
    return parser


def self_test() -> int:
    """Smoke-тест сборки: проверяет импорт PBIX-модулей и бинарных зависимостей."""
    try:
        import pbixray  # noqa: F401

        from pbi_modules.dispatcher import analyze_file  # noqa: F401
        from pbi_modules.file_kind import detect_file_kind  # noqa: F401
        from pbi_modules.pbix_extractor import PBIXModelExtractor  # noqa: F401
    except ImportError as error:
        print(f"Ошибка зависимостей: {error}", file=sys.stderr)
        return 1
    print("Self-test OK")
    return 0


def normalize_focus(value: str) -> str:
    """Приводит фокус lineage к внутреннему ID (как в GUI).

    Поддерживает как чистые ID (``S4``, ``A601``, ``M22``, ``VIS_ID :<id>``),
    так и подписи из выпадающего списка GUI (``Тип | ID | Имя``).
    """
    raw = str(value or "").strip()
    if not raw:
        return ""

    if "|" in raw:
        for part in (item.strip() for item in raw.split("|")):
            if FOCUS_ID_PATTERN.match(part):
                return normalize_focus(part)
        return ""

    upper_raw = raw.upper()
    if upper_raw.startswith("VIS_ID"):
        match = re.match(r"VIS_ID\s*:\s*(.+)$", raw, re.IGNORECASE)
        return f"VIS:{match.group(1).strip()}" if match else ""
    if upper_raw.startswith("VIS:"):
        return f"VIS:{raw[4:].strip()}"
    if upper_raw.startswith("S"):
        return f"s{raw[1:]}"
    if upper_raw.startswith("A"):
        return f"A{raw[1:]}"
    if upper_raw.startswith("M"):
        return f"M{raw[1:]}"
    return raw


def _print_summary(pbit_path: Path, dfs: dict, focus: str, quiet: bool) -> None:
    if quiet:
        return

    kind = str(dfs.get("file_kind", "pbit")).upper()
    print()
    print("=" * 60)
    print(f"PBIT Analyzer — результат (без GUI, {kind})")
    print("=" * 60)
    print(f"Файл: {pbit_path}")

    stats = dfs.get("stats")
    if stats is not None and not stats.empty:
        for _, row in stats.iterrows():
            print(f"{str(row['Параметр']):<38} {row['Значение']}")

    if focus:
        print(f"{'Фокус lineage':<38} {focus}")

    print("-" * 60)
    excel_path = dfs.get("excel_path") or ""
    if excel_path:
        print(f"Excel-отчет: {excel_path}")
    else:
        print("Excel-отчет: не создавался (--no-excel)")
    print("=" * 60)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.self_test:
        return self_test()
    if not args.file:
        print("Ошибка: не указан файл .pbit/.pbix", file=sys.stderr)
        return 2

    file_path = Path(args.file).expanduser()
    if not file_path.is_file():
        print(f"Ошибка: файл не найден — {file_path}", file=sys.stderr)
        return 2
    if file_path.suffix.lower() not in {".pbit", ".pbix"}:
        print(
            f"Ошибка: неподдерживаемое расширение {file_path.suffix or '(нет)'} — "
            "поддерживаются .pbit и .pbix",
            file=sys.stderr,
        )
        return 2

    output_dir = Path(args.output_dir).expanduser() if args.output_dir else file_path.parent
    focus = normalize_focus(args.focus)
    if args.focus and not focus:
        print(f"Ошибка: не удалось разобрать фокус lineage — {args.focus!r}", file=sys.stderr)
        return 2

    # Импорты после проверки аргументов: pbi_modules тянет pandas/openpyxl.
    from pbi_modules.dispatcher import analyze_file
    from pbi_modules.file_kind import PBIX, FileFormatError, detect_file_kind
    from pbi_modules.logging_config import setup_logging
    from pbi_modules.pbix_analyzer import PBIXAnalysisError

    setup_logging("ERROR" if args.quiet else None)

    try:
        file_kind = detect_file_kind(str(file_path))
    except FileFormatError as error:
        print(f"Ошибка формата: {error}", file=sys.stderr)
        return 3

    if file_kind == PBIX and focus and VISUAL_FOCUS_PATTERN.match(focus):
        print(
            "Ошибка: у PBIX нет слоя отчёта, фокус на визуале невозможен. "
            "Используйте таблицу (S…) или меру (M…).",
            file=sys.stderr,
        )
        return 2
    if file_kind == PBIX and args.keep_temp:
        print("Примечание: --keep-temp учитывается только при чтении PBIX через on_disk.", file=sys.stderr)

    on_disk = True if args.on_disk else (False if args.in_memory else None)

    def on_progress(value: int, message: str) -> None:
        if not args.quiet:
            print(f"[{value:>3}%] {message}", file=sys.stderr, flush=True)

    def on_log(message: str) -> None:
        if not args.quiet:
            print(f"        {message}", file=sys.stderr, flush=True)

    try:
        dfs = analyze_file(
            str(file_path),
            str(output_dir),
            keep_temp=args.keep_temp,
            lineage_focus=focus,
            create_excel=not args.no_excel,
            progress_callback=on_progress,
            log_callback=on_log,
            on_disk=on_disk,
        )
    except (FileFormatError, PBIXAnalysisError) as error:
        print(f"Ошибка анализа: {error}", file=sys.stderr)
        if args.debug:
            traceback.print_exc()
        return 3
    except Exception as error:  # noqa: BLE001 — CLI обязан сообщить об ошибке, а не падать трейсбеком
        print(f"Ошибка анализа: {error}", file=sys.stderr)
        if args.debug:
            traceback.print_exc()
        return 1

    _print_summary(file_path, dfs, focus, args.quiet)
    return 0


if __name__ == "__main__":
    sys.exit(main())
