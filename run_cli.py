"""
Запуск PBIT Analyzer без GUI (headless).

Анализирует .pbit файл: распаковывает архив, извлекает модель данных и layout,
строит lineage и (по умолчанию) сохраняет Excel-отчет. Qt не импортируется,
поэтому скрипт работает на серверах без графической среды.

Использование:
    python run_cli.py report.pbit
    python run_cli.py report.pbit --output-dir out
    python run_cli.py report.pbit --no-excel --focus M22
    python run_cli.py report.pbit --focus "Мера | M22 | Total Sales" --keep-temp
    python run_cli.py report.pbit --quiet

Коды возврата:
    0 — успех, 1 — ошибка анализа, 2 — неверные аргументы/файл не найден.
"""
from __future__ import annotations

import argparse
import re
import sys
import traceback
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
WORK_DIR = PROJECT_ROOT / "_work"

# pbi_modules находится в _work: добавляем путь до импорта пакета.
if str(WORK_DIR) not in sys.path:
    sys.path.insert(0, str(WORK_DIR))

FOCUS_ID_PATTERN = re.compile(r"(?i)^(s\d+|a\d+|m\d+|vis_id\s*:.+|vis:.+)$")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="run_cli.py",
        description="Анализ .pbit без GUI: модель, layout, lineage и Excel-отчет.",
        epilog="Пример: python run_cli.py report.pbit --output-dir out --focus M22",
    )
    parser.add_argument("pbit", help="Путь к .pbit файлу")
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
        help="Сохранить распакованные файлы PBIT",
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
    return parser


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

    print()
    print("=" * 60)
    print("PBIT Analyzer — результат (без GUI)")
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

    pbit_path = Path(args.pbit).expanduser()
    if not pbit_path.is_file():
        print(f"Ошибка: файл не найден — {pbit_path}", file=sys.stderr)
        return 2
    if pbit_path.suffix.lower() != ".pbit":
        print(f"Предупреждение: расширение файла не .pbit — {pbit_path.name}", file=sys.stderr)

    output_dir = Path(args.output_dir).expanduser() if args.output_dir else pbit_path.parent
    focus = normalize_focus(args.focus)
    if args.focus and not focus:
        print(f"Ошибка: не удалось разобрать фокус lineage — {args.focus!r}", file=sys.stderr)
        return 2

    # Импорты после проверки аргументов: pbi_modules тянет pandas/openpyxl.
    from pbi_modules.analyzer import analyze_pbit_file
    from pbi_modules.logging_config import setup_logging

    setup_logging("ERROR" if args.quiet else None)

    def on_progress(value: int, message: str) -> None:
        if not args.quiet:
            print(f"[{value:>3}%] {message}", file=sys.stderr, flush=True)

    def on_log(message: str) -> None:
        if not args.quiet:
            print(f"        {message}", file=sys.stderr, flush=True)

    try:
        dfs = analyze_pbit_file(
            str(pbit_path),
            str(output_dir),
            keep_temp=args.keep_temp,
            lineage_focus=focus,
            create_excel=not args.no_excel,
            progress_callback=on_progress,
            log_callback=on_log,
        )
    except Exception as error:  # noqa: BLE001 — CLI обязан сообщить об ошибке, а не падать трейсбеком
        print(f"Ошибка анализа: {error}", file=sys.stderr)
        if args.debug:
            traceback.print_exc()
        return 1

    _print_summary(pbit_path, dfs, focus, args.quiet)
    return 0


if __name__ == "__main__":
    sys.exit(main())
