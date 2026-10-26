# PBIT Analyzer

Desktop-приложение для анализа `Power BI Template` (`.pbit`) файлов.
markdown
![Скриншот](screenshot/measures.png)
![Скриншот](screenshot/leneage.png)


## Что делает

- распаковывает `.pbit` и читает `DataModelSchema` + `Layout`;
- извлекает таблицы, столбцы, меры, связи, источники;
- анализирует листы (включая размеры width x height), визуалы и условное форматирование;
- показывает данные во вкладках GUI;
- опционально формирует Excel-отчет (по галке `Создать Excel`);
- строит lineage (общий и фокусный для выбранного элемента).

## GUI: ключевые возможности

- все чекбоксы по умолчанию выключены;
- чекбокс `Создать Excel` управляет созданием `.xlsx`;
- чекбокс `Открыть Excel файл` работает только если Excel был создан;
- поле `Фокус для lineage` активно после завершения анализа;
- `Фокус для lineage` — выпадающий список с автодополнением и ручным вводом;
- кнопка `Найти lineage` пересчитывает `Lineage Upstream/Downstream` только для выбранного фокуса.

## Формат ссылок и ID

- ссылки на таблицы: `📊 S...` (например, `📊 S4`);
- ссылки на столбцы: `📋 A...` (например, `📋 A601`);
- ссылки на меры: `📐 M...` (например, `📐 M22`);
- ID визуала: `VIS_ID :<id>` (например, `VIS_ID :f48726d8e921d4f3f01f`).

## Формат элементов в "Фокус для lineage"

- общий формат: `Тип элемента | ID | Название`;
- для визуалов: `Тип элемента | ID | Тип визуала`;
- названия подтягиваются из соответствующих вкладок:
  - `Таблица` -> `Справочник таблиц`;
  - `Столбец` -> `Столбцы`;
  - `Мера` -> `Меры`;
  - `Визуал` -> `Визуальные элементы` (тип визуала).

## Требования

- Python 3.10+
- Windows (основной сценарий использования)

## Установка

```bash
python -m pip install -r requirements.txt
```

## Запуск

Через bat-файл:

```bash
tasks.bat
```
```
откроется окно для выбора
==========================================
            PBIT Analyzer Tasks
==========================================
1. Run GUI application
   Starts: python _work\last_work_gui_work.py
2. Lint code
   Runs:   python -m ruff check .
3. Run tests
   Runs:   python -m pytest
4. Check all (lint + test)
   Runs lint first, then tests
5. Exit

CHOICE=Enter option number [1-5]: 
```


Или напрямую из корня проекта:

```bash
python run_gui.py
```

(старая точка входа `python _work\last_work_gui_work.py` тоже работает)

Примечание: авто-установка недостающих библиотек

## Запуск без GUI (headless)

`run_cli.py` выполняет тот же анализ, но без PyQt6 — подходит для серверов и скриптов.

```bash
python run_cli.py report.pbit                    # Excel рядом с .pbit
python run_cli.py report.pbit -o out             # папка для результата
python run_cli.py report.pbit --no-excel         # только анализ
python run_cli.py report.pbit --focus M22        # фокус lineage
python run_cli.py report.pbit --keep-temp -q     # временные файлы, без логов
```

Коды возврата: `0` — успех, `1` — ошибка анализа, `2` — неверные аргументы или файл не найден.

## Сборка EXE

```bash
python -m pip install -r requirements.txt pyinstaller

python build_exe.py            # GUI + CLI
python build_exe.py --gui      # только GUI
python build_exe.py --cli      # только CLI (без Qt, заметно меньше)
python build_exe.py --dir      # папка вместо одного файла (быстрее запуск)
python build_exe.py --clean    # очистить build/ dist/ и .spec
```

Скрипт сам ставит недостающие зависимости, подставляет версию из `plan/CHANGELOG.md`
и копирует готовые `PBIT_Analyzer_GUI_v*.exe` / `PBIT_Analyzer_CLI_v*.exe` в корень проекта.
В GUI-сборку попадают `styles.qss`, `schema_documentation.md` и `plan/CHANGELOG.md`,
поэтому справка и оформление работают и в EXE.

## Структура

- `run_gui.py` — точка входа GUI (корень проекта).
- `run_cli.py` — точка входа без GUI (headless).
- `build_exe.py` — сборка EXE через PyInstaller.
- `_work/last_work_gui_work.py` — исходная точка входа GUI с авто-установкой пакетов.
- `_work/pbi_modules/gui.py` — интерфейс и фоновый worker.
- `_work/pbi_modules/analyzer.py` — сервис анализа `.pbit` (оркестрация).
- `_work/pbi_modules/extractors.py` — извлечение из `DataModelSchema` и `Layout`.
- `_work/pbi_modules/excel_report.py` — генерация Excel.
- `_work/pbi_modules/help_content.py` — HTML-текст справки.
- `_work/pbi_modules/help_dialog.py` — диалоговое окно справки с вертикальным скроллом.
- `_work/pbi_modules/formatters.py` — форматирование данных для отображения.
- `_work/pbi_modules/lineage.py` — построение графа зависимостей.
- `_work/pbi_modules/logging_config.py` — конфигурация логирования.
