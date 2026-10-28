# План рефакторинга PBIT Analyzer

**Дата:** 2026-10-28
**Версия кода:** 2.2.0 (по `docs/CHANGELOG.md`)
**Статус документа:** план согласован с решением «код не меняем — только зафиксировать, что необходимо сделать»
**Область:** `_work/pbi_modules/*`, `run_cli.py`, `run_gui.py`, `build_exe.py`, `create_exe.bat`, `tests/*`, структура репозитория.

> **Важно.** По решению владельца проекта на этом шаге **никакие исходники не изменялись**.
> Всё ниже — карта проблем, целевая структура и пошаговый план работ, готовый к исполнению
> после отдельного согласования по каждому этапу (правило `confirm-changes` из `AGENTS.md`).

---

## 0. Зачем этот документ

Проект вырос из одного скрипта в приложение с двумя форматами (`.pbit`/`.pbix`), GUI, CLI,
Excel-отчётом, lineage-графом и сборкой EXE. Рост шёл итерациями «добавить вкладку/лист»,
поэтому логика расползлась: два гигантских модуля (`extractors.py` 1347 строк, `gui.py` 1172 строки),
три источника правды о колонках и наборах вкладок, по три копии нормализации focus-ID,
мёртвый код и отладочные артефакты в рабочем пути.

Цель плана — сделать код предсказуемым для правок, **не меняя поведение приложения**:
те же вкладки, те же листы Excel, те же заголовки колонок, те же коды возврата CLI.

---

## 1. Исходное состояние: измерения

Замеры сделаны на текущем рабочем дереве (`py -3` + `ast`-разбор, `ruff`, `pytest`).

| Показатель | Значение |
|---|---|
| Всего Python-файлов (без кэшей) | 36 |
| Всего строк Python | 7254 |
| Модулей приложения (`_work/pbi_modules`) | 14 |
| Тестовых файлов в `tests/` | 10 |
| Собрано тестов | 44 |
| Результат `pytest -q` | **44 passed, 23.5 s** |
| `ruff check .` | 4 замечания — все в черновиках `_work/test_*.py` |
| Публичных функций/методов в `pbi_modules` | 64 |
| Из них ни разу не упомянуты в `tests/` | 53 |
| Покрытие (line coverage) | не измеряется: `pytest-cov` не установлен |

### 1.1. Размеры файлов

| Строк | Функций | Классов | Файл |
|---:|---:|---:|---|
| 1347 | 35 | 3 | `_work/pbi_modules/extractors.py` |
| 1172 | 52 | 3 | `_work/pbi_modules/gui.py` |
| 560 | 26 | 1 | `_work/pbi_modules/pbix_extractor.py` |
| 424 | 13 | 0 | `build_exe.py` |
| 421 | 28 | 3 | `_work/pbi_modules/formatters.py` |
| 414 | 14 | 0 | `_work/pbi_modules/lineage.py` |
| 301 | 6 | 0 | `_work/pbi_modules/help_content.py` |
| 273 | 8 | 0 | `run_cli.py` |
| 256 | 7 | 1 | `_work/pbi_modules/excel_report.py` |
| 167 | 5 | 1 | `_work/pbi_modules/pbix_analyzer.py` |
| 151 | 5 | 0 | `_work/pbi_modules/analyzer.py` |
| 102 | 3 | 0 | `run_gui.py` |
| 96 | 4 | 0 | `_work/pbi_modules/file_kind.py` |
| 76 | 3 | 1 | `_work/pbi_modules/help_dialog.py` |
| 60 | 2 | 0 | `_work/last_work_gui_work.py` |
| 13 | 1 | 0 | `_work/pbi_modules/logging_config.py` |

### 1.2. Самые длинные функции

| Строк | Ветвлений | Место | Функция |
|---:|---:|---|---|
| 182 | 28 | `extractors.py:1073` | `LayoutExtractor.extract_conditional_formatting` |
| 171 | 49 | `lineage.py:118` | `build_lineage_edges` |
| 148 | 21 | `extractors.py:924` | `LayoutExtractor.extract_visuals` |
| 127 | 40 | `gui.py:796` | `PBITAnalyzerMainWindow._apply_lineage_focus` |
| 104 | 17 | `pbix_analyzer.py:61` | `analyze_pbix_file` |
| 100 | 10 | `analyzer.py:52` | `analyze_pbit_file` |
| 100 | 20 | `extractors.py:397` | `LayoutExtractor._try_resolve_field` |
| 95 | 32 | `gui.py:700` | `_populate_lineage_focus_options` |
| 91 | 24 | `help_content.py:6` | `_md_to_html` |
| 87 | 19 | `extractors.py:498` | `_extract_from_projections` |
| 85 | 21 | `run_cli.py:185` | `main` |
| 70 | 24 | `extractors.py:789` | `_extract_filter_details` |
| 66 | 0 | `excel_report.py:136` | `ExcelReportGenerator._get_columns` |
| 65 | 26 | `extractors.py:41` | `DependencyAnalyzer.find_dependencies` |
| 59 | 14 | `gui.py:1026` | `_create_table_from_df` |
| 58 | 11 | `build_exe.py:278` | `build` |

### 1.3. «Классы-боги»

| Методов | Место | Класс |
|---:|---|---|
| 40 | `gui.py:220` | `PBITAnalyzerMainWindow` |
| 23 | `extractors.py:330` | `LayoutExtractor` |
| 21 | `pbix_extractor.py:116` | `PBIXModelExtractor` |
| 16 | `formatters.py:4` | `DAXFormatter` |
| 9 | `extractors.py:112` | `DataModelExtractor` |
| 7 | `excel_report.py:22` | `ExcelReportGenerator` |
| 6 | `gui.py:138` | `CustomTableWidget` |

---

## 2. Карта модулей «как есть»

Слои и фактические зависимости (стрелка = «импортирует»):

```
run_gui.py ──► pbi_modules.gui ──► dispatcher ──► analyzer ──► extractors ──► formatters
    │              │                    │              │            │
    │              │                    │              │            └──► pandas
    │              │                    │              └──► lineage ──► pandas
    │              │                    └──► pbix_analyzer ──► pbix_extractor ──► pbixray
    │              ├──► excel_report ──► openpyxl, pandas
    │              ├──► file_kind ──► zipfile
    │              ├──► lineage
    │              └──► help_dialog ──► help_content ──► (файлы .md из бандла)
run_cli.py ──► dispatcher, excel_report (те же пути, без PyQt6)
build_exe.py ──► PyInstaller (внешний процесс)
```

| Модуль | Ответственность сейчас | Проблема слоя |
|---|---|---|
| `file_kind.py` | определение формата по расширению и содержимому zip | чистый, оставить как есть |
| `analyzer.py` | распаковка `.pbit`, извлечение, lineage, Excel, temp-dir | смешивает I/O, оркестрацию и отчёт |
| `pbix_analyzer.py` | то же для `.pbix` через pbixray | дублирует скелет `analyzer.py` |
| `dispatcher.py` | выбор анализатора по типу | чистый фасад, оставить |
| `extractors.py` | DAX-зависимости + модель + Layout (3 класса) | три ответственности в одном файле |
| `formatters.py` | DAX/SQL/M форматирование | внутреннее дублирование |
| `lineage.py` | рёбра и направления графа | гигантская функция сборки рёбер |
| `excel_report.py` | листы, стили, merge, статистика | третий список схем колонок |
| `gui.py` | окно, воркер, таблицы, фильтры, lineage-UI | 40 методов в одном классе |
| `help_content.py` | Markdown→HTML, справка | самописный конвертер |
| `logging_config.py` | базовая настройка логов | не используется? проверить |

---

## 3. Каталог проблем

Формат: **ID — приоритет (P0 критично / P1 важно / P2 желательно) — место — суть — что делать — как проверить.**

### A. Корректность и риск потери данных

**P-01 — P0 — `extractors.py:290-328` (`DataModelExtractor.extract_measures`)**
Меры собираются циклом `for t in self.tables`, а `self.tables` — только **видимые** таблицы
(`extractors.py:117-120` фильтрует `isHidden`). Мера, закреплённая за скрытой таблицей, исчезает
из вкладки «Меры», из Excel и из lineage. При этом `CHANGELOG.md` (2.2.0, «Исправлено») и
`PBIX_PLAN.md` (п. 7b) обещают обратное — но исправление сделано **только в PBIX-пути**
(`pbix_extractor.py:224`, покрыто `tests/test_pbix_extractor.py:116`). Для `.pbit` поведение
осталось прежним, теста нет.
*Что делать:* добавить `self.all_tables` (все таблицы модели) для сбора мер и построения
`measure_to_number`/`measure_to_table`, при этом `table_number` для скрытой таблицы = «—»;
`extract_measures` итерировать по `all_tables`, `extract_tables_ref` — по видимым.
*Как проверить:* новый тест `tests/test_extractors.py::test_hidden_table_measure_kept` на модели
с `isHidden: true` + мере; проверить, что мера есть в `measures` и в `lineage`.

**P-02 — P0 — `extractors.py:118`**
`t["name"]` вместо `t.get("name", "")` — `KeyError` на таблице без имени (битые/частичные
`DataModelSchema`). Рядом фильтр `t["name"] != "DateTableTemplate"` сравнивает **только точное
имя**, хотя план (`PBIX_PLAN.md`, п. 6) требует отсекать `DateTableTemplate_*` и
`LocalDateTable_*`. В итоге системные таблицы дат могут попадать в «Справочник таблиц».
*Что делать:* `name = t.get("name") or ""`, `name.startswith(("DateTableTemplate", "LocalDateTable"))`.
*Как проверить:* тест на модель с `LocalDateTable_abc` → таблицы нет в `tables_ref`; тест на
таблицу без `name` → нет исключения.

**P-03 — P1 — `extractors.py:944-945`**
`visual_id_label = self._format_visual_id(visual_id)` продублирована дважды подряд — безвредно,
но показывает, что файл правился вручную без ревью.
*Что делать:* удалить дубль (при рефакторинге `extract_visuals`).

**P-04 — P1 — `extractors.py:982, 1028, 1060-1063`**
Отладочные ветки с захардкоженными ID визуалов (`d7de85c5ac4282aac955`, `6e69509a76aa9b4efc3e`)
и `logger.debug("DEBUG ...")` в рабочем пути; `from collections import Counter as _Counter`
импортируется внутри функции.
*Что делать:* удалить ветки и импорт; для диагностики достаточно уровня `DEBUG` в вызывающем коде.
*Риск:* нулевой — код исполняется только на двух конкретных ID из чужих отчётов.

**P-05 — P1 — `extractors.py:1090-1094`, `excel_report.py:118-127`**
`except Exception: continue` при разборе `config` визуала и `except Exception` вокруг merge ячеек —
тихо теряют визуал/форматирование без следа в логах.
*Что делать:* ловить `json.JSONDecodeError`/`TypeError`, при прочих — `logger.warning` с ID визуала.

**P-06 — P1 — `gui.py:1042-1070` (`_create_table_from_df`)**
Хрупкая конвертация значений: `pd.isna` с `hasattr(result, '__len__')`, затем `hasattr(value, '__len__')`
и `str(value)` для массивоподобных. На `NaT`, `Decimal`, `timedelta`, вложенных `dict` поведение
зависит от типа pandas-объекта.
*Что делать:* вынести в чистую функцию `format_cell(value) -> str` с явными ветками
(`pd.isna` через `pd.api.types.is_scalar`, `isinstance(value, str)`), покрыть параметризованным тестом.
*Проверка:* тест на набор значений `None, pd.NA, pd.NaT, float('nan'), "a\nb", [1,2], {"a":1}`.

### B. Дублирование (главный источник будущих багов)

**P-07 — P1 — три источника правды о колонках**
`excel_report.py:136-201` (`_get_columns`, 66 строк), `pbix_extractor.py:60-78` (`*_SCHEMA`),
`extractors.py:181-328` (колонки задаются литералами в каждом `dict`).
Любое переименование колонки надо делать в 2-3 местах; рассинхрон даёт пустые листы.
*Что делать:* единый реестр `schemas.py`: `SCHEMAS: dict[str, tuple[str, ...]]` + порядок листов;
`extractors.py` собирает строки через `dict(zip(SCHEMAS[key], values))` либо через dataclass-строки;
`excel_report._get_columns` и `pbix_extractor` берут схему из реестра.
*Проверка:* тест «каждый ключ из `SHEETS_ORDER*` имеет схему», «колонки DataFrame ⊆ схемы».

**P-08 — P1 — два списка вкладок**
`gui.py:52-84` (`TAB_MAPPING`, `TAB_MAPPING_PBIX`) и `excel_report.py:24-57`
(`SHEETS_ORDER`, `SHEETS_ORDER_PBIX`) описывают одно и то же двумя списками с разными именами.
*Что делать:* один `schemas.py`: `TABS_BY_KIND`, `SHEETS_BY_KIND` (лист ≠ вкладка только для
«Статистика хранения»), GUI и Excel импортируют оттуда.

**P-09 — P1 — три копии нормализации focus-ID**
`gui.py:648` (`_normalize_focus_input`), `lineage.py:85` (`_normalize_focus_id`),
`run_cli.py:126` (`normalize_focus`) — три разные реализации одного правила (`S4`, `M3`, `A105`,
`VIS_ID :x`, `VIS:x`, префиксы `T:`/`C:`/`M:`).
*Что делать:* один `focus.py` с `normalize_focus_id()` и `parse_focus_label()`, все три места
делегируют в него. `gui._display_focus_id` и вложенная `normalize_visual_id` (`gui.py:885`) — туда же.

**P-10 — P1 — две реализации `get_column_number`**
`extractors.py:164` и `pbix_extractor.py:261` — одинаковая логика `A{t}{cc}`.
*Что делать:* общий `ids.py::column_number(table_index, column_index)`.

**P-11 — P1 — дубли в `formatters.py`**
`_add_spaces_operators` (стр. ~100-150) и `_apply_operator_spacing` (стр. ~174-215) — почти
посимвольно одинаковые (плейсхолдеры `__SQ/__DQ/__CMT`, правила пробелов).
`_format_with_name` (стр. 35-75) и `format_dax_formula` (стр. 284-333) — две копии разбора
`VAR/RETURN`. `format_column`/`format_measure`/`format_table` — три обёртки над одной функцией.
*Что делать:* оставить один `_spacing(expr)` и один `_render(expr, prefix)`; публичные методы —
делегаты.

**P-12 — P2 — скелет анализаторов**
`analyzer.analyze_pbit_file` (100 строк) и `pbix_analyzer.analyze_pbix_file` (104 строки) —
одинаковый каркас: `_progress/_log`, temp-dir, Excel по флагу, `dfs["excel_path"]`, очистка в
`finally`.
*Что делать:* `pipeline.py::run_analysis(source, extract, ...)`, куда передаётся функция извлечения.
Даёт единое место для прогресса/очистки/ошибок.

**P-13 — P2 — `_work/last_work_gui_work.py:16`**
`check_and_install_packages()` дублирует `build_exe.ensure_runtime_packages()` и **ставит пакеты
через pip при запуске** — сеть в рантайме, непредсказуемый старт.
*Что делать:* удалить legacy-entrypoint или свести к `from pbi_modules.gui import main`.

### C. Архитектура и читаемость

**P-14 — P1 — `extractors.py` = 3 ответственности**
`DependencyAnalyzer` (разбор DAX), `DataModelExtractor` (модель), `LayoutExtractor` (отчёт).
23 метода Layout-класса, 100-строчный `_try_resolve_field`, обращения к приватным полям
`self.ext._normalize`, `self.ext.norm_*`, `self.ext.measure_to_number`.
*Что делать:* пакет `pbi_modules/extract/`:
```
extract/
  __init__.py      # реэкспорт: DataModelExtractor, LayoutExtractor, load_json_utf16
  json_io.py       # load_json_utf16
  model.py         # DataModelExtractor + ModelIndex
  index.py         # ModelIndex: table_number, column_number, measure_number, normalize
  dependencies.py  # DependencyAnalyzer (зависит только от ModelIndex)
  layout/
    __init__.py    # LayoutExtractor (фасад)
    page.py        # extract_sheets
    fields.py      # _try_resolve_field, _resolve_field_with_table, projections
    filters.py     # _extract_filter_details, _parse_json_if_string, _extract_filter_target
    formatting.py  # _extract_formatting, _extract_position
    conditional.py # extract_conditional_formatting + _simplify_condition/_extract_expression_str
    visuals.py     # extract_visuals (сборка строк)
```
Публичный API (`DataModelExtractor`, `LayoutExtractor`, `load_json_utf16`) сохраняется реэкспортом,
поэтому `analyzer.py`, `test_extractors.py` и hidden-imports не ломаются.

**P-15 — P1 — `gui.py` = окно + воркер + таблица + фильтры + lineage**
40 методов в `PBITAnalyzerMainWindow`, `_apply_lineage_focus` 127 строк/40 ветвлений.
*Что делать:* пакет `pbi_modules/ui/`:
```
ui/
  window.py       # PBITAnalyzerMainWindow: только сборка UI и связывание
  actions.py      # запуск анализа, экспорт, открытие файлов/папок
  tables.py       # CustomTableWidget, TableFactory (df -> QTableWidget), format_cell
  search.py       # SearchBar + фильтрация строк
  lineage_panel.py# модель выбора focus, применение/сброс фильтра, сбор focus-опций
  worker.py       # AnalysisWorker
  styles.py       # _load_stylesheet
```
`_apply_lineage_focus` разбивается на: `parse_focus()`, `collect_related_visuals()`,
`filter_dfs(dfs, focus)` (чистая функция, тестируемая без Qt), `render()`.

**P-16 — P1 — `lineage.py:118` (`build_lineage_edges`, 171 строка/49 ветвлений)**
Четыре почти одинаковых блока «пройти DataFrame → добавить рёбра».
*Что делать:* `_names_from(df, id_col, name_col)`, `_add_table_edges`, `_add_column_edges`,
`_add_measure_edges`, `_add_visual_edges` + общий `EdgeBuilder`.

**P-17 — P2 — `help_content._md_to_html` (91 строка)**
Самописный Markdown-конвертер регулярками; риск неверного HTML и XSS при подстановке текста.
*Что делать:* либо оставить (текст локальный), либо заменить на минимальный набор правил
(`markdown-it` как зависимость — отдельное решение), обязательно экранировать `<`/`>`/`&`.

**P-18 — P2 — магические значения**
`LINEAGE_HEAVY_TABLE_ROWS_THRESHOLD = 1200` (`gui.py:89`), обрезка 200 символов (`gui.py:1048`),
`min(max_len + 2, 80)` (`excel_report.py:227`), 500 МБ (`pbix_analyzer.py:25`), `60/70/85/100`
прогресса, строковые ключи `"lineage_upstream"`, `"№ меры"`, `"—"`.
*Что делать:* `constants.py` (пороги, проценты), `keys.py`/`Enums` для ключей наборов и колонок;
`"—"` → `EMPTY = "—"`.

**P-19 — P2 — отсутствие типизации и контрактов**
Публичные методы `extractors.py` без аннотаций; `dfs` — «мешок» DataFrame и строк
(`excel_path`, `file_kind`, `warnings`, `model_size`).
*Что делать:* `AnalysisResult` (TypedDict или dataclass) + `from __future__ import annotations`
во всех модулях; `mypy --strict` на новых пакетах (внедрять поэтапно).

### D. Производительность

**P-20 — P1 — `get_column_number` линейный и вызывается в цикле**
`extractors.py:164`: поиск таблицы по списку + поиск колонки по списку. Вызов из
`find_dependencies` (`extractors.py:41`) для **каждой** ссылки **каждой** меры/столбца → O(T·C) на
ссылку. На моделях с сотнями таблиц и тысячами колонок lineage строится заметно дольше нужного.
*Что делать:* в `_build_lookups` построить `self.column_number[(table, column)] = "A0103"`;
`get_column_number` станет O(1). То же для `LayoutExtractor._resolve_field_with_table`
(`extractors.py:621`), где снова линейный проход по `self.ext.tables`.
*Метрика до/после:* время `extract_measures` + `build_lineage_edges` на реальном `.pbit`
(нужен замер на файле владельца — сейчас его нет в репозитории).

**P-21 — P2 — `iterrows()`**
`lineage.py` (4 места), `gui.py:_populate_lineage_focus_options` (4 места), `excel_report` — построчный
проход по DataFrame вместо векторизации/`itertuples`. Для 10k+ строк lineage это единицы секунд.
*Что делать:* `itertuples(index=False)` там, где нужны значения; словари через `zip(df[col], df[col2])`.

**P-22 — P2 — `_create_table_from_df` строит ячейки по одной**
`gui.py:1026`: `len(df) × len(columns)` вызовов `QTableWidgetItem` в UI-потоке. Для 5-10k строк —
подвисание.
*Что делать:* `setUpdatesEnabled(False)` вокруг заполнения (минимальная правка), затем
QAbstractTableModel (правильное решение) — оценить отдельно.

### E. Тесты и качество

**P-23 — P1 — 53 из 64 публичных функций не упоминаются в тестах**
Полностью без прямых тестов: `extract_columns`, `extract_sources`, `extract_relationships`,
`extract_sheets`, `extract_conditional_formatting`, весь `DAXFormatter`/`SQLFormatter`/`SourceFormatter`,
`build_lineage_directions`, `build_stats`, `find_model_and_layout_files`, все `extract_*` PBIX
(кроме `extract_all`), GUI-классы.
*Что делать:* целевые тесты по слоям (см. этапы 4-6): форматтеры (вход/выход на 10-15 примерах DAX),
lineage (граф из 5 узлов), Excel (набор листов и заголовки), focus-фильтрация (чистая функция).

**P-24 — P1 — нет измерения покрытия**
`pytest-cov` не установлен, порог не задан.
*Что делать:* `pytest-cov` в dev-зависимости, `--cov=pbi_modules --cov-fail-under=60` в CI
(после этапов 4-6 — поднять до 75).

**P-25 — P2 — черновики в `_work/`**
7 файлов `_work/test_*.py` (~500 строк): `test_filter.py`, `test_filter2.py`, `test_fix.py`,
`test_fix2.py`, `test_fix_gui.py`, `test_fix_gui2.py`, `test_visual_filter2.py`. Не собираются
`pytest` (`testpaths = ["tests"]`), но ломают `ruff` (4 из 4 замечаний) и вводят в заблуждение
(«есть тесты на фильтр»).
*Что делать:* перенести полезные сценарии в `tests/test_gui_lineage.py` (с pytest-qt или без Qt —
через чистую функцию фильтрации), остальное удалить.

**P-26 — P2 — конфигурация линтера слабая**
`pyproject.toml` включает только правила по умолчанию (E/F). Нет `I` (порядок импортов),
`B` (bugbear), `UP` (pyupgrade), `SIM`, `RUF`.
*Что делать:* добавить `select = ["E","F","I","B","UP","SIM","RUF"]`, `ignore` для кириллицы в
строках не нужен; прогнать `--fix` отдельным коммитом.

### F. Сборка и упаковка

**P-27 — P2 — `build_exe.py` (424 строки, 13 функций)**
Смешивает: определение версии, список hidden-import'ов, установку пакетов, вызов PyInstaller,
копирование, очистку. `APP_VERSION` читается регуляркой из `docs/CHANGELOG.md`
(`build_exe.py:64-78`) — падение при смене формата заголовка даёт `_v0`.
*Что делать:* `pbi_modules/version.py` с `__version__` (единственный источник), CHANGELOG читается
только для проверки совпадения; разнести на `packaging/hidden_imports.py`, `packaging/clean.py`.

**P-28 — P2 — `pyproject.toml` без метаданных проекта**
Нет `[project]`, версии, зависимостей; рантайм-зависимости живут в `requirements.txt`, где вместе
с ними `pytest` и `ruff` (попадают в окружение сборки EXE), а `pyinstaller` не указан вовсе.
*Что делать:* `[project] dependencies` (PyQt6, pandas, openpyxl, pbixray) + `[project.optional-dependencies] dev` (pytest, ruff, pytest-cov, pyinstaller), `requirements.txt` оставить
тонкой ссылкой (`-e .`) или удалить.

### G. Репозиторий и процесс

**P-29 — P0 — рабочее дерево не там, где git-репозиторий проекта**
`git rev-parse --show-toplevel` в `D:/git/PBIT-Analyzer` отвечает `D:/git` (родительская папка с
чужими проектами). В этом репозитории tracked всего **7** файлов проекта:
`.github/workflows/build-exe.yml`, `README.md`, `build_exe.py`, `create_exe.bat`, `run_cli.py`,
`run_gui.py`, `_work/pbi_modules/extractors.py`. Остальные 20+ файлов (`gui.py`, `analyzer.py`,
`tests/*`, `plan/*`, `pyproject.toml`) — **untracked**. Коммиты при этом делались в отдельном
временном клоне GitHub-репозитория.
*Следствие:* правки в рабочей папке не попадают в историю; легко «потерять» файл или закоммитить
чужой проект (`../BO`, `../IAM` уже видны в `git status`).
*Что делать (требует отдельного согласования):* сделать `D:/git/PBIT-Analyzer` самостоятельным
репозиторием (`git init`, `.gitignore` уже готов, первый коммит текущего состояния) **или**
клонировать `neznamokto80-source/PBIT-Analyzer` и перенести рабочее дерево туда. До решения —
не запускать `git add -A` из `D:/git`.

**P-30 — P2 — документация расходится с кодом**
`CHANGELOG.md` утверждает исправление, которого нет в PBIT-пути (см. P-01); `PBIX_PLAN.md` п. 6
описывает фильтр системных таблиц, которого нет (P-02). `improvements.md` (11.5 КБ) и
`ANALYSIS_AND_PLAN.md` (21 КБ) частично устарели и не помечены как архив.
*Что делать:* починить код или документ; старые планы перенести в `plan/archive/` с шапкой «архив».

---

## 4. Целевая структура

```
_work/pbi_modules/
  __init__.py            # __version__, публичный API
  constants.py           # пороги, проценты прогресса, EMPTY
  schemas.py             # SCHEMAS, TABS_BY_KIND, SHEETS_BY_KIND  ← единый источник правды
  focus.py               # normalize_focus_id, parse_focus_label, display_focus_id
  ids.py                 # table_number/column_number/measure_number
  json_io.py             # load_json_utf16
  extract/
    __init__.py          # реэкспорт DataModelExtractor/LayoutExtractor
    model.py, index.py, dependencies.py
    layout/{page,fields,filters,formatting,conditional,visuals}.py
  lineage.py             # EdgeBuilder + build_lineage_edges + build_lineage_directions
  formatters.py          # DAXFormatter(_spacing/_render) + SQLFormatter + SourceFormatter
  excel_report.py        # листы/стили/merge, схемы из schemas.py
  pipeline.py            # общий каркас анализа (прогресс, temp, Excel)
  analyzer.py            # .pbit: распаковка → extract → pipeline
  pbix_analyzer.py       # .pbix: pbixray → extract → pipeline
  file_kind.py           # без изменений
  dispatcher.py          # без изменений
  ui/
    window.py, actions.py, tables.py, search.py, lineage_panel.py, worker.py, styles.py
  help_content.py, help_dialog.py, logging_config.py
```

Правила переезда (обязательны, чтобы не сломать поведение):

1. Публичные имена сохраняются; старое место импорта остаётся рабочим через реэкспорт
   (`from .extract import DataModelExtractor`).
2. `build_exe.py` hidden-imports обновляются **в том же коммите**, что и переезд пакета
   (`pbi_modules.extract`, `pbi_modules.ui.*`), иначе EXE соберётся без модулей.
3. Ни один этап не меняет: имена ключей `dfs`, порядок и названия вкладок/листов, заголовки
   колонок, коды возврата CLI (`0/2/3`), флаги `--self-test`, имена EXE, формат имени отчёта.

---

## 5. Этапы работ

Порядок выбран так, чтобы каждый этап был независимо проверяемым и обратимым. После каждого —
`ruff` + `pytest` + (для этапов 5-8) сборка и `--self-test`.

### Этап 0. Подготовка (0.5 шага)
1. Зафиксировать текущее состояние: `pytest -q` → 44 passed; `ruff check .` → 4 замечания в черновиках.
2. Добавить в `tests/` «золотые» снимки поведения: заголовки всех листов Excel для `.pbit` и `.pbix`,
   набор вкладок для каждого `file_kind`, коды возврата CLI на 4 сценариях (валидный pbit, `.txt`,
   переименованный pbix, битый файл), формат `build_report_filename`.
3. Записать эталонный `pbi_report_*.xlsx` (лист → колонки → первые 5 строк) в
   `tests/fixtures/expected_sheets.json`.
**DoD:** тесты падают при любом изменении контракта; рефакторинг этапов 1-8 их не ломает.

### Этап 1. Безопасные мелкие правки (P-03, P-04, P-05, P-25)
1. Удалить дубль `visual_id_label` (`extractors.py:944-945`).
2. Удалить захардкоженные debug-ветки и `Counter` из `extract_visuals`.
3. Заменить `except Exception` на конкретные типы + `logger.warning`.
4. Перенести полезное из `_work/test_*.py` в `tests/`, остальное удалить.
**DoD:** `ruff check .` → 0 замечаний; `pytest` → 44+ passed; diff в `extract_visuals` не меняет
выходной DataFrame (проверяется тестом из этапа 0).

### Этап 2. Единые схемы и вкладки (P-07, P-08, P-18)
1. Создать `schemas.py` со `SCHEMAS` (колонки по каждому ключу), `TABS_BY_KIND`, `SHEETS_BY_KIND`,
   `EMPTY`, `LINEAGE_HEAVY_TABLE_ROWS_THRESHOLD`, `PROGRESS_STEPS`.
2. `excel_report._get_columns` → `SCHEMAS.get(key, ())`; `SHEETS_ORDER*` → `SHEETS_BY_KIND`.
3. `gui.TAB_MAPPING*`/`tab_mapping_for` → `TABS_BY_KIND`.
4. `pbix_extractor.*_SCHEMA` → `SCHEMAS[...]` (или оставить как источник, на который ссылается реестр —
   главное, чтобы список был **один**).
5. `extractors.py`: собрать строки через общий хелпер `row(SCHEMAS[key], values)`, чтобы литеральные
   `dict` не расходились с реестром.
**DoD:** тест «у каждого листа/вкладки есть схема»; «колонки собранных DataFrame совпадают со схемой
по порядку»; Excel и GUI не изменились (эталон этапа 0).

### Этап 3. Focus-ID в одном месте (P-09, P-10)
1. `focus.py` с `normalize_focus_id`, `parse_focus_label`, `display_focus_id`.
2. `gui._normalize_focus_input`, `lineage._normalize_focus_id`, `run_cli.normalize_focus` → делегаты.
3. `ids.py::column_number`; `extractors.get_column_number` и `pbix_extractor.get_column_number` → делегаты.
**DoD:** параметризованный тест на 15 входов (`S4`, `s04`, `M3`, `a105`, `VIS_ID :abc`, `VIS:abc`,
`Тип | S4 | Продажи`, мусор) даёт одинаковый результат во всех трёх местах.

### Этап 4. Починка корректности (P-01, P-02, P-06)
1. `ModelIndex`/`all_tables`: меры со скрытых таблиц попадают в «Меры» и lineage; `№ таблицы` = `—`.
2. Фильтр системных таблиц: `startswith(("DateTableTemplate", "LocalDateTable"))`, `t.get("name")`.
3. `format_cell(value)` как чистая функция + тесты.
**DoD:** тесты P-01/P-02/P-06; на реальном `.pbit` со скрытой таблицей количество мер совпадает
с Power BI Desktop (нужен файл владельца).

### Этап 5. Разбор `extractors.py` (P-14, P-16, P-20)
1. Вынести `load_json_utf16` → `json_io.py` (импорт сохранить реэкспортом).
2. `index.py::ModelIndex` с O(1)-поиском; `dependencies.py` зависит только от `ModelIndex`
   (убрать инъекцию функции `get_column_number_func`).
3. `layout/` по файлам из §4; `LayoutExtractor` остаётся фасадом с прежними публичными методами.
4. `lineage.py`: `EdgeBuilder` + пять сборщиков рёбер; выход идентичен (`drop_duplicates` сохраняется).
5. Обновить `COMMON_HIDDEN_IMPORTS` в `build_exe.py`.
**DoD:** эталон этапа 0 без изменений; замер времени lineage до/после на файле ≥3000 мер
(числа внести в этот план); `pytest` зелёный; собранный EXE проходит `--self-test`.

### Этап 6. Разбор `gui.py` (P-15, P-22)
1. `ui/tables.py`: `CustomTableWidget`, `TableFactory`, `format_cell`.
2. `ui/search.py`, `ui/lineage_panel.py` (`filter_dfs(dfs, focus)` — чистая функция, без Qt),
   `ui/worker.py`, `ui/actions.py`, `ui/styles.py`.
3. `ui/window.py` — только композиция; `_apply_lineage_focus` ≤ 25 строк.
**DoD:** `filter_dfs` покрыт тестами (фокус на таблицу/меру/визуал/несуществующий ID);
`run_gui.py --self-test` = 0; ручной smoke: открыть `.pbit`, применить фильтр, сбросить, экспорт.

### Этап 7. Форматтеры и Excel (P-11, P-17)
1. Один `_spacing`, один `_render`; `format_column/measure/table` — делегаты.
2. `excel_report`: стили и merge без `except Exception`; ширины и `auto_filter` через константы.
3. Справка: экранирование HTML в `_md_to_html`, тест на `<script>` во входном тексте.
**DoD:** golden-тесты форматтеров (10-15 DAX-выражений: `VAR/RETURN`, вложенные `CALCULATE`,
строки с `&`, комментарии `--`); Excel совпадает с эталоном.

### Этап 8. Сборка, конфигурация, репозиторий (P-27, P-28, P-29, P-30)
1. `version.py::__version__`, `build_exe.py` сверяет с CHANGELOG и падает при расхождении.
2. `pyproject.toml`: `[project]` + `[project.optional-dependencies].dev`; `requirements.txt` →
   `-e .` или удалить (решение владельца).
3. `ruff` с расширенным набором правил + `pytest-cov` в CI, порог 60%.
4. **Согласовать** P-29: превратить `D:/git/PBIT-Analyzer` в собственный репозиторий
   или перенести дерево в клон GitHub-репозитория.
5. Актуализировать `CHANGELOG.md`/`PBIX_PLAN.md`; архивные планы — в `plan/archive/`.
**DoD:** CI собирает оба EXE и гоняет `--self-test`; локальная сборка оставляет только EXE в корне;
`git status` в папке проекта показывает только осознанно изменённые файлы.

---

## 6. Инварианты (нельзя менять ни на одном этапе)

1. Имена ключей `dfs`: `tables_ref, columns, sources, relationships, measures, sheets, visuals,
   conditional, lineage, lineage_upstream, lineage_downstream, stats, dax, power_query,
   m_parameters, rls, ols, perspectives, statistics, warnings, excel_path, file_kind, model_size`.
2. Наборы вкладок и листов для `pbit`/`pbix` — ровно как сейчас (включая отсутствие «Перспектив»,
   «DAX» и «Power Query»).
3. Заголовки колонок с «№ таблицы», «№ столбца/меры», эмодзи `📊/📋/📐`, формат `M1`, `A0102`, `s3`,
   `VIS_ID :<id>` — это контракт lineage и тестов.
4. Коды возврата CLI: `0` успех, `2` ошибка аргументов/неподдерживаемый фокус, `3` битый/чужой файл.
5. Имена файлов сборки: `PBIT_Analyzer_GUI_v<версия>.exe`, `PBIT_Analyzer_CLI_v<версия>.exe`;
   после сборки в корне остаётся только EXE.
6. `--self-test` у GUI и CLI продолжает работать без дисплея.
7. Ошибки пользователю — по-русски, без трейсбеков в диалоге.

---

## 7. Риски и то, что требует отдельного согласования

| Риск | Вероятность | Влияние | Смягчение |
|---|---|---|---|
| Переезд пакетов ломает hidden-imports EXE | высокая | EXE падает у пользователя | обновлять `build_exe.py` в том же коммите; обязательный `--self-test` после сборки |
| Изменение порядка колонок ломает Excel у потребителей | средняя | отчёты «съезжают» | golden-тест заголовков (этап 0) |
| Разбиение `gui.py` меняет поведение фильтров | средняя | lineage показывает не то | `filter_dfs` тестируется без Qt, ручной smoke |
| P-01 меняет состав вкладки «Меры» для `.pbit` | высокая (это и есть цель) | новые строки в отчёте | согласовать как исправление, отметить в CHANGELOG |
| Расширение правил `ruff` даёт большой diff | высокая | шум в ревью | отдельный коммит «только форматирование/импорты» |
| Перенос рабочего дерева в git-репозиторий (P-29) | — | потеря несохранённых правок | делать после коммита текущего состояния, с бэкапом папки |
| Реальные `.pbix`/`.pbit` для замеров и тестов | — | производительность не проверить | нужен файл от владельца (без него — синтетика на 5000 мер) |

**Отдельного «да» требуют:** P-01 (меняет содержимое отчёта), P-29 (git-структура),
P-28 (перенос зависимостей в `pyproject.toml`), P-17 (возможная новая зависимость),
этап 6 (крупный переезд GUI).

---

## 8. Как проверять каждый этап

```bash
# 1. Линтер
py -3 -m ruff check .

# 2. Тесты
py -3 -m pytest -q

# 3. Покрытие (после этапа 8)
py -3 -m pytest --cov=pbi_modules --cov-report=term-missing

# 4. Сборка и проверка бандла
py -3 build_exe.py --gui
./PBIT_Analyzer_GUI_v*.exe --self-test
py -3 build_exe.py --cli
./PBIT_Analyzer_CLI_v*.exe --self-test

# 5. Контракт Excel (этап 0)
py -3 -m pytest tests/test_excel_contract.py -q

# 6. Ручной smoke (после этапов 5-6)
#    .pbit: анализ → вкладки → фильтр S1/M1 → сброс → экспорт → «Открыть Excel»
#    .pbix: анализ → RLS/OLS/Параметры M → фильтр меры → экспорт
#    битый файл и live-connection .pbix → понятное сообщение, код 3
```

Прогон перед коммитом — в чистом клоне (как делалось ранее), чтобы не тащить артефакты
`build/`, `dist/`, `*.spec`, `__pycache__`.

---

## 9. Порядок и оценка

| Этап | Содержание | Приоритет | Оценка |
|---|---|---|---|
| 0 | golden-тесты контракта | P0 | 0.5 |
| 1 | мелкие правки и уборка черновиков | P1 | 0.5 |
| 2 | `schemas.py`: колонки, вкладки, листы | P1 | 1.0 |
| 3 | `focus.py`, `ids.py` | P1 | 0.5 |
| 4 | P-01/P-02/P-06 (корректность) | P0 | 1.0 |
| 5 | `extract/` + `lineage.py` | P1 | 2.0 |
| 6 | `ui/` | P1 | 2.0 |
| 7 | форматтеры + Excel + справка | P2 | 1.0 |
| 8 | сборка, конфиг, репозиторий | P2 | 1.0 |

Оценка в «сессиях» (одна сессия ≈ один согласованный блок правок с прогоном тестов).
Рекомендуемый порядок: **0 → 1 → 4 → 2 → 3 → 5 → 6 → 7 → 8** (сначала корректность и контракт,
потом структура).

---

## 10. Чек-лист ревьюера

- [ ] `pytest -q` зелёный, число тестов не уменьшилось.
- [ ] `ruff check .` без замечаний.
- [ ] Эталон этапа 0 не изменился (листы, колонки, порядок).
- [ ] Публичные имена импортируются из прежних путей (реэкспорт).
- [ ] `build_exe.py` содержит новые модули в hidden-imports.
- [ ] EXE проходит `--self-test`, после сборки в корне только EXE.
- [ ] В diff нет `build/`, `dist/`, `*.spec`, `__pycache__`, `*.pyc`.
- [ ] CHANGELOG отражает фактическое поведение (P-30).
- [ ] Ни один этап не тронул разделы §6 (инварианты).

---

## Приложение A. Инвентарь: где что лежит сейчас

| Что ищем | Где |
|---|---|
| Наборы вкладок GUI | `gui.py:52` (`TAB_MAPPING`), `gui.py:66` (`TAB_MAPPING_PBIX`), `gui.py:87` (`tab_mapping_for`) |
| Наборы листов Excel | `excel_report.py:24` (`SHEETS_ORDER`), `excel_report.py:41` (`SHEETS_ORDER_PBIX`), `excel_report.py:59` (`sheets_for`) |
| Схемы колонок | `excel_report.py:136` (`_get_columns`), `pbix_extractor.py:60-78` (`*_SCHEMA`) |
| Нормализация focus | `gui.py:648`, `lineage.py:85`, `run_cli.py:126`, `gui.py:885` (вложенная), `gui.py:641` (`_display_focus_id`) |
| Номера столбцов | `extractors.py:164`, `pbix_extractor.py:261` |
| Мёртвый код | `extractors.py:586` (`_extract_from_prototype_query`, 35 строк), `extractors.py:699` (`_extract_filters`, 47 строк), `help_dialog.py:73` (`set_help_text`) |
| Отладка в проде | `extractors.py:982`, `1028`, `1060-1063` |
| Дубли в форматтерах | `formatters.py:100-150` ↔ `174-215`; `formatters.py:35-75` ↔ `284-333` |
| Промежуточные файлы сборки | `build_exe.py:232` (`cleanup_build_artifacts`), `build_exe.py:157` (`clean`) |

## Приложение B. Что уже сделано (не переделывать)

- `.gitignore` переведён в UTF-8, `dist/`, `build/`, `*.spec`, `__pycache__`, `*.pyc` игнорируются;
  `__pycache__` убран из индекса.
- `create_exe.bat` — обёртка над `build_exe.py`; `tasks.bat` удалён.
- После сборки остаются только EXE в корне (`dist/` удаляется; сохраняется для `--no-copy`/CI).
- Справка в EXE находит `docs/schema_documentation.md` и `docs/CHANGELOG.md` (ищет `sys._MEIPASS`).
- Кнопки «ОК»/«Открыть Excel» в итоговом диалоге.
- Вкладка/лист «Перспективы» убраны; «DAX» и «Power Query» объединены с «Мерами»/«Столбцами»/«Источниками».
- `Layout` в `Report/` поддерживается (`file_kind`, `analyzer`).
