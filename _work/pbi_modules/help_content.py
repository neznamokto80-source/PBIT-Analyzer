import re
import sys
from pathlib import Path


def _md_to_html(md_text: str) -> str:
    """Простой конвертер Markdown → HTML для справки."""
    lines = md_text.split("\n")
    html_lines = []
    in_list = False
    in_table = False
    table_header_done = False

    for line in lines:
        stripped = line.strip()

        # Пустая строка — закрываем список/таблицу
        if not stripped:
            if in_list:
                html_lines.append("</ul>")
                in_list = False
            if in_table:
                html_lines.append("</table>")
                in_table = False
                table_header_done = False
            html_lines.append("")
            continue

        # Таблица
        if stripped.startswith("|") and stripped.endswith("|"):
            cells = [c.strip() for c in stripped.strip("|").split("|")]
            # Разделитель таблицы (---|---)
            if all(re.match(r"^[-:]+$", c) for c in cells):
                table_header_done = True
                continue
            if not in_table:
                html_lines.append('<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse;">')
                in_table = True
            if not table_header_done:
                html_lines.append("<tr style='background-color:#4472C4;color:white;'>")
                for c in cells:
                    html_lines.append(f"  <th>{_inline(c)}</th>")
            else:
                html_lines.append("<tr>")
                for c in cells:
                    html_lines.append(f"  <td>{_inline(c)}</td>")
            html_lines.append("</tr>")
            continue

        # Заголовки
        if stripped.startswith("# "):
            if in_list:
                html_lines.append("</ul>")
                in_list = False
            html_lines.append(f"<h1>{_inline(stripped[2:])}</h1>")
            continue
        if stripped.startswith("## "):
            if in_list:
                html_lines.append("</ul>")
                in_list = False
            html_lines.append(f"<h2>{_inline(stripped[3:])}</h2>")
            continue
        if stripped.startswith("### "):
            if in_list:
                html_lines.append("</ul>")
                in_list = False
            html_lines.append(f"<h3>{_inline(stripped[4:])}</h3>")
            continue
        if stripped.startswith("#### "):
            if in_list:
                html_lines.append("</ul>")
                in_list = False
            html_lines.append(f"<h4>{_inline(stripped[5:])}</h4>")
            continue

        # Списки
        if re.match(r"^[-*+]\s", stripped):
            if not in_list:
                html_lines.append("<ul>")
                in_list = True
            content = re.sub(r"^[-*+]\s+", "", stripped)
            html_lines.append(f"  <li>{_inline(content)}</li>")
            continue

        # Обычный текст
        if in_list:
            html_lines.append("</ul>")
            in_list = False
        html_lines.append(f"<p>{_inline(stripped)}</p>")

    if in_list:
        html_lines.append("</ul>")
    if in_table:
        html_lines.append("</table>")

    return "\n".join(html_lines)


def _inline(text: str) -> str:
    """Обработка inline-разметки: **bold**, `code`, ссылки."""
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"`(.+?)`", r"<code>\1</code>", text)
    text = re.sub(r"\[(.+?)\]\((.+?)\)", r'<a href="\2">\1</a>', text)
    return text


_PROJECT_MARKERS = (
    ("schema_documentation.md",),
    ("plan", "CHANGELOG.md"),
    ("pyproject.toml",),
)


def _candidate_roots() -> list[Path]:
    """Возможные корни проекта: папка пакета, бандл PyInstaller, выше по дереву."""
    package_root = Path(__file__).resolve().parent.parent
    candidates: list[Path] = [package_root]

    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        candidates.append(Path(meipass))
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent)

    candidates.extend(package_root.parents)

    unique: list[Path] = []
    for path in candidates:
        if path not in unique:
            unique.append(path)
    return unique


def _find_project_root() -> Path:
    """Найти папку с документацией справки.

    В onefile-сборке PyInstaller файлы лежат в sys._MEIPASS, а `__file__`
    модуля — в `_MEIPASS/pbi_modules/`, поэтому родитель папки пакета
    проверяется первым: иначе справка не находит schema_documentation.md
    и plan/CHANGELOG.md.
    """
    candidates = _candidate_roots()
    for candidate in candidates:
        for marker in _PROJECT_MARKERS:
            if candidate.joinpath(*marker).exists():
                return candidate
    return candidates[0]


def load_schema_doc_html() -> str:
    """Загрузить schema_documentation.md и конвертировать в HTML."""
    root = _find_project_root()
    md_path = root / "schema_documentation.md"
    if md_path.exists():
        return _md_to_html(md_path.read_text(encoding="utf-8"))
    return "<p><i>Файл schema_documentation.md не найден.</i></p>"


def load_changelog_html() -> str:
    """Загрузить plan/CHANGELOG.md и конвертировать в HTML."""
    root = _find_project_root()
    md_path = root / "plan" / "CHANGELOG.md"
    if md_path.exists():
        return _md_to_html(md_path.read_text(encoding="utf-8"))
    return "<p><i>Файл plan/CHANGELOG.md не найден.</i></p>"


HELP_TEXT = """
<h2>PBIT Analyzer — Справка</h2>
<p><b>Описание:</b> Приложение для анализа шаблонов Power BI Template (.pbit) и семантических моделей (.pbix) с построением lineage и опциональной выгрузкой в Excel. Формат определяется автоматически.</p>

<h3>Сценарий работы:</h3>
<ol>
    <li>Нажмите <b>"Выбрать PBIT/PBIX"</b> или <b>перетащите</b> .pbit/.pbix файл в окно приложения — формат определится сам</li>
    <li>При необходимости укажите <b>папку вывода</b></li>
    <li>Настройте галочки:
        <ul>
            <li><b>Создать Excel</b> — формирует .xlsx отчет по результатам анализа</li>
            <li><b>Сохранить временные файлы</b> — не удалять распакованные служебные файлы</li>
            <li><b>Открыть папку результатов</b> — открыть каталог после завершения анализа</li>
            <li><b>Открыть Excel файл</b> — открыть .xlsx после анализа (если Excel создавался)</li>
        </ul>
    </li>
    <li>Нажмите <b>"Начать анализ"</b></li>
    <li>Просмотрите результат во вкладках справа</li>
    <li>В окне завершения нажмите <b>"Открыть Excel"</b> — отчёт откроется сразу,
        либо <b>"ОК"</b>, чтобы вернуться к результатам</li>
</ol>

<h3>Вкладки с данными (детальное описание):</h3>
<ul>
    <li><b>Справочник таблиц</b> — таблицы модели Power BI с идентификаторами, именами, количеством строк, размером, описанием и типом источника.</li>
    <li><b>Столбцы</b> — столбцы таблиц, включая тип данных, формулу DAX (если есть), зависимости от других столбцов/мер, атрибуты сортировки и скрытия.</li>
    <li><b>Источники</b> — типы источников данных (DirectQuery, Import, Calculated и т.д.) и способ создания таблиц (например, через M-запрос).</li>
    <li><b>Связи</b> — отношения между таблицами с указанием типов (один-ко-многим), активностью, направлением фильтрации.</li>
    <li><b>Меры</b> — меры DAX с формулами, форматом отображения, зависимостями от таблиц/столбцов, а также категорией (например, Time Intelligence).</li>
    <li><b>Листы</b> — страницы отчета (sheets) с количеством визуальных элементов, размерами и настройками навигации.</li>
    <li><b>Визуальные элементы</b> — поля визуалов, роли (ось, значение, легенда), ссылки на таблицы/поля, позиция (x, y, ширина, высота) и тип визуала (столбец, линейчатая, карта и т.д.).</li>
    <li><b>Линейдж</b> — базовые зависимости между объектами (таблицы → столбцы → меры → визуалы) в виде графа.</li>
    <li><b>Lineage Upstream</b> — все входящие зависимости выбранного фокуса (что влияет на выбранный объект).</li>
    <li><b>Lineage Downstream</b> — все исходящие зависимости выбранного фокуса (на что влияет выбранный объект).</li>
    <li><b>Условное форматирование</b> — найденные правила форматирования визуалов (цветовые шкалы, гистограммы, значки) с условиями и приоритетами.</li>
    <li><b>Параметры M</b> (PBIX) — параметры Power Query с описанием и выражением.</li>
    <li><b>Безопасность (RLS/OLS)</b> (PBIX) — правила row-level и object-level security по ролям.</li>
    <li><b>Статистика хранения</b> (PBIX) — кардинальность и размеры столбцов модели.</li>
    <li><b>Статистика</b> — агрегированные метрики анализа: количество таблиц, столбцов, мер, визуалов, листов, а также общий размер модели и время анализа.</li>
    <li><b>Схема PBIT</b> — справочная документация по структуре файлов PBIT (DataModelSchema, Layout).</li>
</ul>

<h3>PBIX — семантическая модель:</h3>
<ul>
    <li>PBIX читается библиотекой <b>pbixray</b> (read-only, без Power BI Desktop).</li>
    <li>В режиме PBIX скрыты вкладки отчёта — <b>Листы</b>, <b>Визуальные элементы</b>,
        <b>Условное форматирование</b>: у PBIX нет слоя отчёта.</li>
    <li>Вместо них показаны <b>Параметры M</b>, <b>Безопасность (RLS/OLS)</b>
        и <b>Статистика хранения</b>.</li>
    <li>Отдельных вкладок DAX и Power Query нет: выражения мер и вычисляемых столбцов есть
        в <b>Мерах</b> и <b>Столбцах</b>, вычисляемые таблицы и M — в <b>Источниках</b>.</li>
    <li>Lineage строится по формулам DAX и связям; фокус — таблица (<b>S…</b>) или мера (<b>M…</b>).</li>
    <li>Файлы от 500 МБ читаются через <b>on_disk</b> (memory-mapped), чтобы не занимать всю RAM.</li>
</ul>

<h3>Поиск по таблицам:</h3>
<p>Над каждой таблицей результатов расположена <b>строка поиска</b>. Введите текст для фильтрации строк по содержимомулюбой колонки. Нажмите <b>"Найти"</b> или <b>Enter</b> для применения фильтра, <b>"Сброс"</b> для очистки.</p>

<h3>Фокус для lineage:</h3>
<ul>
    <li>Поле становится доступным <b>после завершения анализа</b></li>
    <li>Поддерживается выбор из выпадающего списка, автодополнение и ручной ввод</li>
    <li>Формат отображения:
        <ul>
            <li><b>Тип элемента | ID | Название</b> — для таблиц/столбцов/мер</li>
            <li><b>Тип элемента | ID | Тип визуала</b> — для визуалов</li>
        </ul>
    </li>
    <li>Кнопка <b>"Найти lineage"</b> пересчитывает upstream/downstream только для выбранного объекта без повторного полного анализа</li>
</ul>

<h3>Форматы ID и ссылок:</h3>
<ul>
    <li>Таблицы: <b>S...</b> (с иконкой 📊 в данных визуалов)</li>
    <li>Столбцы: <b>A...</b> (с иконкой 📋)</li>
    <li>Меры: <b>M...</b> (с иконкой 📐)</li>
    <li>Визуалы: <b>VIS_ID :&lt;id&gt;</b></li>
</ul>

<h3>Интерпретация Lineage:</h3>
<p>Lineage (линейдж) показывает зависимости между объектами модели Power BI. Это помогает понять, как данные перемещаются от источников через таблицы и столбцы к мерам и визуальным элементам.</p>
<ul>
    <li><b>Upstream</b> (входящие зависимости) — объекты, которые влияют на выбранный фокус. Например, для меры upstream могут включать столбцы и таблицы, используемые в её формуле.</li>
    <li><b>Downstream</b> (исходящие зависимости) — объекты, которые зависят от выбранного фокуса. Например, для столбца downstream могут включать меры и визуалы, которые его используют.</li>
    <li>Граф отображается в виде таблицы с колонками: тип, ID, имя, уровень удаленности.</li>
</ul>

<h3>Советы по производительности:</h3>
<ul>
    <li>Для больших .pbit файлов (более 100 МБ) анализ может занять несколько минут. Рекомендуется закрыть другие ресурсоемкие приложения.</li>
    <li>Если не нужен Excel-отчет, снимите галочку <b>"Создать Excel"</b> для ускорения.</li>
    <li>Временные файлы занимают место на диске; оставляйте галочку <b>"Сохранить временные файлы"</b> только для отладки.</li>
    <li>Lineage для сложных моделей с тысячами объектов может быть тяжелым; используйте фильтрацию по фокусу.</li>
</ul>

<h3>Ограничения:</h3>
<ul>
    <li>Для PBIX недоступны визуалы, страницы и условное форматирование, а строки данных не выгружаются — только метаданные модели.</li>
    <li>Файлы .abf и .xlsx (PowerPivot) не поддерживаются.</li>
    <li>PBIX с внешней моделью (Live Connection) не читается: в файле нет встроенной модели.</li>
    <li>Некоторые элементы Power BI (например, динамические меры, кастомные визуалы) могут быть не полностью распознаны.</li>
    <li>Lineage не учитывает зависимости через вычисляемые таблицы (Calculated Tables) в полном объеме.</li>
    <li>Экспорт в Excel может не сохранять сложное форматирование ячеек.</li>
</ul>

<h3>Часто задаваемые вопросы (FAQ):</h3>
<ul>
    <li><b>Вопрос:</b> Почему после анализа некоторые столбцы отображаются без формул?<br>
        <b>Ответ:</b> Формулы DAX извлекаются только для вычисляемых столбцов. Для обычных столбцов источник данных — это запрос M или прямое подключение.</li>
    <li><b>Вопрос:</b> Можно ли анализировать несколько файлов одновременно?<br>
        <b>Ответ:</b> Нет, приложение работает с одним файлом за раз. Для пакетной обработки используйте скрипты.</li>
    <li><b>Вопрос:</b> Как интерпретировать ID визуалов?<br>
        <b>Ответ:</b> ID визуалов генерируются Power BI и являются уникальными в пределах файла. Они используются для ссылок в lineage.</li>
    <li><b>Вопрос:</b> Почему lineage для некоторых объектов пустой?<br>
        <b>Ответ:</b> Возможно, объект не имеет зависимостей (например, таблица-источник без вычислений) или зависимости не удалось извлечь из-за сложности модели.</li>
</ul>

<h3>Клавиши быстрого доступа:</h3>
<ul>
    <li><b>Ctrl+O</b> — открыть диалог выбора файла PBIT/PBIX</li>
    <li><b>Ctrl+S</b> — выгрузить данные в Excel</li>
    <li><b>Ctrl+H</b> — открыть справку</li>
    <li><b>Ctrl+Q</b> — выйти из приложения</li>
</ul>

<h3>Перетаскивание файлов:</h3>
<p>Вы можете перетащить файл <b>.pbit</b> или <b>.pbix</b> из проводника прямо в окно приложения. Файл автоматически будет выбран для анализа.</p>

<h3>Структура PBIT файла:</h3>
<p>PBIT файл — это ZIP-архив, содержащий JSON-файлы с метаданными модели, данные, настройки визуалов и т.д. Приложение распаковывает архив во временную папку, анализирует JSON-структуры и строит отчет.</p>

<p><b>Контакт разработчика:</b> @BDV_80</p>
<p><b>Версия:</b> 2.2</p>
"""
