import json
import logging
import re
from typing import Any, Dict, List

import pandas as pd

from .formatters import DAXFormatter, SourceFormatter

logger = logging.getLogger(__name__)
TABLE_ICON = "📊"
COLUMN_ICON = "📋"
MEASURE_ICON = "📐"


class DependencyAnalyzer:
    """Анализ зависимостей в DAX формулах.
    Формат вывода как в эталоне:
    📊 s1
    📊 s12
    📊 s2
    📊 s3

    📋 A101
    📋 A102
    ...
    """
    PAT_QUOTED = re.compile(r"'([^']+)'\[([^\]]+)\]")
    PAT_UNQUOTED = re.compile(r'(?<!\w)([A-Za-zА-Яа-яёЁ0-9_]+)\[([^\]]+)\]')
    PAT_MEASURE = re.compile(r'\[([^\]]+)\]')

    def __init__(self, table_names: set, table_number: dict,
                 column_to_table: dict, measure_to_number: dict,
                 get_column_number_func):
        self.table_names = table_names
        self.table_number = table_number
        self.column_to_table = column_to_table
        self.measure_to_number = measure_to_number
        self.get_column_number = get_column_number_func

    def find_dependencies(self, formula: str, exclude_measure: str = None) -> str:
        if not formula:
            return "—"
        formula = str(formula)
        table_nums = set()
        col_nums = set()
        measure_nums = set()

        # Собираем все таблицы и столбцы из формулы
        for tbl, col in self.PAT_QUOTED.findall(formula):
            # Таблица
            if tbl in self.table_number:
                table_nums.add(self.table_number[tbl])
            # Столбец
            if tbl in self.table_names and col:
                col_num = self.get_column_number(tbl, col)
                if col_num:
                    col_nums.add(col_num)
            elif col in self.column_to_table:
                real_tbl = self.column_to_table[col]
                col_num = self.get_column_number(real_tbl, col)
                if col_num:
                    col_nums.add(col_num)

        for tbl, col in self.PAT_UNQUOTED.findall(formula):
            if tbl in self.table_names:
                # Таблица
                if tbl in self.table_number:
                    table_nums.add(self.table_number[tbl])
                # Столбец
                col_num = self.get_column_number(tbl, col)
                if col_num:
                    col_nums.add(col_num)

        # Меры
        for meas in self.PAT_MEASURE.findall(formula):
            if meas != exclude_measure:
                if meas in self.measure_to_number:
                    measure_nums.add(self.measure_to_number[meas])

        # Формируем вывод
        parts = []
        # Таблицы
        if table_nums:
            sorted_tables = sorted(table_nums)
            for t in sorted_tables:
                parts.append(f"{TABLE_ICON} {str(t).upper()}")
        # Разделитель
        if col_nums or measure_nums:
            if table_nums:
                parts.append("")  # пустая строка
        # Столбцы
        if col_nums:
            sorted_cols = sorted(col_nums)
            for c in sorted_cols:
                parts.append(f"{COLUMN_ICON} {str(c).upper()}")
        # Меры
        if measure_nums:
            if col_nums:
                parts.append("")
            sorted_measures = sorted(measure_nums)
            for m in sorted_measures:
                parts.append(f"{MEASURE_ICON} {str(m).upper()}")

        return "\n".join(parts) if parts else "—"


# ============================================================
# 3. DataModelExtractor
# ============================================================

class DataModelExtractor:
    def __init__(self, model_data: dict):
        self.model = model_data.get("model", {})
        self.tables_raw = self.model.get("tables", [])
        self.tables = [
            t for t in self.tables_raw
            if not t.get("isHidden", False) and t["name"] != "DateTableTemplate"
        ]
        self.table_names = {t["name"] for t in self.tables}
        self.table_number = {}
        for idx, t in enumerate(self.tables, 1):
            self.table_number[t["name"]] = f"s{idx}"

        self.column_to_table: dict[str, str] = {}
        self.measure_to_table: dict[str, str] = {}
        self.measure_to_number: dict[str, str] = {}
        self.column_to_number: dict[str, dict] = {}
        self.norm_measure_to_name: dict[str, str] = {}
        self.norm_column_to_name: dict[str, str] = {}
        self.norm_table_to_name: dict[str, str] = {}
        self._build_lookups()

        # Инициализируем AFTER _build_lookups
        self.dep_analyzer = DependencyAnalyzer(
            self.table_names,
            self.table_number,
            self.column_to_table,
            self.measure_to_number,
            self.get_column_number,
        )

    @staticmethod
    def _normalize(s: str) -> str:
        return re.sub(r'[\s_\-]+', '', s).lower()

    def _build_lookups(self):
        m_counter = 0
        for t in self.tables:
            t_name = t["name"]
            self.norm_table_to_name[self._normalize(t_name)] = t_name
            for col in t.get("columns", []):
                c_name = col["name"]
                self.column_to_table[c_name] = t_name
                self.norm_column_to_name[self._normalize(c_name)] = c_name
            for m in t.get("measures", []):
                m_counter += 1
                m_num = f"M{m_counter}"
                m_name = m["name"]
                self.measure_to_table[m_name] = t_name
                self.measure_to_number[m_name] = m_num
                self.norm_measure_to_name[self._normalize(m_name)] = m_name

    def get_column_number(self, table_name: str, col_name: str) -> str:
        t_num = self.table_number.get(table_name)
        if not t_num:
            return ""
        t_idx = t_num[1:]
        t_cols = None
        for t in self.tables:
            if t["name"] == table_name:
                t_cols = t.get("columns", [])
                break
        if t_cols is None:
            return ""
        for ci, col in enumerate(t_cols, 1):
            if col["name"] == col_name:
                return f"A{t_idx}{ci:02d}"
        return ""

    def extract_tables_ref(self):
        rows = []
        for t in self.tables:
            rows.append({
                "№ таблицы": self.table_number[t["name"]],
                "Имя таблицы": t["name"],
                "Описание": t.get("description", "—")
            })
        return pd.DataFrame(rows)

    def extract_columns(self):
        rows = []
        for t in self.tables:
            t_num = self.table_number[t["name"]]
            for ci, col in enumerate(t.get("columns", []), 1):
                col_num = f"A{t_num[1:]}{ci:02d}"
                col_type = col.get("type", "")
                data_type = col.get("dataType", "")
                if col_type == "calculated":
                    value_type = "Вычисляемый столбец"
                elif col_type == "calculatedTableColumn":
                    value_type = "Столбец вычисляемой таблицы"
                else:
                    dt_map = {
                        "int64": "Целое число", "double": "Десятичное число",
                        "string": "Строка", "dateTime": "Дата/время",
                        "boolean": "Логическое", "currency": "Денежный",
                        "binary": "Двоичный",
                    }
                    value_type = dt_map.get(data_type, data_type)
                formula = "—"
                expr = col.get("expression")
                col_name = col["name"]
                if expr:
                    formula = DAXFormatter.format_column(col_name, expr)
                elif col_type == "calculated":
                    formula = "Нет формулы"
                deps = "—"
                if expr:
                    deps = self.dep_analyzer.find_dependencies(expr)
                description = col.get("description", "—")
                format_string = col.get("formatString", "—")
                sort_by_column = col.get("sortByColumn", "—")
                summarize_by = col.get("summarizeBy", "—")
                rows.append({
                    "№ столбца": col_num, "№ таблицы": t_num, "Таблица": t["name"],
                    "Столбец": col["name"], "Описание": description,
                    "Тип данных": data_type, "Тип значения": value_type,
                    "Формат": format_string, "Сортировка по столбцу": sort_by_column,
                    "Агрегация по умолчанию": summarize_by,
                    "Формула (если вычисляемый)": formula, "Зависимости": deps,
                })
        return pd.DataFrame(rows)

    def extract_sources(self):
        rows = []
        for t in self.tables:
            t_num = self.table_number[t["name"]]
            partitions = t.get("partitions", [])
            if not partitions:
                rows.append({
                    "№ таблицы": t_num, "Таблица": t["name"],
                    "Тип источника": "Не определён", "Как создается": "—",
                    "Зависимости": "—",
                })
                continue
            for part in partitions:
                source = part.get("source", part)
                src_type = source.get("type", part.get("sourceType", ""))
                expression = source.get("expression", part.get("expression", ""))
                partition_name = part.get("name", t["name"])
                if isinstance(expression, list):
                    expression = "".join(str(x) for x in expression)
                type_name = {
                    "m": "Power Query (M)", "calculated": "Вычисляемая таблица (DAX)",
                    "entity": "Entity", "structuredData": "Structured Data",
                }.get(src_type, src_type or "Не определён")
                deps = "—"
                if src_type == "calculated" and expression:
                    deps = self.dep_analyzer.find_dependencies(expression)
                if src_type == "calculated":
                    formatted = DAXFormatter.format_table(partition_name, expression)
                else:
                    formatted = SourceFormatter.format_source(src_type, expression)
                rows.append({
                    "№ таблицы": t_num, "Таблица": t["name"],
                    "Тип источника": type_name, "Как создается": formatted or "—",
                    "Зависимости": deps,
                })
        return pd.DataFrame(rows)

    def extract_relationships(self):
        rows = []
        for ri, r in enumerate(self.model.get("relationships", []), 1):
            from_tbl = r.get("fromTable", "")
            from_col = r.get("fromColumn", "")
            to_tbl = r.get("toTable", "")
            to_col = r.get("toColumn", "")
            is_active = r.get("isActive", True)
            from_num = self.table_number.get(from_tbl, "?")
            to_num = self.table_number.get(to_tbl, "?")
            rows.append({
                "Название связи": r.get("name", f"Связь {ri}"),
                "Откуда": f"{from_tbl}[{from_col}] ({from_num})",
                "Куда": f"{to_tbl}[{to_col}] ({to_num})",
                "Состояние": "Активна" if is_active else "Не активна",
            })
        return pd.DataFrame(rows)

    def extract_measures(self):
        rows = []
        for t in self.tables:
            t_num = self.table_number[t["name"]]
            for m in t.get("measures", []):
                m_num = self.measure_to_number[m["name"]]
                expr = m.get("expression", "")
                m_name = m["name"]
                if expr:
                    formula = DAXFormatter.format_measure(m_name, expr)
                else:
                    formula = "—"
                fmt = m.get("formatString", "")
                if not fmt:
                    fsd = m.get("formatStringDefinition", {})
                    fmt = fsd.get("expression", "") if isinstance(fsd, dict) else ""
                if not fmt:
                    for ann in m.get("annotations", []):
                        if ann.get("name") == "Format":
                            fmt = ann.get("value", "")
                            break
                fmt = fmt or "—"
                folder = m.get("displayFolder", "—")
                description = m.get("description", "—")
                deps = self.dep_analyzer.find_dependencies(expr, exclude_measure=m_name) if expr else "—"
                rows.append({
                    "№ меры": m_num, "№ таблицы": t_num, "Таблица": t["name"],
                    "Мера": m_name, "Описание": description, "Формула (DAX)": formula,
                    "Формат": fmt, "Папка отображения": folder, "Зависимости": deps,
                })
        return pd.DataFrame(rows)


# ============================================================
# 4. LayoutExtractor — расширенная версия с зависимостями
# ============================================================

class LayoutExtractor:
    """Извлечение листов и визуальных элементов из Layout."""

    def __init__(self, layout_data: dict, extractor: DataModelExtractor):
        self.sections = layout_data.get("sections", [])
        self.ext = extractor

    @staticmethod
    def _iconize_ref(ref_value: str, icon: str) -> str:
        if not ref_value or ref_value == "—":
            return "—"
        return f"{icon} {ref_value}"

    def _format_visual_refs(self, field_type: str, table_num: str, col_measure_num: str) -> tuple[str, str]:
        visual_table_ref = self._iconize_ref(table_num.upper(), TABLE_ICON) if table_num != "—" else "—"

        if field_type == "мера":
            visual_field_ref = self._iconize_ref(col_measure_num.upper(), MEASURE_ICON)
        elif field_type == "столбец":
            visual_field_ref = self._iconize_ref(col_measure_num.upper(), COLUMN_ICON)
        else:
            visual_field_ref = col_measure_num or "—"

        return visual_table_ref, visual_field_ref

    @staticmethod
    def _format_visual_id(visual_id: str) -> str:
        if not visual_id:
            return ""
        return f"VIS_ID :{visual_id}"

    def extract_sheets(self):
        rows = []
        for si, sec in enumerate(self.sections, 1):
            display_name = sec.get("displayName", sec.get("name", ""))
            name_id = sec.get("name", "")
            visuals_count = len(sec.get("visualContainers", []))
            is_hidden = sec.get("visibility", 0) == 1
            width = sec.get("width")
            height = sec.get("height")
            if width is not None and height is not None:
                dimensions = f"{width} x {height}"
            else:
                dimensions = "—"
            rows.append({
                "№ листа": si, "Имя листа": display_name,
                "ID листа": name_id, "Видимость": "Скрыт" if is_hidden else "Видим",
                "Количество визуалов": visuals_count,
                "Размеры": dimensions,
            })
        return pd.DataFrame(rows)

    def _resolve_query_ref(self, query_ref: str):
        if not query_ref or not query_ref.strip():
            return None
        ref = query_ref.strip()
        dot_idx = ref.find('.')
        if dot_idx == -1:
            field_name = ref
            table_name = ""
        else:
            table_name = ref[:dot_idx].strip()
            field_name = ref[dot_idx + 1:].strip()
        if not field_name:
            return None
        return self._try_resolve_field(field_name, table_name)

    def _try_resolve_field(self, field_name: str, table_name: str):
        norm = self.ext._normalize(field_name)

        # === 1. Сначала проверяем меру (не зависит от таблицы) ===
        # 1a) Точное совпадение меры
        if field_name in self.ext.measure_to_number:
            m_num = self.ext.measure_to_number[field_name]
            t_name = self.ext.measure_to_table.get(field_name, table_name)
            t_num = self.ext.table_number.get(t_name, "")
            return {"table_name": t_name, "field": field_name, "display": field_name,
                    "type": "мера", "table_num": t_num, "col_measure_num": m_num}
        # 1b) Нормализованное совпадение меры
        if norm in self.ext.norm_measure_to_name:
            real_name = self.ext.norm_measure_to_name[norm]
            m_num = self.ext.measure_to_number[real_name]
            t_name = self.ext.measure_to_table.get(real_name, table_name)
            t_num = self.ext.table_number.get(t_name, "")
            return {"table_name": t_name, "field": real_name, "display": real_name,
                    "type": "мера", "table_num": t_num, "col_measure_num": m_num}

        # === 2. Столбец в КОНКРЕТНОЙ таблице (приоритет!) ===
        resolved_table = table_name
        if table_name and table_name not in self.ext.table_names:
            t_norm = self.ext._normalize(table_name)
            if t_norm in self.ext.norm_table_to_name:
                resolved_table = self.ext.norm_table_to_name[t_norm]

        if resolved_table and resolved_table in self.ext.table_names:
            for t in self.ext.tables:
                if t["name"] == resolved_table:
                    cols = {c["name"] for c in t.get("columns", [])}
                    # Точное совпадение
                    if field_name in cols:
                        t_num = self.ext.table_number.get(resolved_table, "")
                        col_num = self.ext.get_column_number(resolved_table, field_name)
                        return {"table_name": resolved_table, "field": field_name, "display": field_name,
                                "type": "столбец", "table_num": t_num, "col_measure_num": col_num}
                    # Нормализованное совпадение
                    field_norm = self.ext._normalize(field_name)
                    for c_name in cols:
                        if self.ext._normalize(c_name) == field_norm:
                            t_num = self.ext.table_number.get(resolved_table, "")
                            col_num = self.ext.get_column_number(resolved_table, c_name)
                            return {"table_name": resolved_table, "field": c_name, "display": c_name,
                                    "type": "столбец", "table_num": t_num, "col_measure_num": col_num}

        # === 3. Глобальный lookup столбцов (если таблица не указана) ===
        # 3a) Точное совпадение столбца
        if field_name in self.ext.column_to_table:
            t_name = self.ext.column_to_table[field_name]
            t_num = self.ext.table_number.get(t_name, "")
            col_num = self.ext.get_column_number(t_name, field_name)
            return {"table_name": t_name, "field": field_name, "display": field_name,
                    "type": "столбец", "table_num": t_num, "col_measure_num": col_num}
        # 3b) Нормализованное совпадение столбца
        if norm in self.ext.norm_column_to_name:
            real_name = self.ext.norm_column_to_name[norm]
            t_name = self.ext.column_to_table[real_name]
            t_num = self.ext.table_number.get(t_name, "")
            col_num = self.ext.get_column_number(t_name, real_name)
            return {"table_name": t_name, "field": real_name, "display": real_name,
                    "type": "столбец", "table_num": t_num, "col_measure_num": col_num}

        # === 4. Суффикс (для queryRef вида "Entity.Prefix.field") ===
        parts = field_name.split('.')
        if len(parts) > 1:
            for start_idx in range(1, len(parts)):
                suffix = '.'.join(parts[start_idx:])
                suffix_norm = self.ext._normalize(suffix)
                # мера
                if suffix in self.ext.measure_to_number:
                    t_name = self.ext.measure_to_table.get(suffix, table_name)
                    t_num = self.ext.table_number.get(t_name, "")
                    m_num = self.ext.measure_to_number[suffix]
                    return {"table_name": t_name, "field": suffix, "display": suffix,
                            "type": "мера", "table_num": t_num, "col_measure_num": m_num}
                if suffix_norm in self.ext.norm_measure_to_name:
                    real_name = self.ext.norm_measure_to_name[suffix_norm]
                    t_name = self.ext.measure_to_table.get(real_name, table_name)
                    t_num = self.ext.table_number.get(t_name, "")
                    m_num = self.ext.measure_to_number[real_name]
                    return {"table_name": t_name, "field": real_name, "display": real_name,
                            "type": "мера", "table_num": t_num, "col_measure_num": m_num}
                # столбец — проверяем в конкретной таблице
                if suffix in self.ext.column_to_table:
                    t_name = self.ext.column_to_table[suffix]
                    t_num = self.ext.table_number.get(t_name, "")
                    col_num = self.ext.get_column_number(t_name, suffix)
                    return {"table_name": t_name, "field": suffix, "display": suffix,
                            "type": "столбец", "table_num": t_num, "col_measure_num": col_num}
                if suffix_norm in self.ext.norm_column_to_name:
                    real_name = self.ext.norm_column_to_name[suffix_norm]
                    t_name = self.ext.column_to_table[real_name]
                    t_num = self.ext.table_number.get(t_name, "")
                    col_num = self.ext.get_column_number(t_name, real_name)
                    return {"table_name": t_name, "field": real_name, "display": real_name,
                            "type": "столбец", "table_num": t_num, "col_measure_num": col_num}

        return {"table_name": table_name, "field": field_name, "display": field_name,
                "type": "неизвестно", "table_num": "", "col_measure_num": ""}

    def _extract_from_projections(self, single_vis: dict):
        """Извлекает поля из projections, используя prototypeQuery для разрешения имён."""
        fields = []
        projections = single_vis.get("projections", {})
        if not projections:
            return fields

        # Build alias -> entity map from prototypeQuery
        pq = single_vis.get("prototypeQuery", {})
        alias_map = {}
        for fr in pq.get("From", []):
            alias_map[fr.get("Name", "")] = fr.get("Entity", "")

        # Build Name -> Select item map from prototypeQuery
        select_map = {}
        for sel in pq.get("Select", []):
            name = sel.get("Name", "")
            if name:
                select_map[name] = sel

        # Column properties for display names
        col_props = single_vis.get("columnProperties", {})

        for role_name, role_items in projections.items():
            if not isinstance(role_items, list):
                continue
            for item in role_items:
                if not isinstance(item, dict):
                    continue
                qref = item.get("queryRef", "")
                if not qref:
                    continue

                # Try to find matching Select item by Name
                sel = select_map.get(qref)
                if sel:
                    # Extract field info from Select
                    field_name = None
                    entity_alias = ""
                    if "Measure" in sel:
                        meas = sel["Measure"]
                        field_name = meas.get("Property", "")
                        sr = meas.get("Expression", {}).get("SourceRef", {})
                        entity_alias = sr.get("Source", "")
                    elif "Column" in sel:
                        col = sel["Column"]
                        field_name = col.get("Property", "")
                        sr = col.get("Expression", {}).get("SourceRef", {})
                        entity_alias = sr.get("Source", "")

                    entity_name = alias_map.get(entity_alias, entity_alias)

                    if field_name and entity_name:
                        # Display name from columnProperties or Name
                        display_name = col_props.get(qref, {}).get("displayName", "")
                        if not display_name:
                            display_name = sel.get("NativeReferenceName", field_name)

                        # Resolve the field
                        resolved = self._resolve_field_with_table(field_name, entity_name)
                        if resolved:
                            resolved["role"] = role_name
                            resolved["display"] = display_name
                            fields.append(resolved)
                        else:
                            # Fallback: use raw info
                            fields.append({
                                "field": qref,
                                "display": display_name or field_name,
                                "role": role_name,
                                "table_num": "—",
                                "col_measure_num": "—",
                                "type": "неизвестно",
                                "table_name": entity_name,
                            })
                else:
                    # No matching Select item — resolve queryRef directly
                    resolved = self._resolve_query_ref(qref)
                    if resolved:
                        # Display name from columnProperties
                        dn = col_props.get(qref, {}).get("displayName", "")
                        if dn:
                            resolved["display"] = dn
                        resolved["role"] = role_name
                        fields.append(resolved)

        return fields

    def _extract_from_prototype_query(self, single_vis: dict):
        fields = []
        pq = single_vis.get("prototypeQuery")
        if not pq:
            return fields
        alias_map = {}
        for fr in pq.get("From", []):
            alias_map[fr.get("Name", "")] = fr.get("Entity", "")
        for sel in pq.get("Select", []):
            field_name = None
            role = sel.get("Name", "")
            meas = sel.get("Measure", {})
            if meas and "Property" in meas:
                field_name = meas["Property"]
            col = sel.get("Column", {})
            if col and "Property" in col:
                field_name = col["Property"]
            if not field_name:
                field_name = sel.get("Property")
            if not field_name:
                continue
            entity_name = ""
            for ref_key in ("Measure", "Column"):
                ref_obj = sel.get(ref_key, {})
                sr = ref_obj.get("Expression", {}).get("SourceRef", {})
                if sr:
                    alias = sr.get("Source", "")
                    entity_name = alias_map.get(alias, alias)
                    break
            resolved = self._resolve_field_with_table(field_name, entity_name)
            if resolved:
                resolved["role"] = role
                fields.append(resolved)
        return fields

    def _resolve_field_with_table(self, field_name: str, table_name: str):
        if not field_name:
            return None
        if field_name in self.ext.measure_to_number:
            m_num = self.ext.measure_to_number[field_name]
            t_name = self.ext.measure_to_table.get(field_name, table_name)
            t_num = self.ext.table_number.get(t_name, "")
            return {"table_name": t_name, "field": field_name, "display": field_name,
                    "type": "мера", "table_num": t_num, "col_measure_num": m_num}
        if table_name and table_name in self.ext.table_names:
            cols = None
            for t in self.ext.tables:
                if t["name"] == table_name:
                    cols = [c["name"] for c in t.get("columns", [])]
                    break
            if cols and field_name in cols:
                t_num = self.ext.table_number.get(table_name, "")
                col_num = self.ext.get_column_number(table_name, field_name)
                return {"table_name": table_name, "field": field_name, "display": field_name,
                        "type": "столбец", "table_num": t_num, "col_measure_num": col_num}
        if field_name in self.ext.column_to_table:
            t_name = self.ext.column_to_table[field_name]
            t_num = self.ext.table_number.get(t_name, "")
            col_num = self.ext.get_column_number(t_name, field_name)
            return {"table_name": t_name, "field": field_name, "display": field_name,
                    "type": "столбец", "table_num": t_num, "col_measure_num": col_num}
        return {"table_name": table_name, "field": field_name, "display": field_name,
                "type": "неизвестно", "table_num": "", "col_measure_num": ""}

    # ---------- Рекурсивный поиск мер в JSON для textbox ----------
    @staticmethod
    def _find_all_measures_in_dict(obj: Any, path: str = "") -> List[Dict]:
        results = []
        if isinstance(obj, dict):
            if "Measure" in obj:
                results.append({"path": path, "type": "Measure",
                                "property": obj["Measure"].get("Property", "?")})
            elif "Column" in obj:
                results.append({"path": path, "type": "Column",
                                "property": obj["Column"].get("Property", "?")})
            else:
                for k, v in obj.items():
                    results.extend(LayoutExtractor._find_all_measures_in_dict(v, f"{path}.{k}" if path else k))
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                results.extend(LayoutExtractor._find_all_measures_in_dict(item, f"{path}[{i}]"))
        return results

    def _extract_text_measure_rows(self, config: dict, page_name: str, visual_id: str,
                                   visual_type: str, position: str) -> List[dict]:
        rows = []
        all_measures = self._find_all_measures_in_dict(config)
        added = set()
        for m in all_measures:
            path_lower = m["path"].lower()
            role = None
            if "subtitle" in path_lower and "conditional" not in path_lower:
                role = "Subtitle"
            elif "title" in path_lower and "conditional" not in path_lower:
                role = "Title"
            else:
                continue
            property_name = m["property"]
            key = (role, property_name)
            if key in added:
                continue
            added.add(key)
            # Resolv measure
            resolved = self._try_resolve_field(property_name, "")
            table_num = resolved["table_num"] if resolved else "—"
            field_num = resolved["col_measure_num"] if resolved else "—"
            rows.append({
                "field": property_name, "display": property_name,
                "role": role, "table_num": table_num, "col_measure_num": field_num,
                "type": "мера"
            })
        return rows

    def _extract_filters(self, vc: dict) -> str:
        """Извлекает фильтры визуального элемента и возвращает строковое описание."""
        filters_str = vc.get("filters", "")
        if not filters_str:
            return "—"
        try:
            if isinstance(filters_str, str):
                filters = json.loads(filters_str)
            else:
                filters = filters_str
        except (json.JSONDecodeError, TypeError):
            return str(filters_str)

        if not isinstance(filters, list) or not filters:
            return "—"

        filter_parts = []
        for f in filters:
            expr_obj = f.get("expression", {})
            filter_type = f.get("type", "")

            # Measure filter
            if "Measure" in expr_obj:
                meas = expr_obj["Measure"]
                prop = meas.get("Property", "?")
                entity = meas.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
                filter_parts.append(f"{entity}.{prop}")
            # Column filter
            elif "Column" in expr_obj:
                col = expr_obj["Column"]
                prop = col.get("Property", "?")
                entity = col.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
                filter_parts.append(f"{entity}.{prop}")
            # Aggregation filter
            elif "Aggregation" in expr_obj:
                agg = expr_obj["Aggregation"]
                prop = agg.get("Expression", {}).get("Column", {}).get("Property", "?")
                entity = agg.get("Expression", {}).get("Column", {}).get("Expression", {}).get("SourceRef", {}).get("Entity", "")
                filter_parts.append(f"{entity}.{prop}")

            # Добавим тип фильтра
            if filter_type:
                filter_parts[-1] += f" ({filter_type})"

        return "; ".join(filter_parts) if filter_parts else "—"

    @staticmethod
    def _parse_json_if_string(value):
        if isinstance(value, str):
            try:
                return json.loads(value)
            except Exception:
                return None
        return value

    @staticmethod
    def _extract_filter_target(item: dict) -> tuple:
        expr_obj = item.get("expression", {}) if isinstance(item, dict) else {}
        entity = ""
        prop = ""
        if "Column" in expr_obj:
            col = expr_obj["Column"]
            prop = col.get("Property", "")
            entity = (
                col.get("Expression", {})
                .get("SourceRef", {})
                .get("Entity", "")
            )
        elif "Measure" in expr_obj:
            meas = expr_obj["Measure"]
            prop = meas.get("Property", "")
            entity = (
                meas.get("Expression", {})
                .get("SourceRef", {})
                .get("Entity", "")
            )
        elif "Aggregation" in expr_obj:
            agg_col = (
                expr_obj.get("Aggregation", {})
                .get("Expression", {})
                .get("Column", {})
            )
            prop = agg_col.get("Property", "")
            entity = (
                agg_col.get("Expression", {})
                .get("SourceRef", {})
                .get("Entity", "")
            )
        return entity, prop

    def _extract_filter_details(self, vc: dict, single_vis: dict) -> dict:
        details = []

        # 1) Фильтры из visualContainer.filters
        vc_filters = self._parse_json_if_string(vc.get("filters", []))
        if isinstance(vc_filters, list):
            for item in vc_filters:
                if not isinstance(item, dict):
                    continue
                entity, prop = self._extract_filter_target(item)
                expr_short = item.get("type", "")
                filter_block = item.get("filter", {})
                where = filter_block.get("Where", []) if isinstance(filter_block, dict) else []
                if where and isinstance(where[0], dict):
                    condition = where[0].get("Condition", {})
                    if condition:
                        expr_short = self._simplify_condition(condition)
                details.append({
                    "entity": entity,
                    "property": prop,
                    "expr": expr_short or "—",
                })

        # 2) Фильтр по умолчанию из singleVisual.objects.general[*].properties.filter.filter
        for general_item in single_vis.get("objects", {}).get("general", []):
            if not isinstance(general_item, dict):
                continue
            props = general_item.get("properties", {})
            filter_payload = props.get("filter", {}).get("filter", {})
            where = filter_payload.get("Where", []) if isinstance(filter_payload, dict) else []
            for where_item in where:
                if not isinstance(where_item, dict):
                    continue
                condition = where_item.get("Condition", {})
                expr_short = self._simplify_condition(condition) if condition else "—"

                # Пытаемся извлечь target из In.Expressions[0]
                entity = ""
                prop = ""
                in_obj = condition.get("In", {}) if isinstance(condition, dict) else {}
                exprs = in_obj.get("Expressions", []) if isinstance(in_obj, dict) else []
                if exprs and isinstance(exprs[0], dict):
                    expr0 = exprs[0]
                    if "Column" in expr0:
                        col = expr0["Column"]
                        prop = col.get("Property", "")
                        entity = (
                            col.get("Expression", {})
                            .get("SourceRef", {})
                            .get("Source", "")
                            or col.get("Expression", {})
                            .get("SourceRef", {})
                            .get("Entity", "")
                        )
                details.append({
                    "entity": entity,
                    "property": prop,
                    "expr": expr_short or "—",
                })

        entities = [d["entity"] for d in details if d.get("entity")]
        props = [d["property"] for d in details if d.get("property")]
        exprs = [d["expr"] for d in details if d.get("expr")]

        return {
            "count": len(details),
            "entities": "; ".join(dict.fromkeys(entities)) if entities else "—",
            "properties": "; ".join(dict.fromkeys(props)) if props else "—",
            "expr_short": "; ".join(dict.fromkeys(exprs)) if exprs else "—",
        }

    def _extract_formatting(self, single_vis: dict) -> str:
        """Извлекает форматирование визуала и возвращает строковое описание."""
        objects = single_vis.get("objects", {})
        if not objects:
            return "—"

        fmt_parts = []

        # Проходим по всем объектам форматирования
        for obj_name, obj_list in objects.items():
            if not isinstance(obj_list, list):
                continue
            for obj in obj_list:
                props = obj.get("properties", {})
                if not props:
                    continue
                for prop_name, prop_val in props.items():
                    # Исправление: проверяем тип prop_val
                    if isinstance(prop_val, dict):
                        expr = prop_val.get("expr", {})
                        if expr and isinstance(expr, dict):
                            if "Literal" in expr:
                                val = expr["Literal"].get("Value", "")
                                fmt_parts.append(f"[{obj_name}] {prop_name}={val}")
                            elif "ThemeDataColor" in expr:
                                tc = expr["ThemeDataColor"]
                                fmt_parts.append(f"[{obj_name}] {prop_name}=Theme({tc.get('ColorId')},{tc.get('Percent')})")
                    elif isinstance(prop_val, list):
                        # Если prop_val - список, пропускаем или обрабатываем по-другому
                        fmt_parts.append(f"[{obj_name}] {prop_name}=[complex value]")
                    else:
                        # Простое значение
                        fmt_parts.append(f"[{obj_name}] {prop_name}={prop_val}")

        return "; ".join(fmt_parts) if fmt_parts else "—"

    def _extract_position(self, vc: dict, config: dict) -> str:
        """Извлекает расположение визуала: x, y, w, h."""
        x = vc.get('x', '')
        y = vc.get('y', '')
        w = vc.get('width', '')
        h = vc.get('height', '')
        parts = []
        if x != '':
            parts.append(f"x:{x}")
        if y != '':
            parts.append(f"y:{y}")
        if w != '':
            parts.append(f"w:{w}")
        if h != '':
            parts.append(f"h:{h}")
        return " ".join(parts)

    def _extract_column_properties_display(self, single_vis: dict) -> dict:
        """Извлекает columnProperties — переопределённые отображаемые имена."""
        col_props = single_vis.get("columnProperties", {})
        result = {}
        for key, val in col_props.items():
            if isinstance(val, dict):
                dn = val.get("displayName", "")
                if dn:
                    result[key] = dn
        return result

    def extract_visuals(self):
        """
        Каждая строка = ОДНО ПОЛЕ одного визуала (как в эталоне).
        """
        rows = []
        for sec in self.sections:
            sheet_name = sec.get("displayName", sec.get("name", ""))
            for vc in sec.get("visualContainers", []):
                config = vc.get("config", {})
                if isinstance(config, str):
                    try:
                        config = json.loads(config)
                    except Exception:
                        config = {}

                visual_id = config.get("name", "")
                if not visual_id:
                    x = vc.get("x", 0)
                    y = vc.get("y", 0)
                    visual_id = f"x{x}y{y}"
                visual_id_label = self._format_visual_id(visual_id)
                visual_id_label = self._format_visual_id(visual_id)

                single_vis = config.get("singleVisual", {})
                visual_type = single_vis.get("visualType", "unknown")

                # columnProperties — переопределённые displayName
                col_props = self._extract_column_properties_display(single_vis)

                # Расположение
                position = self._extract_position(vc, config)
                filter_details = self._extract_filter_details(vc, single_vis)

                # Собираем поля из projections (теперь использует prototypeQuery для мер)
                fields = []
                proj_fields = self._extract_from_projections(single_vis)
                fields.extend(proj_fields)

                # Дедупликация по (role, field) среди ВСЕХ полей
                seen = set()
                unique_fields = []
                for f in fields:
                    key = (f.get("role", ""), f.get("field", ""))
                    if key not in seen:
                        seen.add(key)
                        unique_fields.append(f)
                fields = unique_fields

                # Текстовые поля (для textbox)
                if visual_type == "textbox":
                    text_fields = self._extract_text_measure_rows(config, sheet_name, visual_id, visual_type, "")
                    for tf in text_fields:
                        key = (tf.get("role", ""), tf.get("field", ""))
                        if key not in seen:
                            seen.add(key)
                            fields.append(tf)

                # Для каждого поля — строка
                if visual_id in ('d7de85c5ac4282aac955', '6e69509a76aa9b4efc3e'):
                    logger.debug("DEBUG %s: fields count = %s", visual_id, len(fields))
                    for ff in fields:
                        logger.debug(
                            "field: role=%s, field=%s",
                            ff.get("role"),
                            ff.get("field"),
                        )

                if not fields:
                    rows.append({
                        "Имя листа": sheet_name,
                        "ID визуала": visual_id_label,
                        "Тип визуала": visual_type,
                        "Роль": "—",
                        "Порядок": "—",
                        "Поле": "—",
                        "Отображаемое имя": "—",
                        "№ таблицы": "—",
                        "№ столбца/меры": "—",
                        "Расположение": position,
                        "FilterTargetEntity": filter_details["entities"],
                        "FilterTargetProperty": filter_details["properties"],
                        "FilterExpressionShort": filter_details["expr_short"],
                    })
                else:
                    for fi, f in enumerate(fields, 1):
                        role = f.get("role", "—")
                        field_type = f.get("type", "неизвестно")
                        field_orig = f.get("field", "—")
                        field_disp = f.get("display", "—")
                        table_num = f.get("table_num", "—") or "—"
                        col_measure_num = f.get("col_measure_num", "—") or "—"
                        visual_table_ref, visual_field_ref = self._format_visual_refs(
                            field_type,
                            table_num,
                            col_measure_num,
                        )

                        # Проверяем columnProperties для переопределения displayName
                        if field_orig in col_props:
                            dn = col_props[field_orig].get("displayName", "")
                            if dn:
                                field_disp = dn

                        # Debug loop
                        if visual_id in ('d7de85c5ac4282aac955', '6e69509a76aa9b4efc3e'):
                            logger.debug("LOOP fi=%s, field=%s", fi, field_orig)

                        # Для первой строки визуала — заполняем общие атрибуты
                        # Для остальных — None (чтобы потом merge_cells)
                        if fi == 1:
                            row = {
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Тип визуала": visual_type,
                                "Роль": role,
                                "Порядок": str(fi),
                                "Поле": field_orig,
                                "Отображаемое имя": field_disp,
                                "№ таблицы": visual_table_ref,
                                "№ столбца/меры": visual_field_ref,
                                "Расположение": position,
                                "FilterTargetEntity": filter_details["entities"],
                                "FilterTargetProperty": filter_details["properties"],
                                "FilterExpressionShort": filter_details["expr_short"],
                            }
                        else:
                            row = {
                                "Имя листа": None,
                                "ID визуала": None,
                                "Тип визуала": None,
                                "Роль": role,
                                "Порядок": str(fi),
                                "Поле": field_orig,
                                "Отображаемое имя": field_disp,
                                "№ таблицы": visual_table_ref,
                                "№ столбца/меры": visual_field_ref,
                                "Расположение": None,
                                "FilterTargetEntity": None,
                                "FilterTargetProperty": None,
                                "FilterExpressionShort": None,
                            }
                        rows.append(row)
        logger.debug("DEBUG extract_visuals: Total rows: %s", len(rows))
        from collections import Counter as _Counter
        id_counts = _Counter(r.get("ID визуала") for r in rows if r.get("ID визуала"))
        for vid, cnt in id_counts.most_common(5):
            logger.debug("%s: %s rows", vid, cnt)
        return pd.DataFrame(rows)

    def extract_conditional_formatting(self):
        """
        Извлекает условное форматирование из Layout.
        Колонки: Имя листа, ID визуала, Поле, Элемент, Правило, Цвет.
        """
        records = []
        for sec in self.sections:
            sheet_name = sec.get("displayName", sec.get("name", ""))
            for vc in sec.get("visualContainers", []):
                config = vc.get("config", {})
                if isinstance(config, str):
                    try:
                        config = json.loads(config)
                    except Exception:
                        continue

                visual_id = config.get("name", "")
                if not visual_id:
                    x = vc.get("x", 0)
                    y = vc.get("y", 0)
                    visual_id = f"x{x}y{y}"
                visual_id_label = self._format_visual_id(visual_id)

                single_vis = config.get("singleVisual", {})
                objects = single_vis.get("objects", {})
                prototype_query = single_vis.get("prototypeQuery", {})
                selects = prototype_query.get("Select", [])

                # Собираем map: property_name -> queryRef (для measures)
                measures_map = {}
                for sel in selects:
                    if "Measure" in sel:
                        prop = sel["Measure"].get("Property", "")
                        name = sel.get("Name", "")
                        if prop and name:
                            measures_map[prop] = name

                # --- values: background color ---
                for item in objects.get("values", []):
                    selector = item.get("selector", {})
                    props = item.get("properties", {})
                    back_color = props.get("backColor", {})
                    field = selector.get("metadata", "multiple")
                    color_expr = back_color.get("solid", {}).get("color", {})
                    conditional = color_expr.get("expr", {}).get("Conditional", {})
                    if conditional:
                        cases = conditional.get("Cases", [])
                        for case in cases:
                            cond = case.get("Condition", {})
                            value = case.get("Value", {}).get("Literal", {}).get("Value", "")
                            color = value.strip("'")
                            cond_str = self._simplify_condition(cond, measures_map)
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": field,
                                "Элемент": "Значение (фон)",
                                "Правило": cond_str,
                                "Цвет": color,
                            })
                    else:
                        color_val = color_expr.get("expr", {}).get("Literal", {}).get("Value", "")
                        if color_val:
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": field,
                                "Элемент": "Значение (фон)",
                                "Правило": "Всегда",
                                "Цвет": color_val.strip("'"),
                            })

                # --- cardTitle: font color ---
                for item in objects.get("cardTitle", []):
                    props = item.get("properties", {})
                    color_prop = props.get("color", {})
                    color_expr = color_prop.get("solid", {}).get("color", {})
                    conditional = color_expr.get("expr", {}).get("Conditional", {})
                    if conditional:
                        cases = conditional.get("Cases", [])
                        for case in cases:
                            cond = case.get("Condition", {})
                            value = case.get("Value", {}).get("Literal", {}).get("Value", "")
                            color = value.strip("'")
                            cond_str = self._simplify_condition(cond, measures_map)
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": "cardTitle",
                                "Элемент": "Заголовок",
                                "Правило": cond_str,
                                "Цвет": color,
                            })
                    else:
                        color_val = color_expr.get("expr", {}).get("Literal", {}).get("Value", "")
                        if color_val:
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": "cardTitle",
                                "Элемент": "Заголовок",
                                "Правило": "Статический",
                                "Цвет": color_val.strip("'"),
                            })

                # --- dataPoint: fill color ---
                for item in objects.get("dataPoint", []):
                    selector = item.get("selector", {})
                    props = item.get("properties", {})
                    fill = props.get("fill", {})
                    field = selector.get("metadata", "multiple")
                    color_expr = fill.get("solid", {}).get("color", {})
                    conditional = color_expr.get("expr", {}).get("Conditional", {})
                    if conditional:
                        cases = conditional.get("Cases", [])
                        for case in cases:
                            cond = case.get("Condition", {})
                            value = case.get("Value", {}).get("Literal", {}).get("Value", "")
                            color = value.strip("'")
                            cond_str = self._simplify_condition(cond, measures_map)
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": field,
                                "Элемент": "Точка данных",
                                "Правило": cond_str,
                                "Цвет": color,
                            })
                    else:
                        color_val = color_expr.get("expr", {}).get("Literal", {}).get("Value", "")
                        if not color_val:
                            theme = color_expr.get("expr", {}).get("ThemeDataColor", {})
                            if theme:
                                color_val = f"ThemeColor_{theme.get('ColorId')}_{theme.get('Percent')}"
                        if color_val:
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": field,
                                "Элемент": "Точка данных",
                                "Правило": "Статический",
                                "Цвет": color_val,
                            })

                # --- labels: background color ---
                for item in objects.get("labels", []):
                    selector = item.get("selector", {})
                    props = item.get("properties", {})
                    bg_color = props.get("backgroundColor", {})
                    field = selector.get("metadata", "multiple")
                    color_expr = bg_color.get("solid", {}).get("color", {})
                    conditional = color_expr.get("expr", {}).get("Conditional", {})
                    if conditional:
                        cases = conditional.get("Cases", [])
                        for case in cases:
                            cond = case.get("Condition", {})
                            value = case.get("Value", {}).get("Literal", {}).get("Value", "")
                            color = value.strip("'")
                            cond_str = self._simplify_condition(cond, measures_map)
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": field,
                                "Элемент": "Метка (фон)",
                                "Правило": cond_str,
                                "Цвет": color,
                            })
                    else:
                        color_val = color_expr.get("expr", {}).get("Literal", {}).get("Value", "")
                        if color_val:
                            records.append({
                                "Имя листа": sheet_name,
                                "ID визуала": visual_id_label,
                                "Поле": field,
                                "Элемент": "Метка (фон)",
                                "Правило": "Статический",
                                "Цвет": color_val.strip("'"),
                            })

        return pd.DataFrame(records) if records else pd.DataFrame(
            columns=["Имя листа", "ID визуала", "Поле", "Элемент", "Правило", "Цвет"]
        )

    @staticmethod
    def _extract_expression_str(expr: dict, measures_map: dict = None) -> str:
        if not expr:
            return "?"
        if "Literal" in expr:
            val = expr["Literal"].get("Value", "")
            if val.endswith("L") or val.endswith("D"):
                try:
                    num = float(val.rstrip("LD"))
                    return f"{num:,.0f}".replace(",", " ")
                except ValueError:
                    return val.strip("'")
            return val.strip("'")
        if "Measure" in expr:
            prop = expr["Measure"].get("Property", "?")
            if measures_map and prop in measures_map:
                return measures_map[prop]
            return prop
        if "Column" in expr:
            return expr["Column"].get("Property", "?")
        return "?"

    @staticmethod
    def _simplify_condition(cond: dict, measures_map: dict = None) -> str:
        if not cond:
            return ""
        if "And" in cond:
            left = LayoutExtractor._simplify_condition(cond["And"].get("Left", {}), measures_map)
            right = LayoutExtractor._simplify_condition(cond["And"].get("Right", {}), measures_map)
            return f"({left} и {right})"
        if "Or" in cond:
            left = LayoutExtractor._simplify_condition(cond["Or"].get("Left", {}), measures_map)
            right = LayoutExtractor._simplify_condition(cond["Or"].get("Right", {}), measures_map)
            return f"({left} или {right})"
        if "Comparison" in cond:
            cmp = cond["Comparison"]
            kind = cmp.get("ComparisonKind")
            left = cmp.get("Left")
            right = cmp.get("Right")
            left_str = LayoutExtractor._extract_expression_str(left, measures_map)
            right_str = LayoutExtractor._extract_expression_str(right, measures_map)
            if kind == 0:
                op = "="
            elif kind == 1:
                op = "<>"
            elif kind == 2:
                op = "<"
            elif kind == 3:
                op = ">"
            elif kind == 4:
                op = "<="
            elif kind == 5:
                op = ">="
            else:
                op = "?"
            return f"{left_str} {op} {right_str}"
        if "Not" in cond:
            inner = LayoutExtractor._simplify_condition(cond["Not"].get("Expression", {}), measures_map)
            return f"НЕ ({inner})"
        if "In" in cond:
            exprs = cond["In"].get("Expressions", [])
            vals = cond["In"].get("Values", [])
            expr_strs = []
            for e in exprs:
                expr_strs.append(LayoutExtractor._extract_expression_str(e, measures_map))
            val_strs = []
            for v_group in vals:
                v_parts = []
                for v in v_group:
                    v_parts.append(LayoutExtractor._extract_expression_str(v, measures_map))
                val_strs.append("(" + ", ".join(v_parts) + ")")
            return f"{'.'.join(expr_strs)} IN {'; '.join(val_strs)}"
        return "?"


# ============================================================
# 5. Utility
# ============================================================

def load_json_utf16(file_path: str) -> dict:
    encodings = ("utf-16-le", "utf-16", "utf-8-sig", "utf-8")
    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as file:
                return json.load(file)
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError(f"Не удалось прочитать JSON-файл с поддерживаемыми кодировками: {file_path}")


