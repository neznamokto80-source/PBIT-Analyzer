"""Разрешение полей визуала: меры, столбцы, нормализованные имена."""
from __future__ import annotations

import logging
from typing import Any, Dict, List

from ..icons import COLUMN_ICON, MEASURE_ICON, TABLE_ICON

logger = logging.getLogger(__name__)


class FieldsMixin:
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
                    results.extend(FieldsMixin._find_all_measures_in_dict(v, f"{path}.{k}" if path else k))
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                results.extend(FieldsMixin._find_all_measures_in_dict(item, f"{path}[{i}]"))
        return results

    def _extract_text_measure_rows(self, config: dict, page_name: str, visual_id: str,
                                   visual_type: str, position: str) -> List[dict]:
        rows = []
        all_measures = FieldsMixin._find_all_measures_in_dict(config)
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
