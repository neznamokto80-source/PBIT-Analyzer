"""Сборка строк визуальных элементов (одно поле — одна строка)."""
from __future__ import annotations

import json
import logging

from ...schemas import frame_for

logger = logging.getLogger(__name__)


class VisualsMixin:
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
                    except (json.JSONDecodeError, TypeError, ValueError):
                        logger.warning("Визуал с нечитаемым config пропущен")
                        config = {}

                visual_id = config.get("name", "")
                if not visual_id:
                    x = vc.get("x", 0)
                    y = vc.get("y", 0)
                    visual_id = f"x{x}y{y}"
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
        logger.debug("extract_visuals: строк %s", len(rows))
        return frame_for(rows, "visuals")
