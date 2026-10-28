"""Условное форматирование визуалов из Layout."""
from __future__ import annotations

import json
import logging

from ...schemas import frame_for

logger = logging.getLogger(__name__)


class ConditionalMixin:
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
                    except (json.JSONDecodeError, TypeError, ValueError):
                        logger.warning("Визуал с нечитаемым config пропущен в условном форматировании")
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
                            cond_str = ConditionalMixin._simplify_condition(cond, measures_map)
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
                            cond_str = ConditionalMixin._simplify_condition(cond, measures_map)
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
                            cond_str = ConditionalMixin._simplify_condition(cond, measures_map)
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
                            cond_str = ConditionalMixin._simplify_condition(cond, measures_map)
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

        return frame_for(records, "conditional")

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
            left = ConditionalMixin._simplify_condition(cond["And"].get("Left", {}), measures_map)
            right = ConditionalMixin._simplify_condition(cond["And"].get("Right", {}), measures_map)
            return f"({left} и {right})"
        if "Or" in cond:
            left = ConditionalMixin._simplify_condition(cond["Or"].get("Left", {}), measures_map)
            right = ConditionalMixin._simplify_condition(cond["Or"].get("Right", {}), measures_map)
            return f"({left} или {right})"
        if "Comparison" in cond:
            cmp = cond["Comparison"]
            kind = cmp.get("ComparisonKind")
            left = cmp.get("Left")
            right = cmp.get("Right")
            left_str = ConditionalMixin._extract_expression_str(left, measures_map)
            right_str = ConditionalMixin._extract_expression_str(right, measures_map)
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
            inner = ConditionalMixin._simplify_condition(cond["Not"].get("Expression", {}), measures_map)
            return f"НЕ ({inner})"
        if "In" in cond:
            exprs = cond["In"].get("Expressions", [])
            vals = cond["In"].get("Values", [])
            expr_strs = []
            for e in exprs:
                expr_strs.append(ConditionalMixin._extract_expression_str(e, measures_map))
            val_strs = []
            for v_group in vals:
                v_parts = []
                for v in v_group:
                    v_parts.append(ConditionalMixin._extract_expression_str(v, measures_map))
                val_strs.append("(" + ", ".join(v_parts) + ")")
            return f"{'.'.join(expr_strs)} IN {'; '.join(val_strs)}"
        return "?"


# ============================================================
# 5. Utility
# ============================================================
