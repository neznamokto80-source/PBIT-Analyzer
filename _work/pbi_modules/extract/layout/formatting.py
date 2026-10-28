"""Форматирование и расположение визуала."""
from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


class FormattingMixin:
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
