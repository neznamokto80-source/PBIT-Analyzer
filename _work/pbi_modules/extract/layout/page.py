"""Страницы (листы) отчёта и отображаемый ID визуала."""
from __future__ import annotations

import logging

from ...schemas import frame_for

logger = logging.getLogger(__name__)


class PageMixin:
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
        return frame_for(rows, "sheets")

    @staticmethod
    def _format_visual_id(visual_id: str) -> str:
        if not visual_id:
            return ""
        return f"VIS_ID :{visual_id}"
