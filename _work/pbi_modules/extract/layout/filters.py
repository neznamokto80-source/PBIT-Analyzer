"""Фильтры визуала: цели фильтра и сокращённое выражение."""
from __future__ import annotations

import json
import logging

logger = logging.getLogger(__name__)


class FiltersMixin:
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
