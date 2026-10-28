"""Фильтрация результатов анализа по фокусу lineage (без Qt).

Чистая функция ``filter_dfs`` используется и GUI, и тестами: ей не нужен
QApplication, поэтому поведение фильтра проверяется напрямую.
"""
from __future__ import annotations

import pandas as pd

from ..focus import extract_visual_id, focus_node_type, normalize_focus_id
from ..lineage import build_lineage_directions

_LINEAGE_KEYS = {"lineage_upstream", "lineage_downstream"}


def _matches(series: pd.Series, focus_id: str) -> pd.Series:
    return series.astype(str).str.upper() == focus_id.upper()


def _related_visual_ids(lineage_df, focus_id: str) -> set[str]:
    """Визуалы, связанные с фокусом через lineage."""
    if not isinstance(lineage_df, pd.DataFrame) or lineage_df.empty:
        return set()
    if "ID источника" not in lineage_df.columns or "ID приемника" not in lineage_df.columns:
        return set()

    mask = _matches(lineage_df["ID источника"], focus_id) | _matches(
        lineage_df["ID приемника"], focus_id
    )
    visual_ids: set[str] = set()
    for _, row in lineage_df[mask].iterrows():
        if str(row.get("Тип источника", "")).strip() == "Визуал":
            source_id = str(row.get("ID источника", "")).strip()
            if source_id:
                visual_ids.add(source_id)
        if str(row.get("Тип приемника", "")).strip() == "Визуал":
            target_id = str(row.get("ID приемника", "")).strip()
            if target_id:
                visual_ids.add(target_id)
    return visual_ids


def _empty_like(frame: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(columns=frame.columns)


def _table_number_of_measure(dfs: dict, focus_id: str) -> str | None:
    measures = dfs.get("measures")
    if not isinstance(measures, pd.DataFrame) or measures.empty:
        return None
    if "№ меры" not in measures.columns or "№ таблицы" not in measures.columns:
        return None
    matched = measures.loc[_matches(measures["№ меры"], focus_id), "№ таблицы"]
    return str(matched.iloc[0]) if not matched.empty else None


def filter_dfs(dfs: dict, focus_id: str, node_type: str = "") -> dict:
    """Новый набор данных, отфильтрованный по фокусу.

    Не-DataFrame значения (``excel_path``, ``file_kind``) сохраняются как есть,
    ``lineage_upstream``/``lineage_downstream`` пересчитываются.
    """
    normalized = normalize_focus_id(focus_id)
    node_type = node_type or focus_node_type(normalized)
    visual_ids = _related_visual_ids(dfs.get("lineage"), normalized)

    filtered: dict = {}
    for key, frame in dfs.items():
        if key in _LINEAGE_KEYS:
            continue
        if not isinstance(frame, pd.DataFrame):
            filtered[key] = frame
            continue
        if frame.empty:
            filtered[key] = frame
            continue

        if key == "measures":
            filtered[key] = (
                frame[_matches(frame["№ меры"], normalized)]
                if "№ меры" in frame.columns
                else frame
            )
        elif key == "tables_ref":
            if "№ таблицы" not in frame.columns:
                filtered[key] = _empty_like(frame)
            elif node_type == "Таблица":
                filtered[key] = frame[_matches(frame["№ таблицы"], normalized)]
            else:
                table_number = _table_number_of_measure(dfs, normalized) if node_type == "Мера" else None
                filtered[key] = (
                    _empty_like(frame)
                    if table_number is None
                    else frame[frame["№ таблицы"].astype(str) == str(table_number)]
                )
        elif key == "visuals":
            if "ID визуала" in frame.columns:
                visual_keys = frame["ID визуала"].map(lambda value: f"VIS:{extract_visual_id(value)}")
                filtered[key] = frame[visual_keys.isin(visual_ids)]
            else:
                filtered[key] = _empty_like(frame)
        elif key == "lineage":
            if "ID источника" in frame.columns and "ID приемника" in frame.columns:
                filtered[key] = frame[
                    _matches(frame["ID источника"], normalized)
                    | _matches(frame["ID приемника"], normalized)
                ]
            else:
                filtered[key] = frame
        else:
            filtered[key] = _empty_like(frame)

    upstream, downstream = build_lineage_directions(filtered.get("lineage"), focus_id=normalized)
    filtered["lineage_upstream"] = upstream
    filtered["lineage_downstream"] = downstream
    return filtered


__all__ = ["filter_dfs"]
