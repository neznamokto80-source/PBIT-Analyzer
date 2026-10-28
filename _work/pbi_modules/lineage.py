import re
from collections import deque

import pandas as pd

from .focus import extract_visual_id, normalize_focus_id

REF_PATTERN = re.compile(r"([sS]\d+|A\d+|M\d+)")


def _iter_refs(value: str):
    if not value or value == "—":
        return
    for ref in REF_PATTERN.findall(str(value)):
        normalized_ref = str(ref).strip()
        if normalized_ref.startswith("S"):
            normalized_ref = f"s{normalized_ref[1:]}"
        yield normalized_ref


def _extract_visual_id(value: str) -> str:
    """Совместимость: логика в focus.extract_visual_id."""
    return extract_visual_id(value)


def _node_type_by_ref(ref: str) -> str:
    norm_ref = str(ref or "").strip()
    if norm_ref.startswith("S"):
        norm_ref = f"s{norm_ref[1:]}"
    if norm_ref.startswith("s"):
        return "Таблица"
    if norm_ref.startswith("A"):
        return "Столбец"
    if norm_ref.startswith("M"):
        return "Мера"
    return "Неизвестно"


def _clean_node_id(value) -> str:
    """ID узла: прочерк (мера на скрытой таблице без номера) — это не узел."""
    node_id = str(value or "").strip()
    return "" if node_id in {"—", "-"} else node_id


def _empty_lineage_df() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "Тип источника",
            "ID источника",
            "Имя источника",
            "Тип приемника",
            "ID приемника",
            "Имя приемника",
            "Тип связи",
            "Контекст",
        ]
    )


def _empty_direction_df() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "ID фокуса",
            "Тип фокуса",
            "Имя фокуса",
            "Тип источника",
            "ID источника",
            "Имя источника",
            "Тип приемника",
            "ID приемника",
            "Имя приемника",
            "Тип связи",
            "Контекст",
        ]
    )


def _normalize_focus_id(focus_id: str) -> str:
    """Совместимость: единые правила в focus.normalize_focus_id."""
    return normalize_focus_id(focus_id)


def _names_from(frame, id_column: str, name_column: str) -> dict[str, str]:
    """Словарь ID → имя для справочника объектов."""
    if frame is None or frame.empty or id_column not in frame.columns:
        return {}
    names: dict[str, str] = {}
    name_values = frame[name_column] if name_column in frame.columns else [""] * len(frame)
    for node_id, node_name in zip(frame[id_column], name_values):
        node_id_text = str(node_id).strip()
        if node_id_text:
            names[node_id_text] = str(node_name).strip() or "—"
    return names


def _visual_names(frame) -> dict[str, str]:
    """ID визуала → «лист | тип» (для подписи узла)."""
    if frame is None or frame.empty or "ID визуала" not in frame.columns:
        return {}
    names: dict[str, str] = {}
    page_values = frame["Имя листа"] if "Имя листа" in frame.columns else [""] * len(frame)
    type_values = frame["Тип визуала"] if "Тип визуала" in frame.columns else [""] * len(frame)
    for visual_id_raw, page, visual_type in zip(frame["ID визуала"], page_values, type_values):
        visual_id = _extract_visual_id(visual_id_raw)
        if not visual_id:
            continue
        page_text = str(page).strip()
        type_text = str(visual_type).strip()
        names[f"VIS:{visual_id}"] = (
            f"{page_text} | {type_text}" if page_text or type_text else visual_id
        )
    return names


class EdgeBuilder:
    """Собирает рёбра lineage: отбрасывает пустые ID и дубликаты."""

    def __init__(self, resolve_name):
        self._resolve_name = resolve_name
        self.edges: list[dict] = []

    def add(
        self,
        from_type: str,
        from_id: str,
        from_name: str,
        to_type: str,
        to_id: str,
        to_name: str,
        relation_type: str,
        context: str = "",
    ) -> None:
        if not from_id or not to_id:
            return
        self.edges.append(
            {
                "Тип источника": from_type,
                "ID источника": from_id,
                "Имя источника": from_name or "—",
                "Тип приемника": to_type,
                "ID приемника": to_id,
                "Имя приемника": to_name or "—",
                "Тип связи": relation_type,
                "Контекст": context or "—",
            }
        )

    def add_dependency(
        self,
        dep_ref: str,
        to_type: str,
        to_id: str,
        to_name: str,
        context: str,
    ) -> None:
        """Ребро «объект зависит от dep_ref»."""
        self.add(
            _node_type_by_ref(dep_ref),
            dep_ref,
            self._resolve_name(dep_ref),
            to_type,
            to_id,
            to_name,
            "Зависит от",
            context,
        )

    def frame(self) -> pd.DataFrame:
        if not self.edges:
            return _empty_lineage_df()
        return pd.DataFrame(self.edges).drop_duplicates().reset_index(drop=True)


def _add_source_edges(builder: EdgeBuilder, sources) -> None:
    """Источник данных → таблица."""
    if sources is None or sources.empty:
        return
    for position, (_, row) in enumerate(sources.iterrows(), 1):
        table_id = str(row.get("№ таблицы", "")).strip()
        table_name = str(row.get("Таблица", "")).strip()
        src_id = f"SRC:{table_id}:{position}"
        src_name = str(row.get("Тип источника", "")).strip()
        builder.add("Источник", src_id, src_name, "Таблица", table_id, table_name, "Подача данных")


def _add_column_edges(builder: EdgeBuilder, columns) -> None:
    """Таблица → столбец и зависимости вычисляемого столбца."""
    if columns is None or columns.empty:
        return
    for _, row in columns.iterrows():
        table_id = str(row.get("№ таблицы", "")).strip()
        table_name = str(row.get("Таблица", "")).strip()
        col_id = str(row.get("№ столбца", "")).strip()
        col_name = str(row.get("Столбец", "")).strip()
        builder.add("Таблица", table_id, table_name, "Столбец", col_id, col_name, "Содержит")
        for dep_ref in _iter_refs(str(row.get("Зависимости", ""))):
            builder.add_dependency(dep_ref, "Столбец", col_id, col_name, "Вычисляемый столбец")


def _add_measure_edges(builder: EdgeBuilder, measures) -> None:
    """Таблица → мера и зависимости формулы меры."""
    if measures is None or measures.empty:
        return
    for _, row in measures.iterrows():
        table_id = _clean_node_id(row.get("№ таблицы", ""))
        table_name = str(row.get("Таблица", "")).strip()
        measure_id = str(row.get("№ меры", "")).strip()
        measure_name = str(row.get("Мера", "")).strip()
        builder.add("Таблица", table_id, table_name, "Мера", measure_id, measure_name, "Содержит")
        for dep_ref in _iter_refs(str(row.get("Зависимости", ""))):
            builder.add_dependency(dep_ref, "Мера", measure_id, measure_name, "Формула меры")


def _add_visual_edges(builder: EdgeBuilder, visuals) -> None:
    """Таблица/поле → визуал."""
    if visuals is None or visuals.empty:
        return
    for _, row in visuals.iterrows():
        page_name = str(row.get("Имя листа", "")).strip()
        visual_id = _extract_visual_id(row.get("ID визуала"))
        if not visual_id:
            continue
        visual_node_id = f"VIS:{visual_id}"
        visual_name = f"{page_name}:{visual_id}" if page_name else visual_id
        role = str(row.get("Роль", "")).strip()

        for table_ref in _iter_refs(str(row.get("№ таблицы", "")).strip()):
            builder.add(
                "Таблица",
                table_ref,
                builder._resolve_name(table_ref),
                "Визуал",
                visual_node_id,
                visual_name,
                "Используется в",
                role,
            )
        for field_ref in _iter_refs(str(row.get("№ столбца/меры", "")).strip()):
            builder.add(
                _node_type_by_ref(field_ref),
                field_ref,
                builder._resolve_name(field_ref),
                "Визуал",
                visual_node_id,
                visual_name,
                "Используется в",
                role,
            )


def build_lineage_edges(dfs: dict) -> pd.DataFrame:
    """Строит рёбра lineage по справочникам и зависимостям."""
    table_names = _names_from(dfs.get("tables_ref"), "№ таблицы", "Имя таблицы")
    column_names = _names_from(dfs.get("columns"), "№ столбца", "Столбец")
    measure_names = _names_from(dfs.get("measures"), "№ меры", "Мера")
    visual_names = _visual_names(dfs.get("visuals"))

    def resolve(node_id: str) -> str:
        node_type = _node_type_by_ref(node_id)
        if node_type == "Таблица":
            return table_names.get(node_id, node_id)
        if node_type == "Столбец":
            return column_names.get(node_id, node_id)
        if node_type == "Мера":
            return measure_names.get(node_id, node_id)
        if str(node_id).startswith("VIS:"):
            return visual_names.get(node_id, node_id)
        return node_id

    builder = EdgeBuilder(resolve)
    _add_source_edges(builder, dfs.get("sources"))
    _add_column_edges(builder, dfs.get("columns"))
    _add_measure_edges(builder, dfs.get("measures"))
    _add_visual_edges(builder, dfs.get("visuals"))
    return builder.frame()


def _build_focus_nodes(edges_df: pd.DataFrame) -> list[tuple[str, str, str]]:
    focus_nodes = {}
    for _, row in edges_df.iterrows():
        source_type = str(row.get("Тип источника", ""))
        source_id = str(row.get("ID источника", ""))
        source_name = str(row.get("Имя источника", ""))
        target_type = str(row.get("Тип приемника", ""))
        target_id = str(row.get("ID приемника", ""))
        target_name = str(row.get("Имя приемника", ""))

        if source_type in {"Таблица", "Мера", "Визуал"} and source_id:
            focus_nodes[source_id] = (source_type, source_name)
        if target_type in {"Таблица", "Мера", "Визуал"} and target_id:
            focus_nodes[target_id] = (target_type, target_name)

    return [(node_id, data[0], data[1]) for node_id, data in focus_nodes.items()]


def _build_adjacency_indexes(
    edges_df: pd.DataFrame,
) -> tuple[dict[str, list[tuple[int, str]]], dict[str, list[tuple[int, str]]]]:
    adjacency_out: dict[str, list[tuple[int, str]]] = {}
    adjacency_in: dict[str, list[tuple[int, str]]] = {}

    for idx, row in edges_df.iterrows():
        src = str(row.get("ID источника", ""))
        dst = str(row.get("ID приемника", ""))
        if not src or not dst:
            continue
        adjacency_out.setdefault(src, []).append((idx, dst))
        adjacency_in.setdefault(dst, []).append((idx, src))

    return adjacency_out, adjacency_in


def _collect_edge_indexes_with_adjacency(
    focus_id: str,
    direction: str,
    adjacency_out: dict[str, list[tuple[int, str]]],
    adjacency_in: dict[str, list[tuple[int, str]]],
) -> set[int]:
    visited_nodes = {focus_id}
    edge_indexes: set[int] = set()
    queue: deque[str] = deque([focus_id])

    while queue:
        current = queue.popleft()
        if direction == "upstream":
            candidate_indexes = adjacency_in.get(current, [])
            for idx, prev_node in candidate_indexes:
                edge_indexes.add(idx)
                if prev_node and prev_node not in visited_nodes:
                    visited_nodes.add(prev_node)
                    queue.append(prev_node)
        else:
            candidate_indexes = adjacency_out.get(current, [])
            for idx, next_node in candidate_indexes:
                edge_indexes.add(idx)
                if next_node and next_node not in visited_nodes:
                    visited_nodes.add(next_node)
                    queue.append(next_node)

    return edge_indexes


def build_lineage_directions(
    edges_df: pd.DataFrame,
    focus_id: str | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if edges_df is None or edges_df.empty:
        return _empty_direction_df(), _empty_direction_df()

    normalized_focus_id = _normalize_focus_id(focus_id or "")

    if normalized_focus_id:
        focus_nodes_map = {node_id: (node_type, node_name) for node_id, node_type, node_name in _build_focus_nodes(edges_df)}
        node_meta = focus_nodes_map.get(normalized_focus_id)
        if not node_meta:
            return _empty_direction_df(), _empty_direction_df()
        focus_nodes = [(normalized_focus_id, node_meta[0], node_meta[1])]
    else:
        focus_nodes = _build_focus_nodes(edges_df)

    upstream_frames: list[pd.DataFrame] = []
    downstream_frames: list[pd.DataFrame] = []
    adjacency_out, adjacency_in = _build_adjacency_indexes(edges_df)
    edge_columns = [
        "Тип источника",
        "ID источника",
        "Имя источника",
        "Тип приемника",
        "ID приемника",
        "Имя приемника",
        "Тип связи",
        "Контекст",
    ]

    for focus_id, focus_type, focus_name in focus_nodes:
        upstream_idx = _collect_edge_indexes_with_adjacency(focus_id, "upstream", adjacency_out, adjacency_in)
        downstream_idx = _collect_edge_indexes_with_adjacency(focus_id, "downstream", adjacency_out, adjacency_in)

        if upstream_idx:
            upstream_slice = edges_df.loc[sorted(upstream_idx), edge_columns].copy()
            upstream_slice.insert(0, "Имя фокуса", focus_name or "—")
            upstream_slice.insert(0, "Тип фокуса", focus_type)
            upstream_slice.insert(0, "ID фокуса", focus_id)
            upstream_frames.append(upstream_slice)

        if downstream_idx:
            downstream_slice = edges_df.loc[sorted(downstream_idx), edge_columns].copy()
            downstream_slice.insert(0, "Имя фокуса", focus_name or "—")
            downstream_slice.insert(0, "Тип фокуса", focus_type)
            downstream_slice.insert(0, "ID фокуса", focus_id)
            downstream_frames.append(downstream_slice)

    upstream_df = pd.concat(upstream_frames, ignore_index=True) if upstream_frames else _empty_direction_df()
    downstream_df = pd.concat(downstream_frames, ignore_index=True) if downstream_frames else _empty_direction_df()

    if not upstream_df.empty:
        upstream_df = upstream_df.drop_duplicates().reset_index(drop=True)
    if not downstream_df.empty:
        downstream_df = downstream_df.drop_duplicates().reset_index(drop=True)

    return upstream_df, downstream_df
