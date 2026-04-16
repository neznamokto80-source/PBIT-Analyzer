import re

import pandas as pd


REF_PATTERN = re.compile(r"([sS]\d+|A\d+|M\d+)")
VISUAL_ID_PATTERN = re.compile(r"VIS_ID\s*:\s*(.+)$", re.IGNORECASE)


def _iter_refs(value: str):
    if not value or value == "—":
        return
    for ref in REF_PATTERN.findall(str(value)):
        normalized_ref = str(ref).strip()
        if normalized_ref.startswith("S"):
            normalized_ref = f"s{normalized_ref[1:]}"
        yield normalized_ref


def _extract_visual_id(value: str) -> str:
    raw = str(value or "").strip()
    if not raw or raw == "None":
        return ""
    match = VISUAL_ID_PATTERN.match(raw)
    if match:
        return match.group(1).strip()
    if raw.upper().startswith("VIS:"):
        return raw[4:].strip()
    return raw


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
    if focus_id is None:
        return ""
    normalized = str(focus_id).strip()
    if not normalized:
        return ""

    normalized = normalized.replace("🔹", "").replace("📊", "").replace("📋", "").replace("📐", "").strip()
    upper_value = normalized.upper()

    if upper_value.startswith("VIS_ID"):
        visual_id = _extract_visual_id(normalized)
        return f"VIS:{visual_id}" if visual_id else ""
    if upper_value.startswith("VIS:"):
        return f"VIS:{normalized[4:].strip()}"

    if upper_value.startswith("T:"):
        return normalized[2:].strip()
    if upper_value.startswith("C:"):
        return normalized[2:].strip()
    if upper_value.startswith("M:"):
        return f"M{normalized[2:].strip().lstrip('M')}"

    if upper_value.startswith("S"):
        return f"s{normalized[1:]}"
    if upper_value.startswith("A"):
        return f"A{normalized[1:]}"
    if upper_value.startswith("M"):
        return f"M{normalized[1:]}"

    return normalized


def build_lineage_edges(dfs: dict) -> pd.DataFrame:
    edges: list[dict] = []
    table_names: dict[str, str] = {}
    column_names: dict[str, str] = {}
    measure_names: dict[str, str] = {}
    visual_names: dict[str, str] = {}

    tables_ref = dfs.get("tables_ref")
    if tables_ref is not None and not tables_ref.empty:
        for _, row in tables_ref.iterrows():
            node_id = str(row.get("№ таблицы", "")).strip()
            node_name = str(row.get("Имя таблицы", "")).strip()
            if node_id:
                table_names[node_id] = node_name or "—"

    columns_ref = dfs.get("columns")
    if columns_ref is not None and not columns_ref.empty:
        for _, row in columns_ref.iterrows():
            node_id = str(row.get("№ столбца", "")).strip()
            node_name = str(row.get("Столбец", "")).strip()
            if node_id:
                column_names[node_id] = node_name or "—"

    measures_ref = dfs.get("measures")
    if measures_ref is not None and not measures_ref.empty:
        for _, row in measures_ref.iterrows():
            node_id = str(row.get("№ меры", "")).strip()
            node_name = str(row.get("Мера", "")).strip()
            if node_id:
                measure_names[node_id] = node_name or "—"

    visuals_ref = dfs.get("visuals")
    if visuals_ref is not None and not visuals_ref.empty:
        for _, row in visuals_ref.iterrows():
            visual_id = _extract_visual_id(row.get("ID визуала"))
            page_name = str(row.get("Имя листа", "")).strip()
            visual_type = str(row.get("Тип визуала", "")).strip()
            if visual_id:
                visual_names[f"VIS:{visual_id}"] = f"{page_name} | {visual_type}" if page_name or visual_type else visual_id

    def _resolve_node_name(node_id: str) -> str:
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

    def add_edge(
        from_type: str,
        from_id: str,
        from_name: str,
        to_type: str,
        to_id: str,
        to_name: str,
        relation_type: str,
        context: str = "",
    ):
        if not from_id or not to_id:
            return
        edges.append(
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

    sources = dfs.get("sources")
    if sources is not None and not sources.empty:
        for idx, row in sources.iterrows():
            table_id = str(row.get("№ таблицы", "")).strip()
            table_name = str(row.get("Таблица", "")).strip()
            src_id = f"SRC:{table_id}:{idx + 1}"
            src_name = str(row.get("Тип источника", "")).strip()
            add_edge("Источник", src_id, src_name, "Таблица", table_id, table_name, "Подача данных")

    columns = dfs.get("columns")
    if columns is not None and not columns.empty:
        for _, row in columns.iterrows():
            table_id = str(row.get("№ таблицы", "")).strip()
            table_name = str(row.get("Таблица", "")).strip()
            col_id = str(row.get("№ столбца", "")).strip()
            col_name = str(row.get("Столбец", "")).strip()
            add_edge("Таблица", table_id, table_name, "Столбец", col_id, col_name, "Содержит")

            for dep_ref in _iter_refs(str(row.get("Зависимости", ""))):
                add_edge(
                    _node_type_by_ref(dep_ref),
                    dep_ref,
                    _resolve_node_name(dep_ref),
                    "Столбец",
                    col_id,
                    col_name,
                    "Зависит от",
                    "Вычисляемый столбец",
                )

    measures = dfs.get("measures")
    if measures is not None and not measures.empty:
        for _, row in measures.iterrows():
            table_id = str(row.get("№ таблицы", "")).strip()
            table_name = str(row.get("Таблица", "")).strip()
            measure_id = str(row.get("№ меры", "")).strip()
            measure_name = str(row.get("Мера", "")).strip()
            add_edge("Таблица", table_id, table_name, "Мера", measure_id, measure_name, "Содержит")

            for dep_ref in _iter_refs(str(row.get("Зависимости", ""))):
                add_edge(
                    _node_type_by_ref(dep_ref),
                    dep_ref,
                    _resolve_node_name(dep_ref),
                    "Мера",
                    measure_id,
                    measure_name,
                    "Зависит от",
                    "Формула меры",
                )

    visuals = dfs.get("visuals")
    if visuals is not None and not visuals.empty:
        for _, row in visuals.iterrows():
            page_name = str(row.get("Имя листа", "")).strip()
            visual_id_raw = row.get("ID визуала")
            visual_id = _extract_visual_id(visual_id_raw)
            if not visual_id:
                continue
            visual_node_id = f"VIS:{visual_id}"
            visual_name = f"{page_name}:{visual_id}" if page_name else visual_id

            table_ref_value = str(row.get("№ таблицы", "")).strip()
            field_ref_value = str(row.get("№ столбца/меры", "")).strip()
            role = str(row.get("Роль", "")).strip()

            for table_ref in _iter_refs(table_ref_value):
                add_edge(
                    "Таблица",
                    table_ref,
                    _resolve_node_name(table_ref),
                    "Визуал",
                    visual_node_id,
                    visual_name,
                    "Используется в",
                    role,
                )
            for field_ref in _iter_refs(field_ref_value):
                add_edge(
                    _node_type_by_ref(field_ref),
                    field_ref,
                    _resolve_node_name(field_ref),
                    "Визуал",
                    visual_node_id,
                    visual_name,
                    "Используется в",
                    role,
                )

    if not edges:
        return _empty_lineage_df()

    edges_df = pd.DataFrame(edges).drop_duplicates().reset_index(drop=True)
    return edges_df


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


def _collect_edge_indexes(edges_df: pd.DataFrame, focus_id: str, direction: str) -> set[int]:
    if edges_df.empty:
        return set()

    visited_nodes = {focus_id}
    edge_indexes: set[int] = set()
    queue = [focus_id]

    while queue:
        current = queue.pop(0)
        for idx, row in edges_df.iterrows():
            src = str(row.get("ID источника", ""))
            dst = str(row.get("ID приемника", ""))

            if direction == "upstream" and dst == current:
                edge_indexes.add(idx)
                if src and src not in visited_nodes:
                    visited_nodes.add(src)
                    queue.append(src)
            elif direction == "downstream" and src == current:
                edge_indexes.add(idx)
                if dst and dst not in visited_nodes:
                    visited_nodes.add(dst)
                    queue.append(dst)

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

    upstream_rows = []
    downstream_rows = []

    for focus_id, focus_type, focus_name in focus_nodes:
        upstream_idx = _collect_edge_indexes(edges_df, focus_id, "upstream")
        downstream_idx = _collect_edge_indexes(edges_df, focus_id, "downstream")

        for idx in sorted(upstream_idx):
            row = edges_df.iloc[idx].to_dict()
            row.update({"ID фокуса": focus_id, "Тип фокуса": focus_type, "Имя фокуса": focus_name or "—"})
            upstream_rows.append(row)

        for idx in sorted(downstream_idx):
            row = edges_df.iloc[idx].to_dict()
            row.update({"ID фокуса": focus_id, "Тип фокуса": focus_type, "Имя фокуса": focus_name or "—"})
            downstream_rows.append(row)

    upstream_df = pd.DataFrame(upstream_rows) if upstream_rows else _empty_direction_df()
    downstream_df = pd.DataFrame(downstream_rows) if downstream_rows else _empty_direction_df()

    if not upstream_df.empty:
        upstream_df = upstream_df.drop_duplicates().reset_index(drop=True)
    if not downstream_df.empty:
        downstream_df = downstream_df.drop_duplicates().reset_index(drop=True)

    return upstream_df, downstream_df
