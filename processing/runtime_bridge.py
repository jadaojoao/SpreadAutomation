"""Bridges the legacy processing runtime to the canonical schema and registry."""

from __future__ import annotations

from collections import defaultdict
from functools import lru_cache

from core.schema import SpreadSchema
from core.utils import col_txt_to_idx
from mapping import MappingRegistry


@lru_cache(maxsize=1)
def spread_schema() -> SpreadSchema:
    return SpreadSchema.load()


@lru_cache(maxsize=1)
def mapping_registry() -> MappingRegistry:
    return MappingRegistry.load()


@lru_cache(maxsize=1)
def layer1_codes_by_label() -> dict[str, tuple[str, ...]]:
    schema = spread_schema()
    registry = mapping_registry()
    row_key_to_label = {
        row_key: row.label.strip()
        for row_key, row in schema.rows.items()
    }
    grouped: dict[str, list[str]] = defaultdict(list)
    for code, row_key in registry.layer1_map.items():
        label = row_key_to_label.get(row_key)
        if label is None:
            continue
        grouped[label].append(code)
    return {label: tuple(codes) for label, codes in grouped.items()}


def layer1_codes_for_label(label: str | None) -> tuple[str, ...]:
    if not label:
        return ()
    return layer1_codes_by_label().get(str(label).strip(), ())


def label_column_1based() -> int:
    return col_txt_to_idx(spread_schema().column_for("label")) + 1


def skip_rows() -> set[int]:
    return set(spread_schema().skip_rows)


def special_row(name: str) -> int:
    return spread_schema().special_rows[name]


def spread_sheet_name() -> str:
    return spread_schema().sheet_name


def spread_start_row(start_row: int | None = None) -> int:
    return spread_schema().data_start_row if start_row is None else start_row


def data_columns(*, include_quarterly: bool = False) -> tuple[str, ...]:
    schema = spread_schema()
    columns = list(schema.columns.annual)
    if include_quarterly:
        columns.append(schema.columns.quarterly)
    return tuple(columns)


def next_data_column(current: str, *, include_quarterly: bool = False) -> str:
    normalized = current.strip().upper()
    columns = data_columns(include_quarterly=include_quarterly)
    try:
        index = columns.index(normalized)
    except ValueError as exc:
        raise ValueError(f"Column {current!r} is not part of the configured spread grid.") from exc
    if index + 1 >= len(columns):
        raise ValueError(f"Column {current!r} has no next slot in the configured spread grid.")
    return columns[index + 1]


__all__ = [
    "data_columns",
    "label_column_1based",
    "layer1_codes_for_label",
    "mapping_registry",
    "next_data_column",
    "skip_rows",
    "special_row",
    "spread_schema",
    "spread_sheet_name",
    "spread_start_row",
]
