"""Typed loader for the current-runtime spread schema."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, ClassVar

from core.exceptions import SchemaError, SchemaNotFoundError
from core.resources import resource_path


@dataclass(frozen=True)
class SpreadColumns:
    label: str
    period_start: str
    annual: tuple[str, ...]
    quarterly: str
    hidden: tuple[str, ...]


@dataclass(frozen=True)
class SpreadSection:
    start_row: int
    end_row: int


@dataclass(frozen=True)
class SpreadRow:
    row: int | None
    label: str
    section: str
    duplicates: tuple[int, ...] = ()


@dataclass(frozen=True)
class SpreadSchema:
    sheet_name: str
    data_start_row: int
    columns: SpreadColumns
    metadata_rows: dict[str, int]
    special_rows: dict[str, int]
    skip_rows: tuple[int, ...]
    sections: dict[str, SpreadSection]
    rows: dict[str, SpreadRow]

    _CACHE: ClassVar[dict[Path, "SpreadSchema"]] = {}

    @classmethod
    def load(cls, path: Path | None = None) -> "SpreadSchema":
        schema_path = (path or resource_path("mapping_tables/spread_schema.json")).resolve()
        cached = cls._CACHE.get(schema_path)
        if cached is not None:
            return cached

        if not schema_path.exists():
            raise SchemaNotFoundError(f"Spread schema not found: {schema_path}")

        try:
            payload = json.loads(schema_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise SchemaError(f"Invalid spread schema JSON: {schema_path}") from exc

        schema = cls._from_payload(payload)
        cls._CACHE[schema_path] = schema
        return schema

    @classmethod
    def _from_payload(cls, payload: dict[str, Any]) -> "SpreadSchema":
        try:
            columns_payload = payload["columns"]
            sections_payload = payload["sections"]
            rows_payload = payload["rows"]
            metadata_payload = payload["metadata_rows"]
            special_payload = payload["special_rows"]
        except KeyError as exc:
            raise SchemaError(f"Missing required spread schema field: {exc.args[0]}") from exc

        columns = SpreadColumns(
            label=_expect_column_name(columns_payload["label"], field_name="columns.label"),
            period_start=_expect_column_name(
                columns_payload["period_start"], field_name="columns.period_start"
            ),
            annual=tuple(
                _expect_column_name(value, field_name="columns.annual")
                for value in _expect_list(columns_payload["annual"], field_name="columns.annual")
            ),
            quarterly=_expect_column_name(
                columns_payload["quarterly"], field_name="columns.quarterly"
            ),
            hidden=tuple(
                _expect_column_name(value, field_name="columns.hidden")
                for value in _expect_list(columns_payload["hidden"], field_name="columns.hidden")
            ),
        )

        metadata_rows = {
            key: _expect_positive_int(value, field_name=f"metadata_rows.{key}")
            for key, value in _expect_dict(metadata_payload, field_name="metadata_rows").items()
        }
        special_rows = {
            key: _expect_positive_int(value, field_name=f"special_rows.{key}")
            for key, value in _expect_dict(special_payload, field_name="special_rows").items()
        }
        sections = {
            key: SpreadSection(
                start_row=_expect_positive_int(
                    value["start_row"], field_name=f"sections.{key}.start_row"
                ),
                end_row=_expect_positive_int(
                    value["end_row"], field_name=f"sections.{key}.end_row"
                ),
            )
            for key, value in _expect_dict(sections_payload, field_name="sections").items()
        }
        rows = {
            key: SpreadRow(
                row=_expect_optional_positive_int(value.get("row"), field_name=f"rows.{key}.row"),
                label=_expect_non_empty_text(value.get("label"), field_name=f"rows.{key}.label"),
                section=_expect_non_empty_text(
                    value.get("section"), field_name=f"rows.{key}.section"
                ),
                duplicates=tuple(
                    _expect_positive_int(item, field_name=f"rows.{key}.duplicates")
                    for item in _expect_list(value.get("duplicates", []), field_name=f"rows.{key}.duplicates")
                ),
            )
            for key, value in _expect_dict(rows_payload, field_name="rows").items()
        }
        skip_rows = tuple(
            _expect_positive_int(value, field_name="skip_rows")
            for value in _expect_list(payload["skip_rows"], field_name="skip_rows")
        )

        schema = cls(
            sheet_name=_expect_non_empty_text(payload["sheet_name"], field_name="sheet_name"),
            data_start_row=_expect_positive_int(
                payload["data_start_row"], field_name="data_start_row"
            ),
            columns=columns,
            metadata_rows=metadata_rows,
            special_rows=special_rows,
            skip_rows=skip_rows,
            sections=sections,
            rows=rows,
        )
        schema._validate()
        return schema

    def row_for(self, key: str) -> int:
        row = self.rows.get(key)
        if row is None:
            raise SchemaError(f"Unknown schema row key: {key}")
        if row.row is None:
            raise SchemaError(f"Schema row key '{key}' is unresolved in the current template.")
        return row.row

    def section_rows(self, section: str) -> range:
        entry = self.sections.get(section)
        if entry is None:
            raise SchemaError(f"Unknown schema section: {section}")
        return range(entry.start_row, entry.end_row + 1)

    def column_for(self, name: str) -> str:
        scalar_columns = {
            "label": self.columns.label,
            "period_start": self.columns.period_start,
            "quarterly": self.columns.quarterly,
        }
        value = scalar_columns.get(name)
        if value is None:
            raise SchemaError(f"Unknown or non-scalar schema column: {name}")
        return value

    def _validate(self) -> None:
        if len(self.columns.annual) != 4:
            raise SchemaError("columns.annual must contain exactly four annual data columns.")
        if self.columns.period_start != self.columns.annual[0]:
            raise SchemaError("columns.period_start must match the first annual data column.")

        skip_set = set(self.skip_rows)
        special_set = set(self.special_rows.values())
        if skip_set != special_set:
            raise SchemaError("skip_rows must match the values declared in special_rows.")

        if set(self.columns.annual) & set(self.columns.hidden):
            raise SchemaError("Annual data columns cannot overlap hidden columns.")
        if self.columns.quarterly in self.columns.hidden:
            raise SchemaError("Quarterly data column cannot be hidden.")

        for name, section in self.sections.items():
            if section.start_row > section.end_row:
                raise SchemaError(
                    f"Section {name} has start_row greater than end_row: "
                    f"{section.start_row}>{section.end_row}"
                )

        for key, row in self.rows.items():
            section = self.sections.get(row.section)
            if section is None:
                raise SchemaError(f"Row {key} references unknown section: {row.section}")
            if row.row is not None and not (section.start_row <= row.row <= section.end_row):
                raise SchemaError(
                    f"Row {key}={row.row} falls outside section {row.section} "
                    f"({section.start_row}-{section.end_row})."
                )
            if row.row is not None and row.row in row.duplicates:
                raise SchemaError(f"Row {key} cannot repeat its primary row in duplicates.")


def _expect_dict(value: Any, *, field_name: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise SchemaError(f"{field_name} must be an object.")
    return value


def _expect_list(value: Any, *, field_name: str) -> list[Any]:
    if not isinstance(value, list):
        raise SchemaError(f"{field_name} must be a list.")
    return value


def _expect_non_empty_text(value: Any, *, field_name: str) -> str:
    if not isinstance(value, str):
        raise SchemaError(f"{field_name} must be a string.")
    text = value.strip()
    if not text:
        raise SchemaError(f"{field_name} cannot be empty.")
    return text


def _expect_positive_int(value: Any, *, field_name: str) -> int:
    if not isinstance(value, int):
        raise SchemaError(f"{field_name} must be an integer.")
    if value <= 0:
        raise SchemaError(f"{field_name} must be greater than zero.")
    return value


def _expect_optional_positive_int(value: Any, *, field_name: str) -> int | None:
    if value is None:
        return None
    return _expect_positive_int(value, field_name=field_name)


def _expect_column_name(value: Any, *, field_name: str) -> str:
    name = _expect_non_empty_text(value, field_name=field_name).upper()
    if not name.isalpha():
        raise SchemaError(f"{field_name} must be an Excel-style column name.")
    return name


__all__ = [
    "SpreadColumns",
    "SpreadRow",
    "SpreadSchema",
    "SpreadSection",
]
