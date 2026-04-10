"""JSON-backed mapping registry with in-memory caching."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, ClassVar

from core.exceptions import MappingError
from core.resources import resource_path


@dataclass(frozen=True)
class MappingRegistry:
    layer1_map: dict[str, str]
    layer2_map: dict[str, str]
    synonym_map: dict[str, tuple[str, ...]]

    _CACHE: ClassVar[dict[Path, "MappingRegistry"]] = {}

    @classmethod
    def load(cls, tables_dir: Path | None = None) -> "MappingRegistry":
        base_dir = (tables_dir or resource_path("mapping_tables")).resolve()
        cached = cls._CACHE.get(base_dir)
        if cached is not None:
            return cached

        if not base_dir.exists():
            raise MappingError(f"Mapping tables directory not found: {base_dir}")

        registry = cls(
            layer1_map=_load_string_map(
                base_dir / "conta_spread_map.json", normalizer=_normalize_code
            ),
            layer2_map=_load_string_map(
                base_dir / "dre_spread_map.json", normalizer=_normalize_text
            ),
            synonym_map=_load_synonym_map(
                base_dir / "account_synonyms.json", normalizer=_normalize_text
            ),
        )
        registry._validate()
        cls._CACHE[base_dir] = registry
        return registry

    def layer1(self, cd_conta: str) -> str | None:
        return self.layer1_map.get(_normalize_code(cd_conta))

    def layer2(self, label: str) -> str | None:
        return self.layer2_map.get(_normalize_text(label))

    def synonyms(self, row_key: str) -> list[str]:
        return list(self.synonym_map.get(_normalize_text(row_key), ()))

    def _validate(self) -> None:
        _validate_lookup_keys(
            self.layer1_map,
            field_name="layer1_map",
            normalizer=_normalize_code,
        )
        _validate_lookup_keys(
            self.layer2_map,
            field_name="layer2_map",
            normalizer=_normalize_text,
        )
        _validate_lookup_keys(
            self.synonym_map,
            field_name="synonym_map",
            normalizer=_normalize_text,
        )


def _load_string_map(path: Path, *, normalizer: Callable[[str], str]) -> dict[str, str]:
    payload = _load_json_object(path)
    mapping: dict[str, str] = {}
    for raw_key, raw_value in payload.items():
        key = _expect_non_empty_text(raw_key, field_name=f"{path.name}.key")
        value = _expect_non_empty_text(raw_value, field_name=f"{path.name}.{key}")
        normalized_key = normalizer(key)
        _ensure_unique_lookup_key(mapping, normalized_key=normalized_key, raw_key=key, path=path)
        mapping[normalized_key] = value
    return mapping


def _load_synonym_map(
    path: Path, *, normalizer: Callable[[str], str]
) -> dict[str, tuple[str, ...]]:
    payload = _load_json_object(path)
    synonyms: dict[str, tuple[str, ...]] = {}
    for raw_key, raw_value in payload.items():
        row_key = _expect_non_empty_text(raw_key, field_name=f"{path.name}.key")
        if not isinstance(raw_value, list):
            raise MappingError(f"{path.name}.{row_key} must be a list of strings.")

        deduped: list[str] = []
        seen: set[str] = set()
        for index, item in enumerate(raw_value):
            synonym = _expect_non_empty_text(item, field_name=f"{path.name}.{row_key}[{index}]")
            normalized = _normalize_text(synonym)
            if normalized in seen:
                continue
            seen.add(normalized)
            deduped.append(synonym)

        normalized_key = normalizer(row_key)
        _ensure_unique_lookup_key(
            synonyms,
            normalized_key=normalized_key,
            raw_key=row_key,
            path=path,
        )
        synonyms[normalized_key] = tuple(deduped)
    return synonyms


def _load_json_object(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise MappingError(f"Mapping table not found: {path}")

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise MappingError(f"Invalid mapping table JSON: {path}") from exc

    if not isinstance(payload, dict):
        raise MappingError(f"Mapping table must be a JSON object: {path}")
    return payload


def _validate_lookup_keys(
    payload: dict[str, Any], *, field_name: str, normalizer: Callable[[str], str]
) -> None:
    normalized_map: dict[str, str] = {}
    for raw_key, value in payload.items():
        normalized_key = normalizer(raw_key)
        existing = normalized_map.get(normalized_key)
        if existing is not None and existing != raw_key:
            raise MappingError(
                f"{field_name} contains conflicting keys after normalization: "
                f"{existing!r} and {raw_key!r}"
            )
        normalized_map[normalized_key] = raw_key

        if isinstance(value, tuple):
            continue
        _expect_non_empty_text(value, field_name=f"{field_name}.{raw_key}")


def _ensure_unique_lookup_key(
    payload: dict[str, Any], *, normalized_key: str, raw_key: str, path: Path
) -> None:
    if normalized_key in payload:
        raise MappingError(
            f"{path.name} contains conflicting keys after normalization for {raw_key!r}."
        )


def _expect_non_empty_text(value: Any, *, field_name: str) -> str:
    if not isinstance(value, str):
        raise MappingError(f"{field_name} must be a string.")
    text = value.strip()
    if not text:
        raise MappingError(f"{field_name} cannot be empty.")
    return text


def _normalize_code(value: str) -> str:
    return _expect_non_empty_text(value, field_name="cd_conta").strip()


def _normalize_text(value: str) -> str:
    return " ".join(_expect_non_empty_text(value, field_name="text").split()).casefold()


__all__ = ["MappingRegistry"]
