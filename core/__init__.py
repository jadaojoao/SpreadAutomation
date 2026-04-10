"""Public core API for canonical v2 models plus legacy compatibility helpers."""

from __future__ import annotations

from importlib import import_module

from core.exceptions import (
    IncompatibleSourceError,
    MappingError,
    PeriodParseError,
    ScannedPDFError,
    SchemaError,
    SchemaNotFoundError,
    SpreadAutomationError,
)
from core.models import (
    EntityType,
    FinancialAccount,
    FinancialDataSet,
    MappingResult,
    SourceType,
)
from core.periods import is_annual, parse_period, periods_for_year
from core.resources import get_resource_path, resource_path

_LEGACY_UTIL_EXPORTS = (
    "normaliza_num",
    "periodos",
    "col_txt_to_idx",
    "shift_formula",
    "adjust_complex_formula",
    "DRE_SPREAD_MAP",
    "XLWINGS",
)

__all__ = [
    "SpreadAutomationError",
    "SchemaError",
    "SchemaNotFoundError",
    "PeriodParseError",
    "IncompatibleSourceError",
    "MappingError",
    "ScannedPDFError",
    "EntityType",
    "SourceType",
    "FinancialAccount",
    "FinancialDataSet",
    "MappingResult",
    "parse_period",
    "periods_for_year",
    "is_annual",
    "resource_path",
    "get_resource_path",
    *_LEGACY_UTIL_EXPORTS,
]


def __getattr__(name: str) -> object:
    if name in _LEGACY_UTIL_EXPORTS:
        utils_module = import_module("core.utils")
        value = getattr(utils_module, name)
        globals()[name] = value
        return value
    raise AttributeError(f"module 'core' has no attribute {name!r}")
