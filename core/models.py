"""Canonical enums and dataclasses for the v2 architecture."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from enum import Enum

from core.periods import parse_period


class EntityType(str, Enum):
    """Supported reporting entity scopes."""

    CONSOLIDATED = "consolidated"
    INDIVIDUAL = "individual"


class SourceType(str, Enum):
    """Supported source families."""

    CVM_CSV = "CVM_CSV"
    CVM_EXCEL = "CVM_EXCEL"
    CVM_ANALYSIS = "CVM_ANALYSIS"
    PDF = "PDF"


@dataclass(slots=True)
class FinancialAccount:
    code: str | None
    description: str
    value: Decimal
    period: str
    section: str
    source: SourceType
    confidence: float = 1.0

    def __post_init__(self) -> None:
        self.code = _normalize_optional_text(self.code)
        self.description = _require_text(self.description, field_name="description")
        self.value = _coerce_decimal(self.value)
        self.period = parse_period(self.period)
        self.section = _require_text(self.section, field_name="section").upper()
        self.source = _coerce_enum(self.source, SourceType)
        self.confidence = _validate_confidence(self.confidence, field_name="confidence")


@dataclass(slots=True)
class FinancialDataSet:
    company: str
    cnpj: str | None
    period: str
    entity_type: EntityType
    source_type: SourceType
    accounts: list[FinancialAccount]

    def __post_init__(self) -> None:
        self.company = _require_text(self.company, field_name="company")
        self.cnpj = _normalize_optional_text(self.cnpj)
        self.period = parse_period(self.period)
        self.entity_type = _coerce_enum(self.entity_type, EntityType)
        self.source_type = _coerce_enum(self.source_type, SourceType)
        self.accounts = list(self.accounts)

        if not self.accounts:
            raise ValueError("accounts must contain at least one FinancialAccount.")
        if not all(isinstance(account, FinancialAccount) for account in self.accounts):
            raise TypeError("accounts must contain only FinancialAccount instances.")


@dataclass(slots=True)
class MappingResult:
    spread_row: int
    label: str
    source_account: FinancialAccount | None
    value: Decimal | None
    confidence: float
    layer: int

    def __post_init__(self) -> None:
        self.spread_row = _coerce_positive_int(self.spread_row, field_name="spread_row")
        self.label = _require_text(self.label, field_name="label")
        if self.source_account is not None and not isinstance(
            self.source_account, FinancialAccount
        ):
            raise TypeError("source_account must be a FinancialAccount or None.")
        if self.value is not None:
            self.value = _coerce_decimal(self.value)
        self.confidence = _validate_confidence(self.confidence, field_name="confidence")
        self.layer = _coerce_layer(self.layer)


def _coerce_decimal(value: Decimal | int | float | str) -> Decimal:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            raise ValueError("Numeric values cannot be empty strings.")
        return Decimal(stripped)
    raise TypeError(f"Unsupported numeric type: {type(value)!r}")


def _coerce_enum(
    value: object,
    enum_type: type[EntityType] | type[SourceType],
) -> EntityType | SourceType:
    if isinstance(value, enum_type):
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            raise ValueError(f"{enum_type.__name__} cannot be empty.")
        try:
            return enum_type[stripped.upper()]
        except KeyError:
            pass
        for candidate in (stripped, stripped.lower(), stripped.upper()):
            try:
                return enum_type(candidate)
            except ValueError:
                continue
    raise ValueError(f"Unsupported {enum_type.__name__}: {value!r}")


def _coerce_positive_int(value: int, *, field_name: str) -> int:
    try:
        coerced = int(value)
    except (TypeError, ValueError) as exc:
        raise TypeError(f"{field_name} must be an integer.") from exc
    if coerced <= 0:
        raise ValueError(f"{field_name} must be greater than zero.")
    return coerced


def _coerce_layer(value: int) -> int:
    layer = _coerce_positive_int(value, field_name="layer")
    if layer not in {1, 2, 3}:
        raise ValueError("layer must be one of 1, 2, or 3.")
    return layer


def _normalize_optional_text(value: str | None) -> str | None:
    if value is None:
        return None
    stripped = value.strip()
    return stripped or None


def _require_text(value: str, *, field_name: str) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{field_name} must be a string.")
    stripped = value.strip()
    if not stripped:
        raise ValueError(f"{field_name} cannot be empty.")
    return stripped


def _validate_confidence(value: float, *, field_name: str) -> float:
    try:
        confidence = float(value)
    except (TypeError, ValueError) as exc:
        raise TypeError(f"{field_name} must be numeric.") from exc
    if not 0.0 <= confidence <= 1.0:
        raise ValueError(f"{field_name} must be between 0.0 and 1.0.")
    return confidence


__all__ = [
    "EntityType",
    "SourceType",
    "FinancialAccount",
    "FinancialDataSet",
    "MappingResult",
]
