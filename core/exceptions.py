"""Domain exceptions for the v2 core layer."""

from __future__ import annotations


class SpreadAutomationError(Exception):
    """Base class for domain-specific failures."""


class SchemaError(SpreadAutomationError):
    """Base class for schema-related failures."""


class SchemaNotFoundError(SchemaError, FileNotFoundError):
    """Raised when a required schema file cannot be located."""


class PeriodParseError(SpreadAutomationError, ValueError):
    """Raised when a raw period cannot be normalized."""


class IncompatibleSourceError(SpreadAutomationError):
    """Raised when a source cannot be handled by the requested workflow."""


class MappingError(SpreadAutomationError):
    """Raised when a mapping operation cannot be completed."""


class ScannedPDFError(SpreadAutomationError):
    """Raised when a PDF appears to be image-only or otherwise unreadable."""


__all__ = [
    "SpreadAutomationError",
    "SchemaError",
    "SchemaNotFoundError",
    "PeriodParseError",
    "IncompatibleSourceError",
    "MappingError",
    "ScannedPDFError",
]
