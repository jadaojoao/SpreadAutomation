"""Validation modules for Mode 1A Orchestration."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Iterable

from core.models import MappingResult


@dataclass
class ValidationMatch:
    is_valid: bool
    missing_labels: list[str]
    discrepancy: Decimal


class CompletenessValidator:
    """Checks if any target row failed to receive a mapped value."""

    def validate(self, schema_targets: list[dict], mapped_results: Iterable[MappingResult]) -> list[str]:
        mapped_rows = {res.spread_row for res in mapped_results if res.spread_row is not None and res.value is not None}
        
        missing_labels = []
        for target in schema_targets:
            row = target.get("row")
            if row and row not in mapped_rows:
                missing_labels.append(target.get("label", f"Row {row}"))
                
        return missing_labels


class ConsistencyValidator:
    """Validates basic accounting equation: Assets = Liabilities + Equity."""

    def validate(self, mapped_results: Iterable[MappingResult]) -> Decimal:
        """
        Calculates discrepancy between total assets and total liabilities + equity.
        Returns 0 if perfect match, or the difference.
        Note: Exact labels depend on the schema mapping, this is a heuristic approach based on CVM labels or mapped values.
        """
        # For simplicity, we search for root lines representing these totals.
        total_assets = Decimal("0.0")
        total_liab_eq = Decimal("0.0")
        
        for res in mapped_results:
            label = res.label.lower()
            if "total ativo" in label or label == "ativo total":
                total_assets = res.value or Decimal("0.0")
            elif "total passivo" in label or "passivo e patrim" in label:
                total_liab_eq = res.value or Decimal("0.0")
                
        return total_assets - total_liab_eq


class ValidationReporter:
    """Consolidates and reports the validation state."""

    def report(self, schema_targets: list[dict], mapped_results: Iterable[MappingResult]) -> ValidationMatch:
        completeness = CompletenessValidator()
        consistency = ConsistencyValidator()

        missing = completeness.validate(schema_targets, mapped_results)
        discrepancy = consistency.validate(mapped_results)
        
        is_valid = (len(missing) == 0) and (abs(discrepancy) < Decimal("1.0")) # Tolerating tiny rounding refs
        
        return ValidationMatch(
            is_valid=is_valid,
            missing_labels=missing,
            discrepancy=discrepancy
        )
