"""Principal core mapper orchestrating Layer 1 and Layer 2 logic."""

from __future__ import annotations

from typing import Iterable
from decimal import Decimal

from core.models import FinancialDataSet, FinancialAccount, MappingResult
from mapping.registry import MappingRegistry
from mapping.layer1 import Layer1Matcher
from mapping.layer2 import Layer2Matcher


class Mapper:
    """
    Executes mapping layers according to descending confidence principle and overriding ADRs.
    According to ADR-003, Layer 2 executes before Layer 1 to maintain output compatibility.
    """

    def __init__(self, registry: MappingRegistry):
        self._registry = registry
        self._layer1_matcher = Layer1Matcher(registry)
        self._layer2_matcher = Layer2Matcher(registry)

    def map_dataset(
        self,
        target_schema: list[dict],
        current_dataset: FinancialDataSet,
        prior_dataset: FinancialDataSet | None = None,
    ) -> list[MappingResult]:
        """
        Orchestrates mapping from a standard FinancialDataSet to a spread format schema.
        target_schema should be a list of dictionaries detailing the structure of the spread, 
        e.g., {'row': 10, 'label': 'CMV Total', 'prior_value': Decimal('100.5') }.
        
        Returns:
            A list of MappingResult instances ready to be serialized to the Spread.
        """
        results: dict[int, MappingResult] = {}
        
        # Split datasets
        current_accounts = current_dataset.accounts
        prior_accounts = prior_dataset.accounts if prior_dataset else []

        for target in target_schema:
            row = target.get("row")
            label = target.get("label")
            prior_val = target.get("prior_value")
            
            if row is None or label is None:
                continue
                
            mapped_result = None
            
            # --- 1. Execute Layer 2 (Numeric Heuristic Match) according to ADR ---
            if prior_val is not None and prior_accounts:
                mapped_result = self._layer2_matcher.match_by_value(
                    target_label=label,
                    target_row=row,
                    target_prior_value=prior_val,
                    current_accounts=current_accounts,
                    prior_accounts=prior_accounts,
                )
                
            # --- 2. Execute Layer 1 (Deterministic Key Match) as fallback ---
            if mapped_result is None:
                # Find current accounts mapped to this exact label using registry layer1
                layer1_accs = [
                    acc for acc in current_accounts 
                    if acc.code and self._registry.layer1(acc.code) == label
                ]
                
                if layer1_accs:
                    total_val = sum(acc.value for acc in layer1_accs)
                    mapped_result = MappingResult(
                        spread_row=row,
                        label=label,
                        source_account=layer1_accs[0],
                        value=total_val,
                        confidence=1.0,
                        layer=1,
                    )
            
            if mapped_result:
                results[row] = mapped_result
                
        return list(results.values())
