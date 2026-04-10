"""Layer 1 Matcher: Exact matching by CVM account code."""

from __future__ import annotations

from typing import Iterable

from core.models import FinancialAccount, MappingResult
from mapping.registry import MappingRegistry


class Layer1Matcher:
    """
    Performs exact matching based on the CVM account code (CD_CONTA).
    """

    def __init__(self, registry: MappingRegistry):
        self._registry = registry

    def match_account(self, account: FinancialAccount) -> MappingResult | None:
        """
        Attempt to find a mapping for a single account based on its code.
        If the account has no code, returns None.
        """
        if not account.code:
            return None

        # Look up the code in the registry (e.g., "1.01.01")
        # The registry maps code -> spread_label
        spread_label = self._registry.layer1(account.code)
        
        if not spread_label:
            return None

        return MappingResult(
            spread_row=0,  # Spread row needs to be resolved by SpreadReader context, setting a dummy 0 for now
            label=spread_label,
            source_account=account,
            value=account.value,
            confidence=1.0,
            layer=1,
        )

    def match_dataset(self, accounts: Iterable[FinancialAccount]) -> list[MappingResult]:
        """
        Run Layer 1 mapping over all accounts and aggregate the results.
        Multiple accounts mapping to the same label should have their values summed.
        We group them by label and then yield a combined MappingResult per label.
        Warning: Layer 1 groups by label but needs the row number eventually.
        """
        from collections import defaultdict
        from decimal import Decimal

        grouped: dict[str, list[FinancialAccount]] = defaultdict(list)
        
        for acc in accounts:
            label = self._registry.layer1(acc.code) if acc.code else None
            if label:
                grouped[label].append(acc)

        results = []
        for label, acc_list in grouped.items():
            # Sum up values mapped to this same label
            total_value = sum((acc.value for acc in acc_list), Decimal(0))
            
            # We keep the first mapped account as source_account for traceability
            # Or we could set source_account to None if aggregated, but typically 
            # returning the top-level or primary one is fine.
            results.append(
                MappingResult(
                    spread_row=0,  # Will be populated by the caller/Spread context
                    label=label,
                    source_account=acc_list[0],  # Using the first as representative
                    value=total_value,
                    confidence=1.0,
                    layer=1,
                )
            )

        return results
