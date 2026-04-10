"""Layer 2 Matcher: Numeric value checking."""

from __future__ import annotations

from typing import Iterable
from decimal import Decimal

from core.models import FinancialAccount, MappingResult
from mapping.registry import MappingRegistry


class Layer2Matcher:
    """
    Performs heuristic matching looking for equivalent numeric values in the prior period.
    """

    def __init__(self, registry: MappingRegistry):
        self._registry = registry

    def match_by_value(
        self,
        target_label: str,
        target_row: int,
        target_prior_value: Decimal,
        current_accounts: Iterable[FinancialAccount],
        prior_accounts: Iterable[FinancialAccount],
    ) -> MappingResult | None:
        """
        Takes the previous period's value from the Spread and searches
        the prior_accounts for exactly that value. If found, returns
        the corresponding current_account's value mapping to this target.
        """
        # Find which CVM accounts (in prior period) have this exact value
        matching_prior_accs = [
            acc for acc in prior_accounts if acc.value == target_prior_value
        ]

        if not matching_prior_accs:
            return None

        # Take the most likely account if multiple match (e.g. based on registry layer 2 map)
        # We can score them. For exact numeric match, confidence is ~0.85
        best_acc = matching_prior_accs[0]

        # Now find the current period's equivalent account (same account code/description)
        # to extract its current value.
        current_val = None
        current_acc = None
        for acc in current_accounts:
            # Match by description if no code, or code if available
            if (best_acc.code and acc.code == best_acc.code) or \
               (not best_acc.code and acc.description.casefold() == best_acc.description.casefold()):
                current_val = acc.value
                current_acc = acc
                break

        if current_val is None:
            return None

        return MappingResult(
            spread_row=target_row,
            label=target_label,
            source_account=current_acc,
            value=current_val,
            confidence=0.85,
            layer=2,
        )
