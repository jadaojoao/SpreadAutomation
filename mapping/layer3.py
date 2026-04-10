"""Layer 3 Matcher: Semantic fuzzy text analysis for PDF integration."""

from __future__ import annotations

from typing import Iterable
from decimal import Decimal

# Rapidfuzz import fallback
try:
    from rapidfuzz import process, fuzz
except ImportError:
    process = None
    fuzz = None

from core.models import FinancialAccount, MappingResult
from mapping.registry import MappingRegistry


class Layer3Matcher:
    """
    Performs fuzzy matching looking for description similarities in OCR outputs.
    Since confidence varies by text similarity, we scale it.
    """

    def __init__(self, registry: MappingRegistry):
        self._registry = registry
        if not process:
            # Not raising hard so tests run if rapidfuzz absent
            pass

    def match_fuzzy(
        self,
        target_label: str,
        target_row: int,
        target_prior_value: Decimal | None,
        current_accounts: Iterable[FinancialAccount]
    ) -> MappingResult | None:
        """
        Uses rapidfuzz to try answering "Does this target label match any OCR description?"
        Uses synonyms inside MappingRegistry to augment searches.
        """
        if not process:
            return None

        # Build list of words to match against for current spread label
        synonyms = [target_label] + self._registry.synonyms(target_label)
        
        # Flatten current_account descriptions
        descriptions = [acc.description for acc in current_accounts]
        if not descriptions:
            return None

        best_score = 0.0
        best_acc = None

        for synonym in synonyms:
            # Extract top matching description
            match = process.extractOne(
                synonym, descriptions,
                scorer=fuzz.token_sort_ratio
            )
            # rapidfuzz returns (matched_str, score, index)
            if match and match[1] > best_score:
                best_score = match[1]
                best_acc = list(current_accounts)[match[2]]

        # We set a threshold of 75/100 to accept the fuzzy match
        if best_score > 75.0 and best_acc:
            return MappingResult(
                spread_row=target_row,
                label=target_label,
                source_account=best_acc,
                value=best_acc.value,
                confidence=round(best_score / 100.0, 2),  # OCR confidence ratio
                layer=3,
            )

        return None
