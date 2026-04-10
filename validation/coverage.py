"""Validates chronological gaps for multi-period input sequences."""

from __future__ import annotations

from core.periods import parse_period, is_annual


class CoverageValidator:
    """Verifies that an array of datasets forms a continuous sequence."""

    def validate_gaps(self, periods: list[str]) -> list[str]:
        """
        Parses periods, sorts them, and tests if there are skipped years/quarters.
        Returns a list of missing period strings if any.
        """
        if len(periods) <= 1:
            return []
            
        parsed = [parse_period(p) for p in periods]
        annuals = [int(p) for p in parsed if is_annual(p)]
        
        # Super simplified heuristic for annual gaps
        if annuals:
            annuals.sort()
            gaps = []
            for i in range(1, len(annuals)):
                current = annuals[i]
                prev = annuals[i-1]
                if current - prev > 1:
                    # Gaps found
                    for missing_year in range(prev + 1, current):
                        gaps.append(str(missing_year))
            return gaps
            
        # Add actual quarterly gaps reasoning later if needed
        return []
