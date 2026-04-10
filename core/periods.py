"""Pure helpers for normalizing and classifying business-facing periods."""

from __future__ import annotations

import re

from core.exceptions import PeriodParseError

_ANNUAL_RE = re.compile(r"^(?P<year>\d{4})$")
_TRIMESTER_RE = re.compile(r"^(?P<quarter>[1-4])T(?P<year>\d{2}|\d{4})$")
_YEAR_QUARTER_RE = re.compile(r"^(?P<year>\d{4})-?Q(?P<quarter>[1-4])$")
_QUARTER_YEAR_RE = re.compile(r"^Q(?P<quarter>[1-4])-?(?P<year>\d{4})$")
_COMPACT_DATE_RE = re.compile(r"^(?P<year>\d{4})(?P<month>\d{2})(?P<day>\d{2})$")
_SEPARATED_DATE_RE = re.compile(
    r"^(?P<year>\d{4})[-/](?P<month>\d{2})[-/](?P<day>\d{2})$"
)


def parse_period(raw: str) -> str:
    """Normalize supported raw period strings to YYYY or nTYY."""

    if not isinstance(raw, str):
        raise PeriodParseError("Period must be provided as a string.")

    period = re.sub(r"\s+", "", raw.strip().upper())
    if not period:
        raise PeriodParseError("Period cannot be empty.")

    annual_match = _ANNUAL_RE.fullmatch(period)
    if annual_match:
        return annual_match.group("year")

    trimester_match = _TRIMESTER_RE.fullmatch(period)
    if trimester_match:
        quarter = trimester_match.group("quarter")
        year = _format_short_year(trimester_match.group("year"))
        return f"{quarter}T{year}"

    year_quarter_match = _YEAR_QUARTER_RE.fullmatch(period)
    if year_quarter_match:
        quarter = year_quarter_match.group("quarter")
        year = _format_short_year(year_quarter_match.group("year"))
        return f"{quarter}T{year}"

    quarter_year_match = _QUARTER_YEAR_RE.fullmatch(period)
    if quarter_year_match:
        quarter = quarter_year_match.group("quarter")
        year = _format_short_year(quarter_year_match.group("year"))
        return f"{quarter}T{year}"

    for matcher in (_COMPACT_DATE_RE, _SEPARATED_DATE_RE):
        date_match = matcher.fullmatch(period)
        if date_match:
            year = int(date_match.group("year"))
            month = int(date_match.group("month"))
            day = int(date_match.group("day"))
            return _normalize_date_period(year=year, month=month, day=day)

    raise PeriodParseError(
        "Unsupported period format. Use YYYY, nTYY, YYYYQn, or a supported date."
    )


def periods_for_year(year: int) -> list[str]:
    """Return the four business-facing quarterly periods for a given year."""

    if not isinstance(year, int):
        raise PeriodParseError("Year must be an integer.")
    if year < 1000 or year > 9999:
        raise PeriodParseError("Year must be between 1000 and 9999.")

    short_year = f"{year % 100:02d}"
    return [f"{quarter}T{short_year}" for quarter in range(1, 5)]


def is_annual(period: str) -> bool:
    """Return True when the normalized period is annual."""

    normalized = parse_period(period)
    return bool(_ANNUAL_RE.fullmatch(normalized))


def _format_short_year(year: str) -> str:
    if len(year) == 2:
        return year
    return year[-2:]


def _normalize_date_period(*, year: int, month: int, day: int) -> str:
    quarter_ends = {
        (3, 31): f"1T{year % 100:02d}",
        (6, 30): f"2T{year % 100:02d}",
        (9, 30): f"3T{year % 100:02d}",
    }

    if (month, day) in quarter_ends:
        return quarter_ends[(month, day)]

    # Annual filings are user-facing as YYYY, even when a raw reference date is used.
    if (month, day) in {(1, 1), (12, 31)}:
        return str(year)

    raise PeriodParseError(
        "Date-based periods must be quarter-end dates or annual reference dates."
    )


__all__ = ["parse_period", "periods_for_year", "is_annual"]
