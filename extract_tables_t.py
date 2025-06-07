"""extract_tables_t.py
----------------------
Utility to extract financial tables from PDF reports.

Tabula is used on text based pages while pdfplumber only inspects pages
to detect whether they contain selectable text.  If no text is found,
the script falls back to OCR via ``pdf2image`` and ``pytesseract``.  Each
report section (Balance Sheet, Income Statement and Cash Flow Statement)
is saved to a dedicated sheet in the resulting Excel file.
Each report section (Balance Sheet, Income Statement and Cash Flow
Statement) is saved to a dedicated sheet in the resulting Excel file.

The implementation avoids the very aggressive splitting rules of the
previous version and tries to preserve column headers and multi-line row
labels.  Simple cleaning of invisible characters and footnotes is
applied so the final DataFrame mirrors the PDF layout as closely as
possible.
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List, Tuple
import logging
import re
import unicodedata

import pandas as pd
import pdfplumber
from tabula import read_pdf
from pdf2image import convert_from_path
import pytesseract

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
REGEX_INVISIBLES = re.compile(r"[\u200b\u200e\u202f\xa0]")
FOOTNOTE_RE = re.compile(
    r"accompanying notes are an integral part", re.IGNORECASE
)


def normalise_cell(value: object) -> object:
    """Return cleaned cell value."""
    if not isinstance(value, str):
        return value
    value = unicodedata.normalize("NFKC", value)
    value = REGEX_INVISIBLES.sub("", value)
    return value.strip()


def merge_multiline_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Join rows where only the first column contains text."""

    rows: List[List[object]] = []
    buffer: List[object] | None = None

    for _, r in df.iterrows():
        first = (r.iloc[0] or "").strip() if isinstance(r.iloc[0], str) else r.iloc[0]
        rest = r.iloc[1:]
        if buffer is not None and (pd.isna(first) or rest.isna().all()):
            # continuation of previous label
            if isinstance(first, str) and first:
                buffer[0] = f"{buffer[0]} {first}".strip()
            continue

        if buffer is not None:
            rows.append(buffer)
        buffer = r.tolist()

    if buffer is not None:
        rows.append(buffer)

    return pd.DataFrame(rows, columns=df.columns)


def clean_table(df: pd.DataFrame) -> pd.DataFrame:
    """Basic clean-up of extracted tables."""

    df = df.applymap(normalise_cell)
    df.dropna(axis=0, how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)

    if not df.empty:
        df = df[~df.iloc[:, 0].astype(str).str.match(FOOTNOTE_RE, na=False)]

    df.reset_index(drop=True, inplace=True)
    df = merge_multiline_rows(df)
    df.reset_index(drop=True, inplace=True)
    return df


# ---------------------------------------------------------------------------
# Detection helpers
# ---------------------------------------------------------------------------

def is_scanned(pdf_path: str, pages: Iterable[int]) -> bool:
    """Return ``True`` if all selected pages contain no text."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for p in pages:
                if p - 1 < len(pdf.pages):
                    if pdf.pages[p - 1].extract_text() or "":
                        return False
    except Exception as exc:  # pragma: no cover - diagnostics only
        logger.warning("Failed to inspect PDF: %s", exc)
    return True


# ---------------------------------------------------------------------------
# Extraction routines
# ---------------------------------------------------------------------------

def extract_tables_text(pdf_path: str, page_numbers: Iterable[int]) -> List[pd.DataFrame]:
    """Extract tables from ``page_numbers`` using Tabula."""

    page_spec = ",".join(str(p) for p in page_numbers)
    try:
        tables = read_pdf(
            pdf_path,
            pages=page_spec,
            multiple_tables=True,
            stream=True,
            guess=True,
            pandas_options={"header": None},
        )
    except Exception as exc:  # pragma: no cover - extraction diagnostics
        logger.error("Tabula extraction failed: %s", exc)
        return []

    results: List[pd.DataFrame] = []
    for df in tables:
        df = clean_table(df)
        if not df.empty:
            results.append(df)
    return results


def extract_tables_ocr(pdf_path: str, pages: Iterable[int]) -> List[pd.DataFrame]:
    """Fallback OCR extraction for scanned pages."""
    images = convert_from_path(pdf_path, first_page=min(pages), last_page=max(pages))
    tables: List[pd.DataFrame] = []

    for img in images:
        text = pytesseract.image_to_string(img, config="--psm 6")
        rows = [re.split(r"\s{2,}", ln.strip()) for ln in text.splitlines() if ln.strip()]
        if rows:
            df = pd.DataFrame(rows)
            df = clean_table(df)
            if not df.empty:
                tables.append(df)
    return tables


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def pages_range(start: int, end: int) -> List[int]:
    return list(range(start, end + 1))


def save_sections_to_excel(
    pdf_path: str, sections: Dict[str, Tuple[int, int]], output: Path
) -> None:
    """Extract the configured sections and write them to ``output``."""

    with pd.ExcelWriter(output) as writer:
        for name, (start, end) in sections.items():
            pages = pages_range(start, end)
            if is_scanned(pdf_path, pages):
                logger.info("Section %s appears scanned. Using OCR.", name)
                tables = extract_tables_ocr(pdf_path, pages)
            else:
                logger.info("Extracting section %s pages %s", name, pages)
                tables = extract_tables_text(pdf_path, pages)

            if not tables:
                logger.warning("No tables found for section %s", name)
                continue

            # One sheet per section
            combined = pd.concat(tables, ignore_index=True)
            combined.to_excel(writer, sheet_name=name, index=False, header=False)
            logger.info("Saved %s (%d rows)", name, len(combined))


# ---------------------------------------------------------------------------
# CLI entry
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    PDF_FILE = "vibraITR.pdf"
    SECTIONS = {
        "BP": (3, 3),
        "DRE": (4, 4),
        "DFC": (7, 7),
    }
    pdf_path = Path(PDF_FILE)
    out_path = pdf_path.with_name(f"{pdf_path.stem}_tables.xlsx")

    try:
        save_sections_to_excel(str(pdf_path), SECTIONS, out_path)
        logger.info("Tables written to %s", out_path)
    except Exception as exc:  # pragma: no cover - entry diagnostics
        logger.error("Failed to extract tables: %s", exc)

