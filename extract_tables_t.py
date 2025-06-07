"""extract_tables_t.py
----------------------
Enhanced extraction of financial tables from corporate PDF reports.

The script keeps ``tabula`` as the preferred backend whenever available and falls back to ``pdfplumber`` or OCR (Tesseract) for scanned pages. Column headers and year references are preserved, multi line descriptions are merged and recurring footer notes are removed. Each section (BP, DRE and DFC) is exported to its own sheet in the resulting workbook.
"""

from __future__ import annotations

import logging
import re
import unicodedata
from pathlib import Path
from typing import Dict, Iterable, List, Tuple, Union


import pandas as pd
try:  # optional dependencies
    from tabula import read_pdf
    TABULA_AVAILABLE = True
except Exception:  # pragma: no cover - not installed in some environments
    TABULA_AVAILABLE = False
    read_pdf = None

try:
    import pdfplumber  # type: ignore
except Exception:  # pragma: no cover
    pdfplumber = None

try:
    from pdf2image import convert_from_path  # type: ignore
    import pytesseract  # type: ignore
except Exception:  # pragma: no cover
    convert_from_path = None
    pytesseract = None

# ---------------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Cleaning helpers (similar to the previous version)
# ---------------------------------------------------------------------------
REGEX_INVISIBLES = re.compile(r"[\u200b\u200e\u202f\xa0]")
FOOTER_PATTERN = re.compile(r"accompanying notes", re.I)

def clean_cell(cell: Union[str, float, int]) -> Union[str, float, int]:
    """Normalize basic whitespace and invisible characters."""

    if not isinstance(cell, str):
        return cell

    cell = unicodedata.normalize("NFKC", cell)
    cell = REGEX_INVISIBLES.sub("", cell).strip()
    return cell


def remove_footer_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Drop rows that contain recurring footer notes."""

    if df.empty:
        return df

    mask = df.apply(
        lambda r: r.astype(str).str.contains(FOOTER_PATTERN).any(), axis=1
    )
    return df[~mask].reset_index(drop=True)


def merge_header_rows(df: pd.DataFrame, header_rows: int = 2) -> pd.DataFrame:
    """Combine the first ``header_rows`` into a single header row."""

    if df.empty:
        return df

    parts = df.iloc[:header_rows].fillna("")
    headers = [
        " ".join(filter(None, parts[col].astype(str).str.strip().tolist()))
        for col in parts.columns
    ]
    body = df.iloc[header_rows:].reset_index(drop=True)
    body.columns = headers[: body.shape[1]]
    return body


def merge_multiline_labels(df: pd.DataFrame, label_col: int = 0) -> pd.DataFrame:
    """Merge rows where the label column is split across lines."""

    if df.empty:
        return df

    rows: List[List[Union[str, float, int]]] = []
    buffer: List[Union[str, float, int]] | None = None

    for _, row in df.iterrows():
        label = row[label_col]
        if pd.isna(label) and buffer is not None:
            txt = " ".join(
                str(row[c]) for c in df.columns if c != label_col and not pd.isna(row[c])
            )
            buffer[label_col] = f"{buffer[label_col]} {txt}".strip()
        else:
            if buffer is not None:
                rows.append(buffer)
            buffer = row.tolist()

    if buffer is not None:
        rows.append(buffer)

    return pd.DataFrame(rows, columns=df.columns)


def preprocess_table(df: pd.DataFrame) -> pd.DataFrame:
    """Apply :func:`clean_cell` to every element of ``df``."""

    for col in df.columns:
        df[col] = df[col].apply(clean_cell)
    return df


def parse_pages(pages: str) -> List[int]:
    """Expand a page specification like ``"1,3-5"`` to a list of numbers."""

    result: List[int] = []
    for part in pages.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            result.extend(range(int(a), int(b) + 1))
        else:
            result.append(int(part))
    return result




# ---------------------------------------------------------------------------
# Detection helpers
# ---------------------------------------------------------------------------

def pages_range(start: int, end: int) -> List[int]:
    return list(range(start, end + 1))


def is_scanned(pdf_path: str, pages: Iterable[int]) -> bool:
    """Return True if pages appear to contain no text (likely scanned)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for p in pages:
                if p - 1 < len(pdf.pages):
                    txt = pdf.pages[p - 1].extract_text()
                    if txt and txt.strip():
                        return False
    except Exception as exc:  # pragma: no cover
        logger.warning("Failed to inspect PDF text: %s", exc)
    return True

# ---------------------------------------------------------------------------
# Extraction routines
# ---------------------------------------------------------------------------

def extract_tables_text(pdf_path: str, pages: str) -> List[pd.DataFrame]:
    """Extract tables using Tabula or ``pdfplumber``."""

    tables: List[pd.DataFrame] = []
    if TABULA_AVAILABLE:
        try:
            tables = read_pdf(
                pdf_path,
                pages=pages,
                multiple_tables=True,
                stream=True,
                pandas_options={"header": None},
            )
        except Exception as exc:  # pragma: no cover
            logger.error("Tabula extraction failed: %s", exc)

    if not tables and pdfplumber:
        for p in parse_pages(pages):
            with pdfplumber.open(pdf_path) as pdf:
                if p - 1 >= len(pdf.pages):
                    continue
                page = pdf.pages[p - 1]
                for tb in page.extract_tables():
                    tables.append(pd.DataFrame(tb))

    return [merge_multiline_labels(merge_header_rows(remove_footer_rows(preprocess_table(t)))) for t in tables]


def extract_tables_ocr(pdf_path: str, page_numbers: Iterable[int]) -> List[pd.DataFrame]:
    """Fallback OCR extraction using ``pdf2image`` and ``pytesseract``."""

    results: List[pd.DataFrame] = []
    if convert_from_path is None or pytesseract is None:
        logger.error("OCR dependencies are missing")
        return results

    try:
        images = convert_from_path(
            pdf_path,
            first_page=min(page_numbers),
            last_page=max(page_numbers),
        )
    except Exception as exc:  # pragma: no cover
        logger.error("Failed to render pages for OCR: %s", exc)
        return results

    for img in images:
        text = pytesseract.image_to_string(img, config="--psm 6")
        rows = [re.split(r"\s{2,}", ln.strip()) for ln in text.splitlines() if ln.strip()]
        if rows:
            df = pd.DataFrame(rows)
            df = preprocess_table(df)
            df = remove_footer_rows(df)
            df = merge_header_rows(df)
            df = merge_multiline_labels(df)
            results.append(df)
    return results

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def save_sections_to_excel(pdf_path: str, sections: Dict[str, Tuple[int, int]], output: Path) -> None:
    """Extract each section and store it in ``output``."""

    with pd.ExcelWriter(output) as writer:
        for section, (start, end) in sections.items():
            page_list = pages_range(start, end)
            if is_scanned(pdf_path, page_list):
                logger.info("Section %s appears scanned. Using OCR.", section)
                tables = extract_tables_ocr(pdf_path, page_list)
            else:
                pages = f"{start}-{end}"
                logger.info("Extracting section %s pages %s", section, pages)
                tables = extract_tables_text(pdf_path, pages)

            if not tables:
                logger.warning("No tables found for section %s", section)
                continue

            combined = pd.concat(tables, ignore_index=True) if len(tables) > 1 else tables[0]
            combined.to_excel(writer, sheet_name=section, index=False)
            logger.info("Saved %s (%d rows)", section, len(combined))

# ---------------------------------------------------------------------------
# CLI entry
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # Default example using the sample PDF shipped with the repository.
    PDF_FILE = Path("vibraITR.pdf")
    SECTIONS = {
        "BP": (3, 3),
        "DRE": (4, 4),
        "DFC": (7, 7),
    }

    out_path = PDF_FILE.with_name(f"{PDF_FILE.stem}_tables.xlsx")
    try:
        save_sections_to_excel(str(PDF_FILE), SECTIONS, out_path)
        logger.info("Tables written to %s", out_path)
    except Exception as exc:  # pragma: no cover
        logger.error("Failed to extract tables: %s", exc)

