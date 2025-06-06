"""extract_tables_t.py
----------------------
Improved table extraction from corporate PDF reports.

Features:
- Detects whether pages are text-based or scanned images.
- Uses Tabula on text-based pages and Tesseract OCR as a fallback.
- Cleans and splits merged cells using regex heuristics.
- Saves each extracted table into an Excel workbook with a sheet per
  report section (BP, DRE, DFC).
- Provides logging and error handling for easier troubleshooting.
"""

from __future__ import annotations

import logging
import re
import unicodedata
from pathlib import Path
from typing import Dict, Iterable, List, Tuple, Union

import pandas as pd
from tabula import read_pdf
import pdfplumber
from pdf2image import convert_from_path
import pytesseract

# ---------------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Cleaning helpers (similar to the previous version)
# ---------------------------------------------------------------------------
REGEX_INVISIBLES = re.compile(r"[\u200b\u200e\u202f\xa0]")


def clean_and_split_cell(cell: Union[str, float, int]) -> Union[str, float, int]:
    """Clean and split concatenated or malformed cell content."""
    if not isinstance(cell, str):
        return cell

    cell = unicodedata.normalize("NFKC", cell)
    cell = REGEX_INVISIBLES.sub("", cell).strip()

    cell = re.sub(r"(\d+)\s*-\s*$", r"\1;-", cell)
    cell = re.sub(r"(-?\d+)\s+-\b", r"\1;-", cell)
    cell = re.sub(r"-\s+(-?\d+)", r"-;\1", cell)
    cell = re.sub(
        r"(?<![\d/])(\d{1,3}(?:\.\d{3})*|\d+)\s+(\d{1,3}(?:\.\d{3})*|\d+)(?![\d/])",
        r"\1;\2",
        cell,
    )
    cell = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", cell)
    cell = re.sub(r"([a-z])([A-Z])", r"\1 \2", cell)
    cell = re.sub(r"(?<=[a-zA-Z])(?=\d)", " ", cell)
    cell = re.sub(r"(?<=\d)(?=\d{3}\b)", " ", cell)
    cell = re.sub(r"(\d{2})\s*/\s*(\d{2})\s*/\s*(\d{4})", r"\1/\2/\3", cell)
    cell = re.sub(r"(\d)\s+(\d{3})", r"\1\2", cell)
    cell = re.sub(r"(\d{4})\s+(\d{4})", r"\1;\2", cell)
    cell = re.sub(r"(\d{3})(\d\.\d{3})", r"\1;\2", cell)
    cell = re.sub(r"(\d{1,3}(?:\.\d{3})+)(\d{3})\b", r"\1;\2", cell)
    cell = re.sub(r"(\d+\.\d{3})(\d{3}\b)", r"\1;\2", cell)
    cell = re.sub(r"(\d+\.\d)\s+(\d{3}\.\d{3})", r"\1;\2", cell)
    cell = re.sub(r"(\d{1,3}(?:\.\d{3})+)(\d{1,3}(?:\.\d{3})+)", r"\1;\2", cell)
    cell = re.sub(r"\((\d+\.\d{3})\)\s+\((\d+\.\d{3})\)", r"(\1);(\2)", cell)
    cell = re.sub(r"\((\d+)\)\s+\((\d+)\)", r"(\1);(\2)", cell)
    cell = re.sub(r"\((\d+)\)\s+(\d+)", r"(\1);\2", cell)
    cell = re.sub(r"(\d+)\s+\((\d+)\)", r"\1;(\2)", cell)
    cell = re.sub(r"(\d+)\s+\((\d+\.\d{3})\)", r"\1;(\2)", cell)
    cell = re.sub(r"\((\d+\.\d{3})\)\s+(\d+)", r"(\1);\2", cell)
    cell = re.sub(r"\((\d+)\)\((\d+\.\d{3})\)", r"\1;(\2)", cell)
    cell = re.sub(r"\((\d+\.\d{3})\)\((\d+\.\d{3})\)", r"(\1);(\2)", cell)
    cell = re.sub(r"\((\d+\.\d{3})\)\((\d+)\)", r"(\1);(\2)", cell)
    cell = re.sub(r"\((\d+)\)\((\d+)\)", r"\1;(\2)", cell)
    cell = re.sub(r"(\d+)\((\d+\.\d{3})\)", r"\1;(\2)", cell)
    cell = re.sub(r"\((\d+(?:\.\d{3})?)\)", r"-\1", cell)
    cell = re.sub(r"([a-zA-Z]+)\s*(\d{1,3})\b", r"\1;\2", cell)
    cell = re.sub(r"(-?\d+)\s+-\b", r"\1;-", cell)
    return cell.strip()


def preprocess_table(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(clean_and_split_cell)
    return df


def manually_split_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: re.split(r"(?<=\d{3})(?=\d{3})", x)
                if isinstance(x, str) and re.search(r"\d{3}\d{3}", x)
                else x
            )
            df[col] = df[col].apply(lambda x: ";".join(x) if isinstance(x, list) else x)
    return df


def split_semicolon_values(df: pd.DataFrame) -> pd.DataFrame:
    new_cols: List[pd.DataFrame] = []
    for col in df.columns:
        if df[col].dtype == "object" and df[col].str.contains(";").any():
            split_data = df[col].str.split(";", expand=True)
            split_data.columns = [f"{col}_{i+1}" for i in range(split_data.shape[1])]
            new_cols.append(split_data)
        else:
            new_cols.append(df[[col]])
    return pd.concat(new_cols, axis=1)

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
    try:
        tables = read_pdf(
            pdf_path,
            pages=pages,
            multiple_tables=True,
            stream=True,
            pandas_options={"header": None},
        )
        return [preprocess_table(manually_split_columns(split_semicolon_values(t))) for t in tables]
    except Exception as exc:
        logger.error("Tabula extraction failed: %s", exc)
        return []


def extract_tables_ocr(pdf_path: str, page_numbers: Iterable[int]) -> List[pd.DataFrame]:
    results: List[pd.DataFrame] = []
    try:
        images = convert_from_path(pdf_path, first_page=min(page_numbers), last_page=max(page_numbers))
    except Exception as exc:
        logger.error("Failed to render pages for OCR: %s", exc)
        return results

    for img in images:
        text = pytesseract.image_to_string(img, config="--psm 6")
        rows = [re.split(r"\s{2,}", ln.strip()) for ln in text.splitlines() if ln.strip()]
        if rows:
            df = pd.DataFrame(rows)
            results.append(preprocess_table(manually_split_columns(split_semicolon_values(df))))
    return results

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def save_sections_to_excel(pdf_path: str, sections: Dict[str, Tuple[int, int]], output: Path) -> None:
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

            for idx, tbl in enumerate(tables, 1):
                sheet = f"{section}_{idx}"
                tbl.to_excel(writer, sheet_name=sheet, index=False, header=False)
                logger.info("Saved %s (%d rows)", sheet, len(tbl))

# ---------------------------------------------------------------------------
# CLI entry
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # Example usage: update these paths/ranges for your PDFs.
    PDF_FILE = "matrixITR.pdf"  # path to your PDF
    SECTIONS = {
        "BP": (8, 10),
        # Add other sections with their page ranges, e.g.:
        # "DRE": (11, 13),
        # "DFC": (14, 15),
    }
    out_path = Path(PDF_FILE).with_suffix("_tables.xlsx")
    try:
        save_sections_to_excel(PDF_FILE, SECTIONS, out_path)
        logger.info("Tables written to %s", out_path)
    except Exception as exc:  # pragma: no cover
        logger.error("Failed to extract tables: %s", exc)

