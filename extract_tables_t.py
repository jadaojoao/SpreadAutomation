from __future__ import annotations

import logging
import re
import unicodedata
from pathlib import Path
from typing import Dict, Iterable, List, Tuple, Union

import pandas as pd

try:
    from tabula import read_pdf
    TABULA_AVAILABLE = True
except Exception:
    TABULA_AVAILABLE = False
    read_pdf = None

try:
    import camelot
    CAMELOT_AVAILABLE = True
except Exception:
    CAMELOT_AVAILABLE = False

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from pdf2image import convert_from_path
    import pytesseract
except Exception:
    convert_from_path = None
    pytesseract = None

# ---------------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Cleaning helpers
# ---------------------------------------------------------------------------
REGEX_INVISIBLES = re.compile(r"[\u200b\u200e\u202f\xa0]")
FOOTER_PATTERN = re.compile(r"accompanying notes", re.I)
NUM_PATTERN = re.compile(r"^-?\d[\d\.,]*$")


def clean_cell(cell: Union[str, float, int]) -> Union[str, float, int]:
    """Remove caracteres invisíveis e normaliza Unicode (NFKC)."""
    if not isinstance(cell, str):
        return cell
    cell = unicodedata.normalize("NFKC", cell)
    cell = REGEX_INVISIBLES.sub("", cell).strip()
    return cell


def remove_footer_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Exclui linhas de rodapé recorrentes."""
    if df.empty:
        return df
    mask = df.apply(
        lambda r: r.astype(str).str.contains(FOOTER_PATTERN).any(), axis=1
    )
    return df[~mask].reset_index(drop=True)


def merge_header_rows(df: pd.DataFrame, header_rows: int | None = None) -> pd.DataFrame:
    """Funde as primeiras linhas do cabeçalho em uma única linha."""
    if df.empty:
        return df
    if header_rows is None:
        header_rows = guess_header_rows(df)
    parts = df.iloc[:header_rows].fillna("")
    headers = [
        " ".join(filter(None, parts[col].astype(str).str.strip().tolist()))
        for col in parts.columns
    ]
    body = df.iloc[header_rows:].reset_index(drop=True)
    body.columns = headers[: body.shape[1]]
    return body


def merge_multiline_labels(df: pd.DataFrame, label_col: int = 0) -> pd.DataFrame:
    """Une descrições que vêm quebradas em múltiplas linhas."""
    if df.empty:
        return df

    rows: List[List[Union[str, float, int]]] = []
    buffer: List[Union[str, float, int]] | None = None

    for _, row in df.iterrows():
        label = row.iloc[label_col]
        if pd.isna(label) or not row.drop(df.columns[label_col]).dropna().any():
            txt = " ".join(str(v) for v in row if isinstance(v, str) and v)
            if buffer is None:
                buffer = [txt] + [None] * (len(df.columns) - 1)
            else:
                buffer[label_col] = f"{buffer[label_col]} {txt}".strip()
        else:
            if buffer is not None:
                rows.append(buffer)
                buffer = None
            rows.append(row.tolist())

    if buffer is not None:
        rows.append(buffer)

    return pd.DataFrame(rows, columns=df.columns)


def preprocess_table(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(clean_cell)
    df = df.dropna(how="all").reset_index(drop=True)
    df = df.dropna(axis=1, how="all")
    return df


def guess_header_rows(df: pd.DataFrame, max_rows: int = 5) -> int:
    for i, row in df.iterrows():
        if i >= max_rows:
            break
        if any(NUM_PATTERN.fullmatch(str(v)) for v in row if isinstance(v, str)):
            return i
    return max_rows


def parse_pages(pages: str) -> List[int]:
    """Expande '1,3-5' para [1,3,4,5]."""
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
    """Retorna True se as páginas parecem não conter texto (provavelmente escaneadas)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for p in pages:
                if p - 1 < len(pdf.pages):
                    txt = pdf.pages[p - 1].extract_text()
                    if txt and txt.strip():
                        return False
    except Exception as exc:
        logger.warning("Failed to inspect PDF text: %s", exc)
    return True

# ---------------------------------------------------------------------------
# Extraction routines
# ---------------------------------------------------------------------------
def extract_tables_text(pdf_path: str, pages: str) -> List[pd.DataFrame]:
    """Extrai tabelas via Tabula (se disponível) ou pdfplumber."""
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
        except Exception as exc:
            logger.error("Tabula extraction failed: %s", exc)

    def bad_result(ts: List[pd.DataFrame]) -> bool:
        return not ts or all(t.empty or t.shape[1] > 10 for t in ts)

    if bad_result(tables) and CAMELOT_AVAILABLE:
        try:
            cam = camelot.read_pdf(pdf_path, pages=pages, flavor="stream")
            tables = [tb.df for tb in cam]
        except Exception as exc:
            logger.error("Camelot extraction failed: %s", exc)

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
    """Fallback de OCR (pdf2image + pytesseract) quando não há texto extraível."""
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
    except Exception as exc:
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
    """Extrai cada seção e salva no Excel, garantindo ao menos uma aba visível."""
    wrote_anything = False

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

            if len(tables) > 1:
                cols = []
                for tb in tables:
                    if len(tb.columns) > len(cols):
                        cols = list(tb.columns)
                tables = [tb.reindex(columns=cols) for tb in tables]
                combined = pd.concat(tables, ignore_index=True)
            else:
                combined = tables[0]
            combined.to_excel(writer, sheet_name=section, index=False)
            logger.info("Saved %s (%d rows)", section, len(combined))
            wrote_anything = True

        if not wrote_anything:
            pd.DataFrame([["Nenhuma tabela encontrada no PDF."]]).to_excel(
                writer, sheet_name="Erro", index=False, header=False
            )
            wrote_anything = True  # garante visibilidade
            logger.warning("Nenhuma aba criada com dados — adicionada aba de erro.")

# ---------------------------------------------------------------------------
# CLI entry
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    PDF_FILE = Path("vibraITR.pdf")       # substitua pelo seu PDF
    SECTIONS = {
        "BP": (3, 3),
        "DRE": (4, 4),
        "DFC": (7, 7),
    }

    out_path = PDF_FILE.with_name(f"{PDF_FILE.stem}_tables.xlsx")
    try:
        save_sections_to_excel(str(PDF_FILE), SECTIONS, out_path)
        logger.info("Tables written to %s", out_path)
    except Exception as exc:
        logger.error("Failed to extract tables: %s", exc)
