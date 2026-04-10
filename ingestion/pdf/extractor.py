"""Extracts raw Markdown content from PDFs using MarkItDown."""

import logging
from pathlib import Path

# Provide a fallback mock or real MarkItDown depending on user environment
try:
    from markitdown import MarkItDown
except ImportError:
    MarkItDown = None

logger = logging.getLogger(__name__)


class MarkdownExtractor:
    """Uses MarkItDown to convert visual PDF tables into Markdown."""

    def __init__(self):
        if not MarkItDown:
            logger.warning("MarkItDown library not found. Extractor will fail if called.")

    def extract(self, pdf_path: str | Path) -> str:
        """Converts a PDF to Markdown text preserving table structures."""
        # This is the POC phase extraction target (~40%+ accuracy).
        pdf_path = Path(pdf_path)
        
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
            
        if not MarkItDown:
            raise RuntimeError("MarkItDown is required for PDF extraction.")

        md_converter = MarkItDown()
        result = md_converter.convert(pdf_path)
        return result.text_content
