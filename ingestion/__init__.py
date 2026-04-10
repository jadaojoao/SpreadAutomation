"""
Ingestion adapters for retrieving raw data into canonical representations.
"""

from .base import IngestionAdapter
from .excel_adapter import CVMExcelAdapter
from .csv_adapter import CVMCSVAdapter
from .pdf_adapter import PDFAdapter

__all__ = [
    "IngestionAdapter",
    "CVMExcelAdapter",
    "CVMCSVAdapter",
    "PDFAdapter",
]
