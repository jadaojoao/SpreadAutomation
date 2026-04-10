"""
Spread reader and writer using openpyxl.
Provides API for extracting the target schema and dumping results visually.
"""

from .reader import SpreadReader
from .writer import SpreadWriter
from .highlights import Highlights
from .template import TemplateManager

__all__ = [
    "SpreadReader",
    "SpreadWriter",
    "Highlights",
    "TemplateManager",
]
