"""Normalizes textual artifacts, negative signs, and units from Markdown OCR."""

import re
from decimal import Decimal


def normalize_financial_value(raw: str) -> Decimal | None:
    """
    Parses messy financial OCR data strings into a standard Decimal.
    Handles:
    - (123.456) => -123456
    - R$ 1.234,56 => 1234.56
    - 1,234.56 => 1234.56
    - Em branco => None
    """
    val_str = str(raw).strip()
    if not val_str or val_str.lower() in ("-", "na", "n/a", ""):
        return None

    # Handle negative via parentheses
    is_negative = False
    if val_str.startswith("(") and val_str.endswith(")"):
        is_negative = True
        val_str = val_str[1:-1]

    # Remove currency abbreviations and arbitrary whitespaces
    val_str = re.sub(r"[A-Za-zR$]", "", val_str).strip()

    # Determine if the comma is a decimal separator or a thousands separator
    if "," in val_str and "." in val_str:
        # Brazilian format 1.234,56
        if val_str.rfind(",") > val_str.rfind("."):
            val_str = val_str.replace(".", "").replace(",", ".")
        # US format 1,234.56
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        # Probably Brazilian decimal 1234,56
        val_str = val_str.replace(",", ".")

    # We now have hopefully clean floats like '1234.56'
    try:
        dec = Decimal(val_str)
        return -dec if is_negative else dec
    except Exception:
        # If OCR completely junked this block up, pass
        return None
