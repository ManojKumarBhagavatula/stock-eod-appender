"""
excel_utils.py — Workbook load, validate, and save helpers
"""

import io

import openpyxl

from config import SYMBOL_COL, EXCHANGE_COL


def load_and_validate(file_bytes: bytes):
    """
    Load workbook from bytes and validate required columns.

    Returns:
        (wb, ws, headers, sym_idx, exch_idx)
    """
    wb      = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws      = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    if SYMBOL_COL not in headers or EXCHANGE_COL not in headers:
        raise ValueError(
            f"Required columns not found. Need '{SYMBOL_COL}' and '{EXCHANGE_COL}'."
        )

    sym_idx  = headers.index(SYMBOL_COL)
    exch_idx = headers.index(EXCHANGE_COL)
    return wb, ws, headers, sym_idx, exch_idx


def wb_to_bytes(wb) -> bytes:
    """Serialize an openpyxl workbook to raw bytes."""
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
