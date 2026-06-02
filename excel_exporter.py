"""
Excel export module for financial statement data.
Creates well-formatted .xlsx files with separate sheets for each statement.
"""

import os
from typing import Dict, Optional
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

from config import OUTPUT_DIR


# Style definitions
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F3864")
SUBTITLE_FONT = Font(name="Calibri", size=10, color="666666")
DATA_FONT = Font(name="Calibri", size=10)
NUMBER_FORMAT = '#,##0.00'
PERCENT_FORMAT = '0.00%'
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
HEADER_BORDER = Border(
    left=Side(style="thin", color="1F3864"),
    right=Side(style="thin", color="1F3864"),
    top=Side(style="thin", color="1F3864"),
    bottom=Side(style="thin", color="1F3864"),
)
ALTERNATE_FILL = PatternFill(start_color="F2F7FC", end_color="F2F7FC", fill_type="solid")


STATEMENT_LABELS = {
    "income_statement": "Income Statement",
    "balance_sheet": "Balance Sheet",
    "cash_flow": "Cash Flow Statement",
    "ratios": "Financial Ratios",
    "kpi_metrics": "KPI Metrics",
}

# Short display names for sheet tabs (max 31 chars for Excel)
SHEET_NAMES = {
    "income_statement": "Income Statement",
    "balance_sheet": "Balance Sheet",
    "cash_flow": "Cash Flow",
    "ratios": "Financial Ratios",
    "kpi_metrics": "KPI Metrics",
}


def ensure_output_dir():
    """Create the output directory if it doesn't exist."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def _auto_column_width(ws, min_width=12, max_width=36):
    """Auto-fit column widths based on content."""
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        adjusted = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = adjusted


def _apply_header_style(ws, row_idx: int, col_count: int):
    """Apply styling to header row."""
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = HEADER_BORDER


def _apply_data_style(ws, start_row: int, end_row: int, col_count: int):
    """Apply styling to data rows with alternating colors."""
    for row in range(start_row, end_row + 1):
        for col in range(1, col_count + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER

            if col == 1:
                # Line item name - left aligned, bold
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = Font(name="Calibri", size=10, bold=True)
            else:
                # Numeric data - right aligned
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = NUMBER_FORMAT

            # Alternating row colors
            if (row - start_row) % 2 == 0:
                cell.fill = ALTERNATE_FILL


def _write_dataframe_to_sheet(
    ws, df: pd.DataFrame, sheet_title: str, start_row: int = 1
) -> int:
    """
    Write a DataFrame to a worksheet with formatting.
    Returns the last row number used.
    """
    # Write title
    title_cell = ws.cell(row=start_row, column=1, value=sheet_title)
    title_cell.font = TITLE_FONT
    ws.merge_cells(
        start_row=start_row, start_column=1,
        end_row=start_row, end_column=len(df.columns) + 1
    )

    # Header row (column names = periods)
    header_row = start_row + 1
    # Write "Line Item" for the first column
    ws.cell(row=header_row, column=1, value="Line Item")
    for col_idx, col_name in enumerate(df.columns, start=2):
        ws.cell(row=header_row, column=col_idx, value=str(col_name))

    col_count = len(df.columns) + 1
    _apply_header_style(ws, header_row, col_count)

    # Data rows
    data_start = header_row + 1
    for row_idx, (index_val, row_data) in enumerate(df.iterrows()):
        current_row = data_start + row_idx
        ws.cell(row=current_row, column=1, value=str(index_val))
        for col_idx, value in enumerate(row_data, start=2):
            if pd.notna(value):
                ws.cell(row=current_row, column=col_idx, value=value)

    data_end = data_start + len(df) - 1
    _apply_data_style(ws, data_start, data_end, col_count)

    # Auto-fit columns
    _auto_column_width(ws)
    # First column wider for line item names
    ws.column_dimensions["A"].width = 38

    return data_end


def export_single_stock(
    symbol: str,
    statements: Dict[str, Optional[pd.DataFrame]],
    output_dir: str = OUTPUT_DIR,
) -> Optional[str]:
    """
    Export financial statements for a single stock to an Excel file.
    Each statement type gets its own sheet.

    Args:
        symbol: Stock ticker symbol
        statements: Dict of statement_type -> DataFrame
        output_dir: Directory to save the file

    Returns:
        Path to the created file, or None if no data
    """
    ensure_output_dir()

    # Filter out None values
    valid_statements = {
        k: v for k, v in statements.items() if v is not None and not v.empty
    }

    if not valid_statements:
        print(f"  ⚠️  No data to export for {symbol}")
        return None

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    first = True
    for stmt_type, df in valid_statements.items():
        sheet_name = SHEET_NAMES.get(stmt_type, stmt_type)
        ws = wb.create_sheet(title=sheet_name)
        label = STATEMENT_LABELS.get(stmt_type, stmt_type)
        _write_dataframe_to_sheet(ws, df, f"{symbol} - {label}")

    filepath = os.path.join(output_dir, f"{symbol}_financials.xlsx")
    wb.save(filepath)
    print(f"  💾 Saved: {filepath}")
    return filepath


def export_all_stocks(
    all_data: Dict[str, Dict[str, Optional[pd.DataFrame]]],
    output_dir: str = OUTPUT_DIR,
) -> list:
    """
    Export financial statements for multiple stocks to separate Excel files.

    Args:
        all_data: Nested dict: symbol -> statement_type -> DataFrame
        output_dir: Directory to save files

    Returns:
        List of created file paths
    """
    ensure_output_dir()
    created_files = []

    for symbol, statements in all_data.items():
        filepath = export_single_stock(symbol, statements, output_dir)
        if filepath:
            created_files.append(filepath)

    return created_files


def create_summary_file(
    all_data: Dict[str, Dict[str, Optional[pd.DataFrame]]],
    output_dir: str = OUTPUT_DIR,
) -> Optional[str]:
    """
    Create a summary Excel file with key metrics from all stocks.

    Args:
        all_data: Nested dict: symbol -> statement_type -> DataFrame
        output_dir: Directory to save the file

    Returns:
        Path to the summary file
    """
    if not all_data:
        return None

    ensure_output_dir()

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    # Write header
    ws.cell(row=1, column=1, value="Stock Financial Data Summary").font = Font(
        name="Calibri", size=14, bold=True, color="1F3864"
    )
    ws.merge_cells("A1:F1")

    ws.cell(row=3, column=1, value="Symbol").font = Font(bold=True)
    ws.cell(row=3, column=2, value="Income Statement Rows").font = Font(bold=True)
    ws.cell(row=3, column=3, value="Balance Sheet Rows").font = Font(bold=True)
    ws.cell(row=3, column=4, value="Cash Flow Rows").font = Font(bold=True)
    ws.cell(row=3, column=5, value="Ratios Rows").font = Font(bold=True)
    ws.cell(row=3, column=6, value="KPI Rows").font = Font(bold=True)

    row = 4
    for symbol, statements in all_data.items():
        ws.cell(row=row, column=1, value=symbol).font = Font(bold=True)

        for col_offset, stmt_type in enumerate(
            ["income_statement", "balance_sheet", "cash_flow", "ratios", "kpi_metrics"], start=2
        ):
            df = statements.get(stmt_type)
            if df is not None and not df.empty:
                ws.cell(row=row, column=col_offset, value=f"{df.shape[0]} rows × {df.shape[1]} periods")
            else:
                ws.cell(row=row, column=col_offset, value="No data")
        row += 1

    # Column widths
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 28

    filepath = os.path.join(output_dir, "_summary.xlsx")
    wb.save(filepath)
    print(f"  📋 Summary saved: {filepath}")
    return filepath
