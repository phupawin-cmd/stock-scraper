"""
Core scraper module for extracting financial statements from StockAnalysis.com.
"""

import time
import re
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from typing import Dict, List, Optional, Tuple
import pandas as pd

from config import BASE_URL, STATEMENT_TYPES, HEADERS, REQUEST_DELAY


def build_url(symbol: str, statement_type: str) -> str:
    """
    Build the full URL for a given stock symbol and statement type.

    Args:
        symbol: Stock ticker symbol (e.g., 'AAPL')
        statement_type: One of 'income_statement', 'balance_sheet', 'cash_flow'

    Returns:
        Full URL string
    """
    path = STATEMENT_TYPES.get(statement_type, "financials")
    return f"{BASE_URL}/{symbol.lower()}/{path}/"


def fetch_page(url: str) -> Optional[str]:
    """
    Fetch the HTML content of a page with error handling.

    Args:
        url: The URL to fetch

    Returns:
        HTML content as string, or None if failed
    """
    try:
        response = requests.get(url, headers=HEADERS, timeout=30)
        response.raise_for_status()
        return response.text
    except requests.RequestException as e:
        print(f"  ❌ Error fetching {url}: {e}")
        return None


def parse_financial_table(html: str) -> Optional[pd.DataFrame]:
    """
    Parse the financial data table from StockAnalysis.com HTML.

    The site uses div-based tables with role attributes.
    We look for the main financial data table.

    Args:
        html: Raw HTML content

    Returns:
        DataFrame with periods as columns and line items as rows, or None
    """
    soup = BeautifulSoup(html, "lxml")

    # Try to find the table element
    # StockAnalysis uses standard HTML tables or div-role tables
    table = soup.find("table", class_=re.compile("financial", re.I))

    if table is None:
        # Try finding any table with financial data
        table = soup.find("table")

    if table is None:
        # Try div-based table (role="table")
        table = soup.find("div", {"role": "table"})
        if table is not None:
            return _parse_div_table(table)

    if table is None:
        print("    ⚠️  Could not find financial table on page")
        return None

    return _parse_html_table(table, soup)


def _extract_period_dates(cell_values: list) -> list:
    """
    Extract actual date strings from Period Ending cells.

    StockAnalysis.com cells have concatenated text like:
    "Mar '26Mar 31, 2026" → extract "Mar 31, 2026"
    "Sep '25Sep 30, 2025" → extract "Sep 30, 2025"

    Non-date cells like "TTM" or "Current" are kept as-is.

    Args:
        cell_values: List of raw cell text values from the Period Ending row

    Returns:
        List of date strings (or original label for non-date columns)
    """
    # Pattern: Month Day, Year at the end of a string
    # e.g. "Mar 31, 2026" or "September 30, 2025"
    date_pattern = re.compile(
        r'([A-Z][a-z]{2,8}\s+\d{1,2},\s+\d{4})$'
    )

    dates = []
    for val in cell_values:
        val_str = str(val) if val is not None else ""
        match = date_pattern.search(val_str)
        if match:
            dates.append(match.group(1))
        else:
            # Keep original label (e.g., "TTM", "Current")
            dates.append(val_str)
    return dates


def _parse_html_table(table, soup: BeautifulSoup) -> pd.DataFrame:
    """
    Parse a standard HTML <table> element into a DataFrame.

    Handles StockAnalysis.com's concatenated date format in Period Ending cells,
    extracting proper dates to use as column headers.
    """
    rows = table.find_all("tr")

    # Extract headers (column names) from first row
    headers = []
    header_row = rows[0] if rows else None
    if header_row:
        for th in header_row.find_all(["th", "td"]):
            headers.append(th.get_text(strip=True))

    # Extract data rows
    data = []
    period_date_headers = None  # Will hold extracted dates if Period Ending row found

    for row in rows[1:]:
        cells = row.find_all(["td", "th"])
        if not cells:
            continue
        row_data = [cell.get_text(strip=True) for cell in cells]
        # Filter out empty or separator rows
        if any(row_data):
            # Check if this is the Period Ending row
            if row_data and row_data[0].lower().startswith("period ending"):
                # Extract actual dates from concatenated values
                period_date_headers = _extract_period_dates(row_data[1:])
                # Prepend the label for the first column
                period_date_headers.insert(0, row_data[0])
            data.append(row_data)

    if not data:
        return None

    # Create DataFrame
    df = pd.DataFrame(data)

    # Determine column headers: use extracted dates > original headers > default
    if period_date_headers is not None and len(period_date_headers) == df.shape[1]:
        df.columns = period_date_headers
    elif len(headers) == df.shape[1]:
        df.columns = headers

    # Set first column as index (line item names)
    if df.shape[1] > 1:
        df = df.set_index(df.columns[0])
        df.index.name = "Line Item"

    return df


def _parse_div_table(table_div) -> Optional[pd.DataFrame]:
    """
    Parse a div-based table (role="table") into a DataFrame.
    StockAnalysis.com sometimes uses this format.
    """
    # Find row groups
    headers = []
    header_divs = table_div.find_all("div", {"role": "columnheader"})
    for hd in header_divs:
        headers.append(hd.get_text(strip=True))

    rows = table_div.find_all("div", {"role": "row"})
    data = []
    for row in rows:
        cells = row.find_all("div", {"role": "cell"})
        if not cells:
            # Try gridcell
            cells = row.find_all("div", {"role": "gridcell"})
        if cells:
            row_data = [cell.get_text(strip=True) for cell in cells]
            if any(row_data):
                data.append(row_data)

    if not data:
        return None

    df = pd.DataFrame(data)
    if len(headers) == df.shape[1]:
        df.columns = headers

    if df.shape[1] > 1:
        df = df.set_index(df.columns[0])
        df.index.name = "Line Item"

    return df


def clean_numeric(value: str) -> Optional[float]:
    """
    Clean a financial value string and convert to float.

    Handles formats like:
    - '1,234.56'
    - '(1,234.56)' → negative
    - '1.23B' → billions
    - '1.23M' → millions
    - '--' → None
    - '-' → None

    Args:
        value: Raw string value from the table

    Returns:
        Float value or None if not parseable
    """
    if not value or value in ("-", "--", "N/A", "NA", ""):
        return None

    value = value.strip()

    # Check for negative (parentheses)
    is_negative = False
    if value.startswith("(") and value.endswith(")"):
        is_negative = True
        value = value[1:-1]

    # Check for minus sign
    if value.startswith("-"):
        is_negative = True
        value = value[1:]

    # Handle billion/million/thousand suffixes
    multiplier = 1.0
    if value.upper().endswith("B"):
        multiplier = 1_000_000_000
        value = value[:-1]
    elif value.upper().endswith("M"):
        multiplier = 1_000_000
        value = value[:-1]
    elif value.upper().endswith("K"):
        multiplier = 1_000
        value = value[:-1]

    # Remove commas and other non-numeric chars (except decimal point and %)
    value = re.sub(r"[^\d.\-%]", "", value)

    # Handle percentage
    is_percent = value.endswith("%")
    if is_percent:
        value = value[:-1]
        try:
            num = float(value) / 100.0
            return -num if is_negative else num
        except ValueError:
            return None

    try:
        num = float(value) * multiplier
        return -num if is_negative else num
    except ValueError:
        return None


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean all numeric values in a financial DataFrame.
    Also converts Period Ending row values to proper dates.

    Args:
        df: Raw DataFrame with string values

    Returns:
        Cleaned DataFrame with numeric values where possible
    """
    # Check if Period Ending row exists
    period_ending_mask = df.index.str.lower().str.startswith("period ending")

    if period_ending_mask.any():
        pe_index = df.index[period_ending_mask]
        # Separate Period Ending row
        df_pe = df.loc[pe_index].copy()
        df_data = df.drop(pe_index)
    else:
        df_pe = None
        df_data = df.copy()

    # Clean numeric data
    cleaned_data = pd.DataFrame(index=df_data.index)
    for col in df_data.columns:
        cleaned_data[col] = df_data[col].apply(
            lambda x: clean_numeric(str(x)) if pd.notna(x) else None
        )
        # Ensure numeric dtype
        cleaned_data[col] = pd.to_numeric(cleaned_data[col], errors='coerce')

    # Clean Period Ending dates
    if df_pe is not None:
        cleaned_pe = pd.DataFrame(index=df_pe.index)
        for col in df_pe.columns:
            cleaned_pe[col] = df_pe[col].apply(
                lambda x: _try_parse_date(str(x))
            )
        # Combine back: PE row first, then data
        cleaned = pd.concat([cleaned_pe, cleaned_data])
    else:
        cleaned = cleaned_data

    cleaned.index.name = df.index.name
    return cleaned


def _try_parse_date(value: str) -> Optional[str]:
    """
    Try to parse a value as a date. Handles StockAnalysis.com concatenated
    format like "Mar '26Mar 31, 2026" → "2026-03-31".

    Returns formatted date string "YYYY-MM-DD" or original value if unparseable.
    """
    if not value or value in ("-", "--", "N/A", "NA", "", "None"):
        return None

    # First, try to extract a date from concatenated text
    # Pattern: Month Day, Year at the end (e.g., "Mar 31, 2026")
    date_match = re.search(r'([A-Z][a-z]{2,8}\s+\d{1,2},\s+\d{4})', str(value))
    if date_match:
        value = date_match.group(1)

    # Try common date formats
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(str(value), fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue

    # If it looks like a number (already parsed incorrectly), keep as-is
    try:
        float(str(value))
        return str(value)
    except ValueError:
        return str(value)


def scrape_statement(symbol: str, statement_type: str) -> Optional[pd.DataFrame]:
    """
    Scrape a single financial statement for a given stock.

    Args:
        symbol: Stock ticker symbol
        statement_type: Type of statement to scrape

    Returns:
        Cleaned DataFrame with financial data, or None
    """
    url = build_url(symbol, statement_type)
    print(f"  📄 Fetching: {url}")

    html = fetch_page(url)
    if html is None:
        return None

    df = parse_financial_table(html)
    if df is None or df.empty:
        print(f"    ⚠️  No data found for {symbol} - {statement_type}")
        return None

    # Clean numeric values
    df = clean_dataframe(df)

    print(f"    ✅ Got {df.shape[0]} rows × {df.shape[1]} columns")
    return df


def scrape_all_statements(symbol: str) -> Dict[str, Optional[pd.DataFrame]]:
    """
    Scrape all financial statements for a given stock.

    Args:
        symbol: Stock ticker symbol

    Returns:
        Dictionary mapping statement type names to DataFrames
    """
    print(f"\n🔍 Scraping financials for {symbol}...")
    results = {}

    for stmt_type, stmt_label in [
        ("income_statement", "Income Statement"),
        ("balance_sheet", "Balance Sheet"),
        ("cash_flow", "Cash Flow"),
        ("ratios", "Financial Ratios"),
        ("kpi_metrics", "KPI Metrics"),
    ]:
        print(f"  📊 {stmt_label}:")
        df = scrape_statement(symbol, stmt_type)
        results[stmt_type] = df
        time.sleep(REQUEST_DELAY)  # Be respectful to the server

    return results
