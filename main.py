"""
Main entry point for the Stock Financial Statement Scraper.

Usage:
    python main.py              # Scrape all stocks in config
    python main.py AAPL         # Scrape a single stock
    python main.py AAPL MSFT    # Scrape multiple stocks
"""

import sys
import time
from typing import Dict, Optional
import pandas as pd

from config import STOCK_SYMBOLS, REQUEST_DELAY
from scraper import scrape_all_statements
from excel_exporter import export_all_stocks, create_summary_file


def print_banner():
    """Print a nice banner."""
    print("=" * 60)
    print("  📈 Stock Financial Statement Scraper")
    print("  Data source: StockAnalysis.com")
    print("=" * 60)


def print_results(all_data: Dict):
    """Print a summary of scraping results."""
    print("\n" + "=" * 60)
    print("  📊 SCRAPING COMPLETE - Summary")
    print("=" * 60)

    for symbol, statements in all_data.items():
        print(f"\n  🏢 {symbol}:")
        for stmt_type, label in [
            ("income_statement", "Income Statement"),
            ("balance_sheet", "Balance Sheet"),
            ("cash_flow", "Cash Flow"),
            ("ratios", "Financial Ratios"),
            ("kpi_metrics", "KPI Metrics"),
        ]:
            df = statements.get(stmt_type)
            if df is not None and not df.empty:
                print(f"    ✅ {label}: {df.shape[0]} rows × {df.shape[1]} periods")
            else:
                print(f"    ❌ {label}: No data")


def main():
    print_banner()

    # Determine which symbols to scrape
    if len(sys.argv) > 1:
        symbols = [s.upper() for s in sys.argv[1:]]
        print(f"\n📋 Scraping specified symbols: {', '.join(symbols)}")
    else:
        symbols = STOCK_SYMBOLS
        print(f"\n📋 Scraping default symbols: {', '.join(symbols)}")

    all_data = {}

    for i, symbol in enumerate(symbols):
        if i > 0:
            print(f"\n⏳ Waiting {REQUEST_DELAY}s before next stock...")
            time.sleep(REQUEST_DELAY)

        statements = scrape_all_statements(symbol)
        all_data[symbol] = statements

    # Print results
    print_results(all_data)

    # Export to Excel
    print("\n" + "=" * 60)
    print("  📦 Exporting to Excel...")
    print("=" * 60)

    files = export_all_stocks(all_data)
    create_summary_file(all_data)

    print(f"\n✅ Done! {len(files)} Excel file(s) created in the 'output/' folder.")
    print("=" * 60)


if __name__ == "__main__":
    main()
