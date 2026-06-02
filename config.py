"""
Configuration for the Stock Financial Statement Scraper.
"""

# Target stock symbols (US stocks)
# Add or remove symbols as needed
STOCK_SYMBOLS = [
    "AAPL",
    "MSFT",
    "GOOGL",
    "AMZN",
    "META",
    "TSLA",
    "NVDA",
    "JPM",
]

# Base URL for StockAnalysis.com financials
BASE_URL = "https://stockanalysis.com/stocks"

# Financial statement types and their URL paths
STATEMENT_TYPES = {
    "income_statement": "financials",
    "balance_sheet": "financials/balance-sheet",
    "cash_flow": "financials/cash-flow-statement",
    "ratios": "financials/ratios",
    "kpi_metrics": "financials/metrics",
}

# Request headers to mimic a browser
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Delay between requests (seconds) to be respectful to the server
REQUEST_DELAY = 1.5

# Output directory for Excel files
OUTPUT_DIR = "output"
