# 📈 Stock Financial Statement Scraper

Scrape detailed financial statements (Income Statement, Balance Sheet, Cash Flow) from **StockAnalysis.com** and export to beautifully formatted **Excel** files.

## 🚀 Features

- Scrape complete financial statements for US stocks
- **3 statement types**: Income Statement, Balance Sheet, Cash Flow
- Clean numeric parsing (handles M/B/K suffixes, parentheses for negatives, percentages)
- Export to **Excel (.xlsx)** — each statement in a separate sheet
- Professional formatting: headers, alternating row colors, auto-column width
- Batch scraping for multiple stocks
- Respectful scraping with configurable delays
- Summary file with data overview

## 📦 Installation

```bash
pip install -r requirements.txt
```

## 🔧 Configuration

Edit `config.py` to customize:

- `STOCK_SYMBOLS` — List of stock tickers to scrape
- `REQUEST_DELAY` — Time between requests (default: 1.5s)
- `OUTPUT_DIR` — Where Excel files are saved

## ▶️ Usage

### Scrape all stocks in config
```bash
python main.py
```

### Scrape specific stock(s)
```bash
python main.py AAPL
python main.py AAPL MSFT GOOGL
```

## 📁 Output Structure

```
output/
├── AAPL_financials.xlsx      # Apple - 3 sheets
├── MSFT_financials.xlsx      # Microsoft - 3 sheets
├── GOOGL_financials.xlsx     # Google - 3 sheets
└── _summary.xlsx             # Overview of all data
```

### Each Excel file contains:
| Sheet | Content |
|-------|---------|
| Income Statement | Revenue, COGS, Operating Income, Net Income, EPS... |
| Balance Sheet | Total Assets, Liabilities, Equity, Cash... |
| Cash Flow | Operating/Investing/Financing Cash Flows, FCF... |

## 🛠️ Tech Stack

- **Python 3.9+**
- `requests` — HTTP requests
- `beautifulsoup4` + `lxml` — HTML parsing
- `pandas` — Data manipulation
- `openpyxl` — Excel export with formatting

## ⚠️ Disclaimer

This tool is for **educational/personal use only**. Always respect:
- Website's Terms of Service
- robots.txt restrictions
- Rate limits (adjust `REQUEST_DELAY` as needed)
