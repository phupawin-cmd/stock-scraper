# 📈 Stock Financial Statement Scraper

ดึงงบการเงินแบบละเอียดจาก **StockAnalysis.com** และ export เป็น **Excel** ที่มี formatting สวยงาม พร้อมใช้งาน

## 🚀 Features

- ✅ **5 ประเภทงบ**: Income Statement, Balance Sheet, Cash Flow, Financial Ratios, KPI Metrics
- ✅ **Annual + Quarterly**: ดึงทั้งงบรายปีและรายไตรมาส
- ✅ **Period Ending เป็นวันที่จริง**: แสดงวันที่สิ้นสุดงวดแบบ `YYYY-MM-DD`
- ✅ **TTM**: คอลัมน์ล่าสุดของ Annual แสดงเป็น "TTM"
- ✅ **Clean numeric**: จัดการ `M`/`B`/`K`, วงเล็บ (ค่าลบ), `%` ให้โดยอัตโนมัติ
- ✅ **Excel สวยงาม**: Freeze panes, Auto-filter, สลับสีแถว, ไฮไลท์แถวสำคัญ, ตัวเลขติดลบสีแดง, % format อัตโนมัติ
- ✅ **Batch scraping**: ดึงหลายหุ้นพร้อมกัน
- ✅ **Respectful scraping**: หน่วงเวลาระหว่าง request (config ได้)
- ✅ **Summary file**: ไฟล์สรุปภาพรวมทุกหุ้น

---

## 📦 การติดตั้ง

```bash
pip install -r requirements.txt
```

**Requirements:** `requests`, `beautifulsoup4`, `lxml`, `pandas`, `openpyxl`

---

## 🔧 การตั้งค่า (`config.py`)

```python
# หุ้นที่จะ scrape (เมื่อไม่ระบุใน command line)
STOCK_SYMBOLS = ["AAPL", "MSFT", "GOOGL", "AMZN", "META", "TSLA", "NVDA", "JPM"]

# หน่วงเวลาระหว่าง request (วินาที)
REQUEST_DELAY = 1.5

# โฟลเดอร์สำหรับไฟล์ Excel
OUTPUT_DIR = "output"
```

---

## ▶️ วิธีใช้งาน

### Scrape ทุกหุ้นใน config
```bash
python main.py
```

### Scrape เฉพาะบางหุ้น
```bash
python main.py AAPL
python main.py AAPL MSFT ANET
```

### ตัวอย่างผลลัพธ์
```
============================================================
  📈 Stock Financial Statement Scraper
  Data source: StockAnalysis.com
============================================================

📋 Scraping specified symbols: ANET

🔍 Scraping Annual financials for ANET...
  📊 Income Statement:       ✅ 34 rows × 6 periods
  📊 Balance Sheet:          ✅ 38 rows × 6 periods
  📊 Cash Flow:              ✅ 32 rows × 6 periods
  📊 Financial Ratios:       ✅ 35 rows × 6 periods
  📊 KPI Metrics:            ✅ 10 rows × 15 periods

🔍 Scraping Quarterly financials for ANET...
  📊 Income Statement:       ✅ 34 rows × 20 periods
  📊 Balance Sheet:          ✅ 38 rows × 20 periods
  📊 Cash Flow:              ✅ 33 rows × 20 periods
  📊 Financial Ratios:       ✅ 35 rows × 20 periods
  📊 KPI Metrics:            ✅ 9 rows × 24 periods

💾 Saved: output\ANET_financials.xlsx
```

---

## 📁 โครงสร้าง Output

```
output/
├── AAPL_financials.xlsx
├── ANET_financials.xlsx
├── MSFT_financials.xlsx
└── _summary.xlsx
```

---

## 📊 โครงสร้างไฟล์ Excel

แต่ละไฟล์มี **10 sheets** (5 ประเภท × 2 periods):

| Sheet | เนื้อหา |
|---|---|
| **Income (A)** | Annual — รายได้, ต้นทุน, กำไร, EPS... |
| **Income (Q)** | Quarterly — รายได้, ต้นทุน, กำไร, EPS... |
| **Balance Sheet (A)** | Annual — สินทรัพย์, หนี้สิน, ส่วนของผู้ถือหุ้น... |
| **Balance Sheet (Q)** | Quarterly — สินทรัพย์, หนี้สิน, ส่วนของผู้ถือหุ้น... |
| **Cash Flow (A)** | Annual — กระแสเงินสดดำเนินงาน/ลงทุน/จัดหา, FCF... |
| **Cash Flow (Q)** | Quarterly — กระแสเงินสดดำเนินงาน/ลงทุน/จัดหา, FCF... |
| **Ratios (A)** | Annual — PE, PB, ROE, ROA, Debt/Equity, Current Ratio... |
| **Ratios (Q)** | Quarterly — PE, PB, ROE, ROA, Debt/Equity, Current Ratio... |
| **KPI Metrics (A)** | Annual — Revenue, EPS, Margins, FCF, Dividends... |
| **KPI Metrics (Q)** | Quarterly — Revenue, EPS, Margins, FCF, Dividends... |

### 🎨 Excel Formatting Features

| Feature | คำอธิบาย |
|---|---|
| 🔒 **Freeze Panes** | ตรึงแถว header + คอลัมน์แรก ไว้ตลอดตอน scroll |
| 🔍 **Auto-Filter** | กรอง/เรียงข้อมูลได้ทุกคอลัมน์ |
| 🎨 **Key Row Highlight** | แถวสำคัญ (Revenue, Net Income, Total Assets, EBITDA...) พื้นหลังฟ้า ตัวหนา |
| 📅 **Period Ending** | แถววันที่พื้นหลังเหลือง ตัวเอียง — แยกชัดจากข้อมูล |
| 🔴 **Negative Numbers** | ตัวเลขติดลบแสดงเป็น**สีแดง** |
| 📊 **Smart % Format** | แถว % (Margin, Growth, ROE, Yield...) ฟอร์แมต `0.00%` อัตโนมัติ |
| 🏷️ **Alternating Rows** | สีพื้นหลังสลับแถวคู่/คี่ |
| 📐 **Print Ready** | ตั้งค่าแนวนอน (Landscape), Fit to Width |

---

## 🛠️ โครงสร้างโปรเจค

| ไฟล์ | หน้าที่ |
|---|---|
| `main.py` | Entry point — ควบคุม flow การ scrape + export |
| `config.py` | ตั้งค่า symbols, delay, output dir, URLs |
| `scraper.py` | Core scraping — fetch HTML, parse tables, clean data |
| `excel_exporter.py` | Excel export — formatting, styles, sheet creation |
| `requirements.txt` | Python dependencies |

---

## 🧪 การใช้จาก Python โดยตรง

```python
from scraper import scrape_all_statements, scrape_statement
from excel_exporter import export_single_stock

# ดึงเฉพาะงบรายปี
annual = scrape_all_statements("AAPL", period="annual")

# ดึงเฉพาะงบรายไตรมาส
quarterly = scrape_all_statements("AAPL", period="quarterly")

# ดึงแค่ income statement รายปี
income_df = scrape_statement("AAPL", "income_statement", period="annual")
print(income_df.head())

# Export เป็น Excel
all_data = {"AAPL": {"Annual": annual, "Quarterly": quarterly}}
export_single_stock("AAPL", all_data)
```

### Statement types ที่ใช้ได้
- `income_statement`
- `balance_sheet`
- `cash_flow`
- `ratios`
- `kpi_metrics`

---

## 🔗 แหล่งข้อมูล

| ข้อมูล | URL |
|---|---|
| Income Statement | `stockanalysis.com/stocks/{symbol}/financials/` |
| Balance Sheet | `stockanalysis.com/stocks/{symbol}/financials/balance-sheet/` |
| Cash Flow | `stockanalysis.com/stocks/{symbol}/financials/cash-flow-statement/` |
| Ratios | `stockanalysis.com/stocks/{symbol}/financials/ratios/` |
| KPI Metrics | `stockanalysis.com/stocks/{symbol}/financials/metrics/` |

> 📝 ต่อท้าย `?p=quarterly` เพื่อดึงข้อมูลรายไตรมาส

---

## ⚠️ Disclaimer

เครื่องมือนี้เพื่อ**การศึกษาและใช้งานส่วนตัว**เท่านั้น กรุณาเคารพ:
- Terms of Service ของเว็บไซต์
- ข้อจำกัด robots.txt
- Rate limits (ปรับ `REQUEST_DELAY` ได้ใน `config.py`)
