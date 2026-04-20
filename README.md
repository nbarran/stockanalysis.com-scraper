# StockAnalysis.com Scraper — v1.1

**Personal Use Only**
Built and maintained by Nick Barran — [nbarran@uw.edu](mailto:nbarran@uw.edu)

---

## Overview

A desktop GUI application that scrapes financial data from [stockanalysis.com](https://stockanalysis.com) and saves it locally as Excel or CSV files. Built with Python and tkinter — no browser or API key required.

---

## Core Features

### Ticker Input
- Enter one or multiple tickers separated by commas, spaces, or semicolons (e.g. `AAPL, MSFT, NVDA`)
- Input is automatically cleaned and normalized — capitalization and spacing errors are handled automatically
- Duplicate tickers are deduplicated before running

### Statement Selection
Choose any combination of the following statements to download:
- **Overview** — key stats (market cap, P/E, EPS, dividend, volume, etc.)
- **Income Statement**
- **Balance Sheet**
- **Cash Flow Statement**
- **Ratios**

### Period
- **Annual** — full-year data (default)
- **Quarterly** — quarter-by-quarter data
- **Both** — annual and quarterly in separate sheets/files

### Output Mode
- **Combined Excel** (default) — all selected statements saved as tabs in a single `.xlsx` workbook per ticker
- **Separate Files** — one file per statement, in either CSV or Excel format

### Output Format
- **Excel (.xlsx)** — includes metadata header (ticker, date, time, currency, source) and auto-fitted column widths
- **CSV** — plain comma-separated format

### Folder Structure
Each run creates a timestamped folder containing a subfolder per ticker:
```
stockscraper/
└── 2026-04-20_02-30-15PM/
    ├── AAPL/
    │   └── AAPL_combined.xlsx
    └── MSFT/
        └── MSFT_combined.xlsx
```

### Performance
- All statements are fetched simultaneously using a thread pool (8 workers)
- Persistent HTTP sessions reuse TCP connections, eliminating SSL handshake overhead per request
- No artificial delay between requests
- Single ticker with all statements on Both periods completes in approximately 3-5 seconds

### Error Handling
- Invalid or unrecognized tickers are detected immediately and skipped
- Any partially created folder for a failed ticker is automatically deleted
- Other tickers in the same run continue unaffected

### Progress Bar
- Blue while running
- Green on successful completion
- Red if any ticker encountered an error
- Resets on each new run

### Collapsible Log
A dropdown log panel shows real-time status for each statement as it is scraped and saved.

### Saved Preferences
The output folder path is automatically saved when the app is closed and restored on next launch.

---

## Requirements

```
pip install requests pandas beautifulsoup4 lxml openpyxl
```

---

## Running the App

**From source:**
```
python stockanalysis_gui_v1_1.py
```

**As a compiled executable (Windows):**
```
& "C:\Users\ndbar\AppData\Local\Programs\Python\Python314\python.exe" -m PyInstaller --onefile --windowed "C:\Users\ndbar\PyCharmMiscProject\stockanalysis_gui_v1_1.py"
```
The `.exe` will be output to the `dist/` folder.

---

## Contact

**Nick Barran**
University of Washington — Foster School of Business
[nbarran@uw.edu](mailto:nbarran@uw.edu)

---

*This tool is intended for personal academic and research use only. Data is sourced from stockanalysis.com — please respect their terms of service.*
