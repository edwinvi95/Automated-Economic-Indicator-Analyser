# Automated Economic Indicator Analyser

A Python tool that pulls financial market data from Yahoo Finance, cleans it automatically, and exports a formatted multi-sheet Excel report.

---

## What it does

1. Downloads monthly OHLCV (Open, High, Low, Close, Volume) data for a configurable set of tickers via `yfinance`
2. Runs an automated cleaning and validation pipeline on each series
3. Computes derived metrics — normalised price, returns, log returns, moving averages
4. Exports a formatted Excel workbook with four sheets: raw data, cleaned data, summary statistics, and a cleaning change log

---

## How to run

### 1. Install dependencies

```bash
pip install yfinance pandas numpy openpyxl
```

### 2. Configure

Open `analyser_fixed.py` and edit the config block at the top:

```python
TICKERS     = ["^GSPC", "AAPL", "MSFT", "NVDA", "AMD"]  # Any valid Yahoo Finance tickers
START_DATE  = "2024-01-01" # Adjust dates for your requirements
END_DATE    = "2026-05-31"
INTERVAL    = "1mo"   # 1d, 1wk, 1mo
OUTPUT_FILE = "Financial_Data_Report.xlsx"
```

### 3. Run

```bash
python analyser_fixed.py
```

Output: `Financial_Data_Report.xlsx` in the same directory.

---

## Output workbook

| Sheet | Contents |
|---|---|
| **Raw Data** | Unmodified OHLCV data as downloaded |
| **Cleaned Data** | Cleaned data with derived columns (see below) |
| **Summary Stats** | Mean, std dev, min, max close price; % of missing values filled per ticker |
| **Change Log** | Duplicates removed, missing values before/after, outliers detected per ticker |

---

## Cleaning steps (in order)

### 1. Standardise date index
Converts the index to `datetime` and sorts chronologically to ensure consistent time ordering across all tickers.

### 2. Remove duplicates
Drops any rows with duplicate date indices, keeping the first occurrence. Prevents double-counting in any downstream calculation.

### 3. Lowercase column names
Standardises all column headers to lowercase (`Open` → `open`) for consistent programmatic access.

### 4. Fill missing values
Applies forward-fill (`ffill`) first — propagating the last known value forward, which is appropriate for market data over non-trading periods. Then back-fills (`bfill`) any remaining gaps at the start of the series.

### 5. OHLC relationship validation
Flags rows where the price relationships are logically inconsistent — e.g. `High < Close`, `Low > Open`, or negative volume. These are recorded in a `price_error` boolean column rather than silently dropped, so the extent of data quality issues is visible.

### 6. Outlier detection (z-score)
For each OHLCV column, computes the z-score of every observation. Values more than 3 standard deviations from the mean are flagged in a corresponding `*_outlier` boolean column. Flagged rather than removed — outliers in financial data can be genuine extreme events, not just noise.

---

## Derived metrics added after cleaning

| Column | Description |
|---|---|
| `close_norm` | Close price indexed to 100 at the first observation — allows cross-asset comparison on the same scale |
| `return` | Simple period-on-period percentage return |
| `log_return` | Log return — preferred for statistical analysis due to time-additivity |
| `MA_3` | 3-period rolling mean of close price |
| `MA_6` | 6-period rolling mean of close price |

---

## Dependencies

| Package | Purpose |
|---|---|
| `yfinance` | Yahoo Finance data download |
| `pandas` | Data manipulation and cleaning |
| `numpy` | Numerical operations (z-scores, log returns) |
| `openpyxl` | Excel workbook creation and formatting |

---
Still in the process of being updated- to make the excel workbook tidier
