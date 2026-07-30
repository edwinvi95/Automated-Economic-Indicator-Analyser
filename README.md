# Automated Equity Analyser

A self-directed Python project that pulls live market data from Yahoo Finance, runs an automated cleaning and validation pipeline, and exports a formatted multi-sheet Excel report with risk, return, and momentum analytics.

Built during the final year of a BSc Economics degree to demonstrate applied data and finance skills.

---

## What it does

1. Downloads monthly OHLCV (Open, High, Low, Close, Volume) data for a configurable set of tickers via the `yfinance` API
2. Runs an automated cleaning pipeline — deduplication, missing value imputation, OHLC validation, and z-score outlier detection
3. Computes derived metrics — normalised price, simple and log returns, moving averages
4. Calculates a full analytics suite — Sharpe ratio, beta, information ratio, drawdown, momentum signals, and cross-asset correlations
5. Exports a colour-coded, formatted Excel workbook across five sheets

---

## How to run

### 1. Install dependencies

```bash
pip install yfinance pandas numpy openpyxl
```

### 2. Configure

Edit the config block at the top of `analyser.py`:

```python
TICKERS     = ["^GSPC", "AAPL", "MSFT", "NVDA", "AMD"]  # any valid Yahoo Finance tickers
START_DATE  = "2024-01-01"
END_DATE    = "2026-05-31"
INTERVAL    = "1mo"    # 1d, 1wk, or 1mo
OUTPUT_FILE = "Financial_Data_Report.xlsx"
RISK_FREE_RATE = 0.05  # annualised, used for Sharpe and IR calculations
```

### 3. Run

```bash
python analyser.py
```

Output: `Financial_Data_Report.xlsx` saved in the same directory.

---

## Output workbook

| Sheet | Tab colour | Contents |
|---|---|---|
| **Raw Data** | Grey | Unmodified OHLCV data as downloaded from Yahoo Finance |
| **Cleaned Data** | Green | Post-pipeline data including all derived columns; price errors and outliers highlighted in red |
| **Summary Stats** | Blue | Mean, std dev, min, max close price and % of missing values filled per ticker |
| **Change Log** | Red | Audit trail — duplicates removed, missing values before/after cleaning, outliers detected |
| **Analytics** | Purple | Full risk/return summary, momentum signals, correlation matrix, rolling Sharpe, and drawdown series |

---

## Cleaning pipeline

Steps are applied in order to each ticker independently.

### 1. Standardise date index
Converts the index to `datetime` and sorts chronologically, ensuring consistent time ordering across all series before any calculation is run.

### 2. Remove duplicates
Drops rows with duplicate date indices (keeping the first occurrence). Prevents double-counting in return calculations.

### 3. Lowercase column names
Standardises all headers to lowercase (`Open` → `open`) for consistent downstream access.

### 4. Fill missing values
Forward-fills first (`ffill`) — carrying the last known value forward, which is the correct convention for market data over non-trading periods. Back-fills (`bfill`) any remaining gaps at the start of the series.

### 5. OHLC relationship validation
Flags rows where price relationships are logically inconsistent — e.g. `High < Close`, `Low > Open`, `High < Low`, or negative volume. Recorded in a `price_error` boolean column and highlighted red in Excel. Flagged rather than silently dropped so data quality issues remain visible.

### 6. Outlier detection (z-score)
For each OHLCV column, flags values more than 3 standard deviations from the mean in a corresponding `*_outlier` column. Outliers in financial data can be genuine extreme events rather than errors, so they are flagged rather than removed.

---

## Analytics

All analytics use the full cleaned return series. ^GSPC (S&P 500) is the benchmark for beta and information ratio calculations.

### Risk & Return Summary

| Metric | Description |
|---|---|
| Annualised Return (%) | Mean monthly return scaled to annual |
| Annualised Vol (%) | Monthly return std dev scaled to annual (×√12) |
| Sharpe Ratio | Excess return over the risk-free rate per unit of total volatility |
| Beta (vs S&P 500) | `Cov(asset, S&P) / Var(S&P)` — sensitivity to broad market moves |
| Information Ratio | Active return vs S&P 500 divided by tracking error — consistency of outperformance |
| Max Drawdown (%) | Largest peak-to-trough decline over the period |

**Note on Sharpe vs Information Ratio:** Sharpe measures absolute risk-adjusted return (relevant for hedge funds and absolute return strategies). IR measures return relative to a benchmark (relevant for long-only funds judged against an index). Both are included because different types of fund use different metrics.

### Momentum Signals
For the latest period, flags whether each asset is trading above its 3-month and 6-month moving averages. Outputs a BULLISH / BEARISH / MIXED signal per ticker.

### Correlation Matrix
Pairwise return correlations across all tickers. High correlations reduce diversification benefit — an asset with low correlation to the rest of the portfolio adds more risk-reduction value than a high-returning but highly correlated asset.

### Rolling 6-Month Sharpe Ratio
Tracks how risk-adjusted performance has evolved over time, rather than collapsing the entire period into a single figure. Useful for identifying periods of sustained outperformance or drawdown.

### Drawdown Series
Month-by-month percentage decline from each asset's previous peak. Max drawdown is a standard institutional risk metric.

---

## Derived columns (Cleaned Data sheet)

| Column | Description |
|---|---|
| `close_norm` | Close price indexed to 100 at first observation — enables cross-asset comparison on the same scale |
| `return` | Simple period-on-period percentage return |
| `log_return` | Logarithmic return — preferred for statistical analysis due to time-additivity |
| `MA_3` | 3-period rolling mean of close |
| `MA_6` | 6-period rolling mean of close |
| `price_error` | True if any OHLC relationship is violated |
| `*_outlier` | True if the column value exceeds 3 standard deviations from the mean |

---

## Dependencies

| Package | Purpose |
|---|---|
| `yfinance` | Yahoo Finance market data download |
| `pandas` | Data manipulation and cleaning |
| `numpy` | Numerical operations — z-scores, covariance, log returns |
| `openpyxl` | Excel workbook creation, formatting, and number formats |

---

## Project structure

```
.
├── analyser.py                  # Main script
├── Financial_Data_Report.xlsx   # Output (generated on run)
└── README.md                    # This file
```
