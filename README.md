# Automated Equity Analyser

A small Python project that pulls live market data from Yahoo Finance, runs an automated cleaning and validation pipeline, and exports a formatted multi-sheet Excel report with risk, return, and momentum analytics.

---

## Output files

Running the script produces two files:

| File | Contents |
|---|---|
| `Financial_Data_Report.xlsx` | All data sheets — raw, cleaned, summary stats, change log, analytics, and chart data tables |
| `Financial_Data_Report_charts.xlsx` | Interactive line charts — normalised price trend and drawdown from peak |

Both files update automatically when tickers are added or removed from the config.

---

## How to run

### 1. Install dependencies

```bash
pip install yfinance pandas numpy openpyxl xlsxwriter
```

### 2. Configure

Edit the config block at the top of `analyser.py`:

```python
TICKERS        = ["^GSPC", "AAPL", "MSFT", "NVDA", "AMD", "^VIX"]
START_DATE     = "2024-01-01"
END_DATE       = "2026-05-31"
INTERVAL       = "1mo"       # 1d, 1wk, or 1mo
OUTPUT_FILE    = "Financial_Data_Report.xlsx"
RISK_FREE_RATE = 0.05        # annualised, used for Sharpe and IR calculations
```

Any valid Yahoo Finance ticker can be added to `TICKERS`. All sheets and charts update automatically.

### 3. Run

```bash
python analyser.py
```

> **Note:** delete any existing `Financial_Data_Report.xlsx` before the first run to avoid carrying forward stale data.

---

## Workbook sheets

| Sheet | Tab colour | Contents |
|---|---|---|
| **Raw Data** | Grey | Unmodified OHLCV data as downloaded from Yahoo Finance |
| **Cleaned Data** | Green | Post-pipeline data with derived columns; price errors and outliers highlighted in red |
| **Summary Stats** | Blue | Mean, std dev, min, max close price and % of missing values filled per ticker |
| **Change Log** | Red | Audit trail — duplicates removed, missing values before/after, outliers flagged |
| **Analytics** | Purple | Full risk/return table, momentum signals, correlation matrix, rolling Sharpe, drawdown series |
| **Charts** | Orange | Source data tables used to build the companion charts file |

---

## Cleaning pipeline

Applied independently to each ticker in order.

### 1. Standardise date index
Converts the index to `datetime` and sorts chronologically, ensuring consistent time ordering before any calculation is run.

### 2. Remove duplicates
Drops rows with duplicate date indices, keeping the first occurrence. Prevents double-counting in return calculations.

### 3. Lowercase column names
Standardises all headers to lowercase (`Open` → `open`) for consistent downstream access.

### 4. Fill missing values
Forward-fills first (`ffill`) — carrying the last known value forward, the correct convention for market data over non-trading periods. Back-fills (`bfill`) any remaining gaps at the start of the series.

### 5. OHLC relationship validation
Flags rows where price relationships are logically inconsistent — e.g. `High < Close`, `Low > Open`, `High < Low`, or negative volume. Recorded in a `price_error` boolean column and highlighted red in Excel. Flagged rather than dropped so data quality issues remain visible.

### 6. Outlier detection (z-score)
For each OHLCV column, flags values more than 3 standard deviations from the mean in a `*_outlier` boolean column. Outliers in financial data can be genuine extreme events, so they are flagged rather than removed.

---

## Analytics

All metrics use the full cleaned return series. `^GSPC` (S&P 500) is the benchmark for beta and information ratio. All return and volatility figures are annualised from monthly data (×12 and ×√12 respectively).

### Risk & Return Summary

| Metric | Description |
|---|---|
| Annualised Return (%) | Mean monthly return scaled to annual |
| Annualised Vol (%) | Monthly return std dev scaled to annual (×√12) |
| Sharpe Ratio | Excess return over the risk-free rate per unit of total volatility — measures absolute risk-adjusted performance |
| Beta (vs S&P 500) | `Cov(asset, S&P) / Var(S&P)` — sensitivity to broad market moves. >1 means more volatile than the index; <1 means more defensive |
| Information Ratio | Active return vs S&P 500 divided by tracking error — measures consistency of outperformance relative to the benchmark |
| Max Drawdown (%) | Largest peak-to-trough decline over the period |

**Sharpe vs Information Ratio:** Sharpe measures absolute risk-adjusted return and is the standard metric for hedge funds and absolute return strategies. IR measures performance relative to a benchmark and is how long-only fund managers are typically evaluated. Both are included because different mandates use different metrics.

### Momentum Signals
For the latest period, flags whether each asset is trading above its 3-month and 6-month moving averages. Outputs BULLISH / BEARISH / MIXED per ticker.

### Return Correlation Matrix
Pairwise return correlations across all tickers. High correlations reduce diversification benefit — an asset with low correlation to the rest of the portfolio contributes more risk-reduction value regardless of its individual return.

### Rolling 6-Month Sharpe Ratio
Tracks how risk-adjusted performance has evolved over time rather than collapsing the period into a single figure. Useful for identifying whether outperformance was sustained or concentrated in one period.

### Drawdown Series
Month-by-month percentage decline from each asset's previous peak. Max drawdown is a standard institutional risk metric used to assess downside exposure.

---

## Charts (companion file)

The `_charts.xlsx` file is generated using `xlsxwriter` rather than `openpyxl` to avoid Excel repair warnings caused by openpyxl's chart XML implementation.

| Chart | Description |
|---|---|
| Normalised Price Trend | All tickers indexed to 100 at the start date — enables direct comparison across assets with very different price levels |
| Drawdown from Peak | Peak-to-trough decline for each asset over time — visually shows the risk profile of each position |

Both charts include a line per ticker and update automatically when the ticker list changes.

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
| `xlsxwriter` | Chart generation — produces clean Excel-compatible chart XML |

---
