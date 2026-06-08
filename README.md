# Financial Analyzer

A personal portfolio analytics dashboard built in Python. Ingests brokerage CSV exports, fetches live market data via Yahoo Finance, and generates a fully interactive HTML dashboard — no subscriptions, no third-party services.

---

## Features

- **Holdings Overview** — real-time price refresh, market value, gain/loss per position
- **Stock Recommendations** — sector-weighted and thematic picks ranked by analyst upside, momentum, and valuation
- **Portfolio Simulation** — model hypothetical buys and instantly see projected P&L
- **Earnings Calendar** — upcoming earnings dates for all holdings and recommendations
- **Price Alerts** — set price targets; alerts persist across sessions via localStorage
- **Market Overview** — live sector heatmap, yield curve, and index strip
- **ETF Picks** — expense-ratio-optimized ETF recommendations to fill sector gaps
- **Excel Export** — full report exported to `.xlsx` alongside the HTML dashboard

---

## Tech Stack

| Layer | Tools |
|---|---|
| Data | Python 3, yfinance, pandas |
| Report Generation | Jinja2 (HTML templating), openpyxl (Excel) |
| Frontend | Vanilla JS, Chart.js, localStorage |
| Live Prices | Yahoo Finance public API (browser-side fetch) |

---

## Getting Started

### 1. Install dependencies
```bash
pip install yfinance pandas jinja2 openpyxl esprima
```

### 2. Add your portfolio
Export a CSV from your brokerage with columns: `Symbol`, `Shares`, `Cost Basis`
(Fidelity and Schwab exports are auto-detected.)

### 3. Run the analyzer
```bash
python analyze.py your_portfolio.csv --cash 5000 --age 22
```
Or double-click **`run_analyze.bat`** (Windows).

### 4. Open the dashboard
```bash
python -m http.server 8080
# then open http://localhost:8080/your_portfolio_report.html
```
Or double-click **`start_dashboard.bat`** (Windows) — launches the server and opens the browser automatically.

---

## Project Structure

```
Financial Analyzer/
├── analyze.py               # Core engine: data fetching, analysis, HTML/Excel generation
├── sample_portfolio.csv     # Example portfolio (8 positions)
├── run_analyze.bat          # One-click: run analysis + update dashboard
├── start_dashboard.bat      # One-click: serve dashboard locally
└── watchlist.json           # Symbols to track beyond your holdings
```

---

## Sample Output

> Dashboard generated from `sample_portfolio.csv` — 8 positions across Tech, Healthcare, Energy, and Financials.

- Portfolio value: **$55,912**
- 18 buy recommendations across sector and thematic strategies
- Live price refresh on every page load

---

## Notes

- No API keys required — uses Yahoo Finance's public endpoints
- All data stays local — nothing is sent to any external server
- Live browser price refresh requires serving via `localhost` (not `file://`) due to CORS
