# ورقة — Waraqah

**A self-serve stock analysis workbook for the Saudi Exchange (Tadawul). Type any stock code, read a full report. No AI, no subscriptions, no API keys.**

Waraqah (Arabic for "sheet of paper") builds a 7-sheet Excel workbook covering **202 Tadawul-listed companies** with price history back to 2015, fundamentals 2021–2025, risk metrics, and a research-backed scoring model adapted to the Saudi market. You refresh it with one double-click; everything is computed locally by a Python pipeline.

![Sheets](https://img.shields.io/badge/sheets-7-blue) ![Stocks](https://img.shields.io/badge/stocks-202-green) ![Python](https://img.shields.io/badge/python-3.10%2B-informational)

## What you get

| Sheet | Purpose |
|---|---|
| **Portfolio** | Enter your holdings (code, shares, cost) → live market value, P/L, weights, concentration flags |
| **Stock Lookup** | Type any 4-digit Tadawul code → full Arabic report card: returns (1W→1Y), RSI(14), volatility regime, SMA200 position, P/E, ROE, dividends, 2021–2025 statements, composite score and verdict |
| **Risk & Horizons** | All 202 symbols: returns, drawdown, oil-beta proxy by sector, and قريب/متوسط/بعيد (short/mid/long) verdicts |
| **DB** | Market database: annual return, volatility, max drawdown, dividends, momentum, P/E per symbol per year since 2015 |
| **Statements** | Revenue, net income, EPS, ROE, D/E, payout 2021–2025 for every symbol |
| **Symbols** | Full Tadawul reference list with Arabic/English names and sectors |
| **Guide** | Arabic manual: how each metric works, the scoring model, risk rules, refresh instructions |

Auto-filters are enabled on every table sheet. Missing fundamentals render as "بيانات ناقصة" (neutral) — never as `#N/A` errors.

## Quick start

```bash
pip install -r requirements.txt
python -m pytest tests/ -v              # 10 unit tests
python builder.py                       # build the workbook from the included seed data
```

Or just double-click **`refresh.bat`** → `1` (quick: re-price your portfolio, ~2 min) or `2` (full market rebuild, ~20–40 min). The finished file opens as `Sahm_Portfolio_Analysis_v2.xlsx` next to the repo folder — open it in Excel or LibreOffice and start typing codes.

To add your real holdings: type your 4-digit codes, shares, and average cost into the Portfolio sheet, then run `refresh.bat` → `1`. Your entries are preserved across rebuilds.

## Scoring model (Saudi-adapted)

Composite score 0–100 = **Value 30% + Quality 20% + Technical 20% + Dividend 15% + Risk 15%**.

The weights follow published evidence on Tadawul: value (B/M) is the only factor with a significant premium across all examined portfolios (Alkhareif 2016; Alshaikhmubarek 2024), dividend yield carries a documented premium, momentum is weaker than in US markets (tactical overlay only), and small-cap tilts are an illiquidity trap. Risk rules use the 200-DMA trend filter (max drawdown roughly halved in backtests), volatility-managed exposure (Moreira & Muir 2017), 10–15 position diversification, and oil-beta balancing — energy is ~60% of TASI, so petchem vs. banks vs. consumer balance matters.

Ratings: ≥80 شراء قوي · 65–79 شراء · 50–64 تعزيز/احتفاظ · 35–49 بيع · <35 بيع قوي.

## How it works

```
fetcher.py → yfinance (.SR tickers) → data/*.csv (resumable cache)
builder.py → openpyxl builds all 7 sheets (Excel-2007-era formulas only)
LibreOffice headless → recalculates formula cache → verify_workbook.py (19 checks)
           → promoted only if every check passes (fail-safe keeps the previous file)
```

- `metrics.py` — pure metric functions, TDD-tested (returns, volatility, drawdown, RSI, SMA200 flag, volatility regime, momentum, composite scoring)
- `fetcher.py` — resilient fetcher: every yfinance call guarded, sleep + retries, resume support
- `validate_symbols.py` — one-off tool that live-validates the symbol universe (drops Nomu codes Yahoo doesn't carry)
- `verify_workbook.py` — 19 pinned-assert checks, exit 0/1

## Data sources & honest limitations

- **yfinance `.SR`** — daily prices back to ~2010, full dividend history; annual financial statements capped at **2021–2025** (Yahoo limit; deeper backfill would require scraping Argaam/Tadawul, deliberately out of scope)
- **No TASI index history** via yfinance — benchmark with portfolio self-metrics instead
- Quotes are delayed; some symbols have missing fundamentals (REITs often lack P/E) — shown as بيانات ناقصة and scored neutral
- Symbol list parsed from the Arabic Wikipedia Tadawul listings page, then validated live against Yahoo

## Disclaimer

This tool is for personal research and education. It is **not investment advice**. Scores are mechanical factor rankings, not recommendations — do your own due diligence.

## License

MIT
