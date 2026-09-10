# ورقة — Waraqah

Waraqah is a personal, auditable Saudi/US investment book. The reporting currency is SAR, every transaction retains its original currency and FX rate, and no component can place an order with a broker.

The workbook contains 12 native tabs:

| Sheet | Purpose |
|---|---|
| **Portfolio** | Positions, value, P/L, weights, Sharia state, three horizons, blockers |
| **Orders** | Ranked Buy/Add/Hold/Trim/Exit/Wait queue with versioning, sizing gates and broker reconciliation |
| **Stock Lookup** | Single-security report card |
| **Performance** | Actual P/L components, external flows, concentration/turnover, dated TWR history and benchmarks |
| **Activity** | Source-of-truth transaction and cash ledger with stable transaction IDs |
| **Sharia** | Authority/method/version, business and ratio evidence, dates, change review and purification ledger |
| **Risk & Horizons** | Saudi/US identifiers, fiscal/source/basis metadata, market rules, 6m/2y/5y views and freshness |
| **DB** | Annual total-return and fundamental series |
| **Statements** | Period-labelled annual financial statements |
| **Symbols** | Security reference list |
| **Checks** | Owner settings, Sharia/risk/data integrity gates, immutable review log |
| **Guide** | Arabic operating guide and calculation definitions |

## Safety model

- `Activity` is the accounting source of truth. Portfolio quantity and cost basis are formulas, not editable balances.
- Broker imports are upserted by deterministic `Transaction ID`, so reimporting the same activity does not duplicate it.
- Sales use moving weighted-average cost and realized P/L is separate from unrealized P/L.
- A non-SAR movement without a positive `FX to SAR` is rejected.
- `Buy` and `Add` proposals are blocked unless Sharia status is `Compliant`, authority/method/version, evidence, reporting period and screen dates are present, and the next review has not expired.
- Executable proposals also remain blocked while cash, broker/account rules, quantity, owner thesis or owner-approved position and sector limits are missing.
- Missing model inputs are excluded and the available weights are re-normalized. Scores below 70% data completeness are labelled non-actionable, not neutral.
- Orders are review records only. Lifecycle values are `Proposed`, `Approved`, `Submitted`, `Part-filled`, `Filled`, `Cancelled`, `Expired`, and `Superseded`. A broker order ID and fill evidence are required before a fill can reconcile to `Activity`.
- The three horizons are research lenses over one shared position. A control prevents more than one active executable order for the same ticker.

## Calculation conventions

- Price history: Yahoo Finance `auto_adjust=True`, explicitly recorded as an adjusted-close total-return basis.
- Annual return: last adjusted close / first adjusted close - 1 for the stated calendar-year window.
- Dividends: separately retained as cash per share in the listing currency.
- Momentum 12–1: price at `t-21` trading sessions divided by price at `t-252`; the most recent month is excluded.
- Volatility level: absolute 60-session annualized volatility. Volatility trend: the 20-session level relative to the 60-session baseline. They are separate.
- Horizons: 6 months, 2 years and 5 years. A horizon is a structured view, not a prediction or guaranteed target.
- Saudi default conventions: SAR, Asia/Riyadh, Saudi Exchange calendar and [T+2 equity settlement](https://www.saudiexchange.sa/wps/portal/saudiexchange/trading/market-services/equities?locale=en).
- US default conventions: USD, America/New_York with daylight saving and [T+1 standard settlement](https://www.sec.gov/rules-regulations/2023/02/34-96930).
- US support is normalization/accounting-tested without inserting fictional US holdings. Fractional-share and order capabilities remain broker-specific.

## Quick start

```bash
pip install -r requirements.txt
python -m pytest -q
python builder.py --out Waraqah.xlsx --preserve previous_Waraqah.xlsx
python verify_workbook.py Waraqah.xlsx --data-dir data
```

The `--preserve` workflow carries forward Activity, Sharia evidence, Orders decisions, Performance history, owner settings and the review log while rebuilding formula/helper columns from current code.

See `CODEX_REVIEW.md` for the reusable agent review contract. Review scheduling is intentionally not activated until the owner chooses a cadence.

## Data limitations

Yahoo Finance is a free, delayed and sometimes incomplete source. Every refresh records observation/retrieval time, price/return basis, fiscal period, source and URL. Missing or stale data is surfaced as a blocker; the last valid record is preserved. TWR, XIRR, drawdown, FX attribution, benchmarks and proposal sizing stay pending until the required broker/history or owner policy inputs exist.

This tool is for personal recordkeeping and research, not investment advice.
