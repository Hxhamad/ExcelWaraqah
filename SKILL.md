---
name: waraqah-investment-book
description: Review and maintain the Waraqah personal Saudi/US investment book in a native Google Sheet, including auditable accounting, Sharia evidence gates, market-data quality, risk analysis, three-horizon research, and non-executable order proposals.
metadata:
  short-description: Auditable Saudi/US portfolio-book review and maintenance
---

# Waraqah Investment Book

Use this skill when working on the owner's Waraqah investment book: the
portfolio, transaction ledger, Sharia screening, Saudi/US market data, risk
and horizon analysis, performance history, or the controlled Orders queue.
The book is a personal recordkeeping and research system. It is not a broker,
an execution system, or a substitute for the owner's Sharia adviser or
investment judgment.

The canonical deliverable is the native Google Sheet named **Waraqah Portfolio
Analysis — 2026-09-09**. When the user refers to “the sheet”, discover and
verify that exact native Sheet through the connected Google Drive/Sheets
workflow before reading or writing. Do not put a private live-sheet URL or
account credentials into source files.

## Non-negotiable boundaries

- Never connect to a broker, place an order, approve your own proposal, or
  claim that a proposal was filled without broker evidence and a matching
  `Activity` transaction ID.
- Never invent a holding, trade, FX rate, cash balance, acquisition date,
  limit, target, benchmark return, company fact, Sharia authority, or Sharia
  methodology. Use `Pending`, `Uncertain`, or `Wait` and record the blocker.
- Never silently change the owner's holdings, decisions, Sharia evidence,
  performance history, or review log. Append or upsert by a stable identifier;
  preserve prior records.
- Treat the `Activity` ledger as the accounting source of truth. Do not type
  balances into derived Portfolio quantity or cost-basis cells.
- A score is a ranking aid, not a probability of return or a recommendation.
  A horizon is a research lens, not a forecast or guarantee.
- Research simulations must remain visibly separate from actual P/L,
  `Activity`, holdings, and cash.

## Workbook topology

The workbook has twelve native tabs. Preserve these names and relationships:

| Tab | Role | Theory boundary |
|---|---|---|
| `Portfolio` | Current positions, value, P/L, weights, Sharia state, horizons and blockers | Derived from `Activity`, market snapshots, `Sharia`, `Risk & Horizons` and `Checks` |
| `Orders` | Ranked research queue for 6-month, 2-year and 5-year lenses | Proposals only; controlled lifecycle and hard gates prevent execution |
| `Stock Lookup` | Single-security report card | Read-through view; it must not become a second source of truth |
| `Performance` | Actual P/L components, dated valuation history, flows and benchmarks | Never manufacture TWR, XIRR, drawdown or FX attribution from insufficient history |
| `Activity` | Transactions, cash, dividends, withholding, fees, tax and corporate actions | Immutable-style source ledger with deterministic transaction IDs |
| `Sharia` | Status, authority, method/version, evidence, review dates and purification | `Uncertain` is not `Compliant`; evidence and owner method are required |
| `Risk & Horizons` | Saudi/US identifiers, price/fundamental observations, technical metrics, rules and horizon views | Facts, estimates/scenarios and agent judgment must be distinguishable |
| `DB` | Annual total-return and fundamental series | Period-labelled source data, not transaction history |
| `Statements` | Annual financial statements | Keep fiscal period labels and missing-data notes visible |
| `Symbols` | Security reference list | Stable symbol/company/exchange mapping |
| `Checks` | Owner settings, integrity gates and review log | Warnings are unresolved owner inputs, not permission to guess |
| `Guide` | User-facing operating rules and calculation definitions | Keep it synchronized with code and workflow changes |

## Theory of the book

### 1. Ledger accounting and cost basis

For each security (s), quantity is the signed accumulation of position
transactions:

`ending quantity = opening balance + buys + transfers in - sells - transfers out`

`Activity` stores positive quantity and price inputs; transaction type supplies
the sign. A split changes quantity by its split factor while preserving the
total cost basis. A transfer changes custody/quantity but is not automatically
profit. An opening balance is a snapshot of what the owner already held; it is
not an asserted historical trade date.

The book uses moving weighted-average cost. Before a sale:

`average cost per share = cost basis before sale / quantity before sale`

The basis removed by a sale is the sold quantity times that average cost. In
reporting currency:

`realized P/L = (sale proceeds - fees - tax) × FX - basis removed`

For a purchase or opening position, the cost basis includes the entered price
and transaction costs, converted using the recorded transaction FX. The
transaction's original currency is retained alongside its SAR conversion.
This prevents later exchange-rate changes from rewriting historical
accounting.

Stable transaction IDs make imports idempotent: importing the same broker file
twice must update the same logical row, not create a duplicate. Prefer the
broker's immutable reference; otherwise derive a deterministic hash from the
account, dates, security, quantity, price, currency, costs and FX.

Cash follows double-entry-like signs: deposits and dividends increase the
relevant currency balance; withdrawals, fees, tax and withholding reduce it.
For any non-SAR record with economic value, a positive contemporaneous
`FX to SAR` is mandatory. Never use a guessed or current FX rate to repair
history.

### 2. Actual performance versus research

The book separates these components:

- market value of open positions;
- open-position cost basis and unrealized P/L;
- realized P/L from closed sales;
- net dividends and withholding;
- standalone fee/tax cash entries;
- external flows such as deposits and withdrawals;
- currency translation/FX effect when the required history exists.

Fees and taxes that are already included in trade basis or proceeds must not be
counted a second time as standalone P/L. The total actual P/L is the sum of the
appropriate realized, unrealized, dividend and standalone components, with
clear notes about what is included.

Simple return is a diagnostic:

`current value / known cost basis - 1`

It is not a money-weighted or time-weighted return when deposits, withdrawals,
or dated valuations are missing. TWR requires a valuation path split around
external cash flows. XIRR requires dated signed cash flows and a terminal
valuation. Drawdown requires a dated value path. If those inputs are
insufficient, display `Pending` and explain why; do not backfill a fictional
history from today's opening snapshot.

Research scenarios, proposed entry prices, simulated exits and hypothetical
returns never alter actual performance, the ledger, or the holdings count.

### 3. Market-data theory

Use a consistent price basis. The current implementation uses yfinance
`auto_adjust=True` and records the resulting adjusted close as a total-return
price basis; cash dividends remain separately identifiable in the listing
currency. Every observation should retain source, URL, observation/as-of time,
retrieval time, fiscal period/publication date where applicable, currency,
exchange, and basis.

The standard technical measures have these meanings:

- annual return: final adjusted close divided by initial adjusted close minus
  one for the stated calendar-year window;
- 12–1 momentum: price around 252 trading sessions ago compared with price
  around 21 sessions ago, excluding the most recent month;
- volatility level: absolute annualized volatility over the longer recent
  window;
- volatility trend: short-window volatility relative to the longer baseline;
- RSI(14): a bounded momentum oscillator, not a valuation measure;
- SMA200 state: current price relative to the 200-trading-day average;
- maximum drawdown: worst peak-to-trough decline in the stated history.

Volatility level and volatility direction are separate concepts. A recent
calming market should not be labelled “high volatility” merely because an
older shock remains in the longer window. Missing history produces `N/A` or
`Pending`, not a fake neutral value.

Data completeness is part of the analysis. Scores below 70% completeness are
non-actionable. Missing inputs are excluded and available weights are
re-normalized; a missing value must never silently become a neutral score.
Stale or weakly sourced data is a blocker for an executable proposal.

### 4. Saudi/US normalization and currency

Use a stable security identity rather than a bare ticker:

- Saudi equities: `SA-<ticker>`, Tadawul, SAR, Asia/Riyadh, Saudi Exchange
  calendar and Saudi equity settlement convention.
- US equities: `US-<ticker>`, exchange such as NYSE or NASDAQ, USD,
  America/New_York, daylight-saving-aware sessions and US settlement
  convention.

The exact broker quantity rule, order type support, fractional-share support,
and account restrictions remain broker-specific and must be confirmed by the
owner. Keep price return, dividend return and USD/SAR translation conceptually
separate. A US security can have a positive local-currency return while its
SAR return differs because of FX; neither effect should be attributed without
the necessary dated FX data.

Do not load the entire US market merely because US support exists. Prioritize
the owner's holdings, serious watchlist candidates, relevant benchmarks and
sector comparators. Use `DB` for annual/quarterly fundamentals and a dedicated
market-history structure for daily prices, corporate actions and FX if daily
history becomes large.

### 5. Sharia evidence as a decision gate

Sharia status is an owner-governed compliance input, not an inference from a
sector label or a model's general knowledge. Each security should have:

- status: `Compliant`, `Non-compliant`, `Uncertain` or `Review overdue`;
- authority/provider;
- methodology and version;
- business-activity evidence;
- financial-ratio evidence;
- evidence URL and reporting period;
- screen date and next review date;
- purification method and, only where supported, purification amount/payment.

`Buy Eligible?` may be `YES` only when the status is `Compliant`, the required
authority/method/version and evidence fields are present, and the next review
is current. `Uncertain`, missing evidence or an overdue review means no
`Buy`/`Add`; use `Wait` and name the missing evidence. Do not choose a Sharia
authority or methodology on the owner's behalf.

### 6. Risk and the three horizons

The owner-defined horizons are:

- short: 6 months;
- medium: 2 years;
- long: 5 years.

These are three analytical lenses over one shared position. They do not create
three independent positions or three pools of capital. A thesis may be
positive over five years and weak over six months without implying that three
orders should be placed. The queue must reconcile horizons to one position and
allow at most one active executable order per ticker.

Risk analysis should consider concentration, sector exposure, cash by
currency, liquidity, drawdown, volatility, thesis invalidation, data quality,
and the owner's soft/hard limits. Position sizing is impossible to authorize
when cash, broker rules or owner limits are unknown. The correct result is
`Wait`, not an invented quantity.

Every research judgment should distinguish:

1. **Facts** — observed prices, reported statements, dated evidence and
   recorded owner inputs.
2. **Estimates/scenarios** — explicitly labelled assumptions or what-if cases.
3. **Agent judgment** — the reasoning that connects the facts to a proposed
   action, including uncertainty and invalidation conditions.

Never turn a composite score into a probability. Never represent a stop level
as a guaranteed maximum loss or imply that touching a level means a fill.

### 7. Orders as a governed research queue

An `Orders` row is a versioned decision record, not an API call. The controlled
actions are `Buy`, `Add`, `Hold`, `Trim`, `Exit` and `Wait`. The controlled
lifecycle is:

`Proposed → Approved → Submitted → Part-filled → Filled`

with `Cancelled`, `Expired` and `Superseded` as terminal alternatives.

Before any proposal can be executable, it needs a stable proposal ID/version,
review ID, horizon, action, order type, quantity, price/entry logic, validity,
currency, broker/account, cash impact, proposed weights, rationale, facts,
scenarios, agent judgment, catalyst, downside, review/exit triggers, Sharia
evidence, freshness, confidence and owner thesis/invalidation criteria. It
must also pass risk, sector, cash, broker and duplicate/horizon-conflict
checks.

If a hard field is missing, leave quantity blank and use `Wait`. An order is
not approved merely because it ranks first. The owner must make the decision.
Only broker evidence plus a matching `Activity` transaction can convert a
submitted order into a filled accounting event.

## Review protocol

For every review or refresh, follow this order:

1. Read `Checks` settings and warnings: reporting currency, timezone, Sharia
   method, position/sector limits, brokers, cash, stale-data threshold and
   cadence.
2. Validate `Activity`: stable-ID uniqueness, supported types, positive FX for
   non-SAR records, no oversells, correct signs and no unexplained duplicates.
3. Reconcile `Portfolio` quantities and cost basis against `Activity`; never
   edit derived position cells.
4. Read `Sharia` and apply the evidence gate before discussing any Buy/Add.
5. Check `Risk & Horizons`, `DB` and `Statements` for observation/retrieval
   times, fiscal periods, sources, price basis, currency and completeness.
6. Review the 6-month, 2-year and 5-year lenses over the shared positions;
   separate facts, scenarios and judgment.
7. Evaluate concentration, sector exposure, drawdown, volatility, cash by
   currency and pending orders. Size only when the owner-approved inputs are
   actually configured.
8. Add or update evidence-backed `Orders` proposal rows with stable version
   and review identifiers. Preserve owner decisions and broker references.
9. Append one review-log record in `Checks` and one dated valuation point in
   `Performance`; do not overwrite historical reviews.
10. Re-read formulas, validation lists, holdings count, Sharia gates, evidence
    links, lifecycle controls and visible layout before reporting completion.

For native Google Sheets work, use bounded ranges and exact visible tab names.
Read metadata before deeper reads, re-read target cells before writes, and
verify the native spreadsheet after each write pass. Prefer the connected
Google Drive/Sheets actions over rebuilding an `.xlsx` when working with the
live native Sheet.

## Allowed and prohibited agent actions

| Allowed | Prohibited |
|---|---|
| Read the workbook and explain its calculations | Place or submit a broker order |
| Refresh public data with source and as-of metadata | Approve a proposal or approve a Sharia method |
| Add evidence-backed research and `Proposed` rows | Mark a row filled without broker evidence and Activity reconciliation |
| Upsert an idempotent ledger import when the owner supplies the source | Invent cash, FX, limits, dates, holdings or missing facts |
| Flag stale data, missing evidence, concentration and horizon conflicts | Convert a score into a probability or guarantee |
| Append dated performance/review records | Rewrite prior history to make performance look complete |

## Owner inputs that unlock analysis

The initial book intentionally leaves several items pending. Ask the owner to
provide or configure them rather than inferring them:

- selected Sharia authority, methodology and version;
- current evidence for each holding and its reporting period/screen date;
- Saudi and US broker/account names and quantity/order rules;
- SAR and USD cash balances with statement dates;
- soft and hard position limits plus sector limits;
- acquisition dates or broker-confirmed historical statements;
- review cadence and notification preference;
- owner thesis and invalidation criteria for each holding or proposal;
- US holdings/watchlist identifiers and the desired amount of price/fundamental
  history.

Until those inputs exist, the safe state is a high-quality auditable book with
`Wait` proposals—not an artificially complete recommendation engine.

## Repository references

- [README.md](README.md): workbook topology, safety model, calculation
  conventions and data limitations.
- [CODEX_REVIEW.md](CODEX_REVIEW.md): compact operational review contract and
  recurring-review prompt.
- [investment_book.py](investment_book.py): testable ledger primitives.
- [metrics.py](metrics.py): N/A-tolerant technical metric definitions.
- [fetcher.py](fetcher.py): Saudi/US normalization, source metadata and
  market-rule defaults.
- [builder.py](builder.py): workbook schema, formulas and formatting.

Before handing off changes, run the repository tests and the workbook verifier,
then confirm that no formula-error strings are present, all original holdings
remain, and the live native Sheet still has the expected tab topology.
