# Codex investment-book review contract

Use this workflow whenever Codex reviews the Waraqah Google Sheet.

## Authority and non-authority

Codex may read the workbook, refresh public data, add evidence-backed review records, and create `Proposed` rows in `Orders`. Codex must never connect to a broker, submit an order, approve its own proposal, mark a row `Part-filled` or `Filled` without broker evidence and a matching Activity transaction ID, choose a Sharia methodology for the owner, infer cash, or invent holdings, trades, FX rates, limits, targets, benchmark returns or facts.

## Read order

1. Read `Checks` settings and unresolved warnings.
2. Read `Activity` and reject duplicate IDs, missing FX, oversells or unsupported transaction types before trusting Portfolio.
3. Reconcile `Portfolio` against Activity; do not edit derived quantities/costs.
4. Read `Sharia`; block every `Buy`/`Add` unless status is `Compliant`, authority, method/version, business and ratio evidence, reporting period and dates are present, and `Next Review` is current. Record purification only under the selected method.
5. Check price/fundamental observation and retrieval times, fiscal periods, publication dates, basis, source URLs and completeness in `Risk & Horizons`.
6. Review 6-month, 2-year and 5-year cases as separate lenses over one shared position. Distinguish facts, estimates/scenarios and agent judgment; record disagreements and do not turn a score into a probability.
7. Size only after cash by currency, pending active orders, broker quantity/order rules, soft/hard position limits and sector limits are known.
8. Write `Buy`, `Add`, `Hold`, `Trim`, `Exit` or `Wait` proposals with stable Proposal ID/version and Review ID. Preserve owner decisions and broker evidence, prevent duplicate versions, and allow at most one active executable order per ticker.
9. Append one immutable review-log row in `Checks` and one valuation point in `Performance`. Never overwrite prior reviews.
10. Verify formulas, validation lists, holdings count, Sharia gates, evidence links and visible layout before reporting completion.

## Minimum proposal evidence

Every non-Wait executable proposal needs: original/reporting currency, broker/account, current/proposed position and sector weights, cash impact and available cash after pending commitments, horizon, order type/quantity/entry/validity, rationale, facts, estimates/scenarios, agent judgment, source URLs, catalyst, downside, review/exit triggers, owner thesis/invalidation, Sharia status/evidence, freshness, confidence and timestamps. If any hard field is unavailable, use `Wait`, leave quantity pending and explain the blocker. A stop is not a guaranteed maximum loss and touching a level is not a fill.

The only valid lifecycle values are `Proposed`, `Approved`, `Submitted`, `Part-filled`, `Filled`, `Cancelled`, `Expired`, and `Superseded`.

## Suggested recurring-review prompt (not scheduled)

> Review the Waraqah investment book in place. Follow CODEX_REVIEW.md and the Checks settings. Refresh only public data whose source and as-of date can be recorded. Reconcile Activity, Portfolio, Sharia and Performance; then update or add evidence-backed Orders proposals for the 6-month, 2-year and 5-year horizons. Never execute trades. Preserve all prior decisions and history. Stay quiet when there is no meaningful change; notify me only about a new blocker, material thesis change, Sharia status/review issue, limit breach, stale-data failure, or proposal needing my decision.
