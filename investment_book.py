"""Accounting primitives for the Waraqah Saudi/US investment book.

The spreadsheet is the user-facing book; this module is the testable reference
for stable transaction IDs, idempotent imports, weighted-average cost basis,
realized P/L, cash balances, FX conversion, and split handling.

Quantities and prices are entered as positive values.  Transaction ``type``
determines the sign.  Reporting currency is SAR and each record retains its
original currency and the contemporaneous SAR FX rate used for accounting.
"""

from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from datetime import date, datetime
import hashlib
import json


POSITION_TYPES = {"Opening balance", "Buy", "Sell", "Transfer in", "Transfer out"}
CASH_IN_TYPES = {"Deposit", "Dividend", "Interest", "FX buy", "Corporate action"}
CASH_OUT_TYPES = {"Withdrawal", "Fee", "Tax", "Withholding", "FX sell"}


def _canonical(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if value is None:
        return ""
    if isinstance(value, float):
        return format(value, ".12g")
    return str(value).strip()


def stable_transaction_id(record):
    """Return a deterministic ID; prefer the broker's immutable reference."""
    supplied = _canonical(record.get("transaction_id"))
    if supplied:
        return supplied
    broker_ref = _canonical(record.get("broker_reference"))
    account = _canonical(record.get("account"))
    if broker_ref:
        payload = {"account": account, "broker_reference": broker_ref}
    else:
        fields = (
            "account", "broker", "trade_date", "settlement_date", "type",
            "security_id", "exchange", "quantity", "price", "currency",
            "fees", "tax", "fx_to_sar",
        )
        payload = {field: _canonical(record.get(field)) for field in fields}
    digest = hashlib.sha256(
        json.dumps(payload, ensure_ascii=False, sort_keys=True,
                   separators=(",", ":")).encode("utf-8")
    ).hexdigest()[:20].upper()
    return "TX-" + digest


def upsert_transactions(existing, incoming):
    """Idempotently merge imports by stable ID; incoming values win by field."""
    merged = {}
    order = []
    for source in (existing or [], incoming or []):
        for original in source:
            row = deepcopy(original)
            txid = stable_transaction_id(row)
            row["transaction_id"] = txid
            if txid not in merged:
                order.append(txid)
                merged[txid] = row
            else:
                merged[txid].update(
                    {k: v for k, v in row.items() if v not in (None, "")}
                )
    return [merged[txid] for txid in order]


def _number(record, field, default=0.0):
    value = record.get(field, default)
    try:
        return float(value if value not in (None, "") else default)
    except (TypeError, ValueError):
        raise ValueError("%s must be numeric" % field)


def apply_ledger(transactions):
    """Apply transactions in supplied chronological order.

    Uses moving weighted-average cost.  Each returned row contains signed
    quantity, SAR cash movement, SAR cost-basis movement and realized P/L.
    Unknown FX on a non-SAR money movement is rejected rather than guessed.
    """
    quantity = defaultdict(float)
    cost_sar = defaultdict(float)
    cash = defaultdict(float)
    applied = []

    for original in transactions or []:
        row = deepcopy(original)
        row["transaction_id"] = stable_transaction_id(row)
        kind = _canonical(row.get("type"))
        security = _canonical(row.get("security_id"))
        currency = (_canonical(row.get("currency")) or "SAR").upper()
        qty = _number(row, "quantity")
        price = _number(row, "price")
        fees = _number(row, "fees")
        tax = _number(row, "tax")
        amount = _number(row, "cash_amount")
        fx = _number(row, "fx_to_sar", 1.0 if currency == "SAR" else 0.0)
        if currency != "SAR" and fx <= 0 and (qty or price or fees or tax or amount):
            raise ValueError("positive fx_to_sar required for non-SAR transaction")
        if currency == "SAR" and fx <= 0:
            fx = 1.0

        signed_qty = 0.0
        basis_change = 0.0
        realized = 0.0
        cash_change_original = 0.0

        if kind in ("Opening balance", "Buy", "Transfer in"):
            if not security or qty < 0:
                raise ValueError("position inflow needs security_id and non-negative quantity")
            signed_qty = qty
            basis_change = (qty * price + fees + tax) * fx
            if kind == "Buy":
                cash_change_original = -(qty * price + fees + tax)
        elif kind in ("Sell", "Transfer out"):
            if not security or qty < 0 or qty > quantity[security] + 1e-9:
                raise ValueError("position outflow exceeds available quantity")
            average = cost_sar[security] / quantity[security] if quantity[security] else 0.0
            signed_qty = -qty
            basis_change = -(qty * average)
            if kind == "Sell":
                cash_change_original = qty * price - fees - tax
                realized = cash_change_original * fx + basis_change
        elif kind == "Split":
            factor = _number(row, "split_factor")
            if not security or factor <= 0:
                raise ValueError("split_factor must be positive")
            signed_qty = quantity[security] * (factor - 1.0)
        else:
            if kind in CASH_IN_TYPES:
                cash_change_original = amount - fees - tax
            elif kind in CASH_OUT_TYPES:
                cash_change_original = -(amount + fees + tax)
            elif kind != "":
                raise ValueError("unsupported transaction type: %s" % kind)

        quantity_before = quantity[security] if security else 0.0
        cost_before = cost_sar[security] if security else 0.0
        if security:
            quantity[security] += signed_qty
            cost_sar[security] += basis_change
            if abs(quantity[security]) < 1e-10:
                quantity[security] = 0.0
                cost_sar[security] = 0.0
        cash[currency] += cash_change_original

        row.update({
            "signed_quantity": signed_qty,
            "quantity_before": quantity_before,
            "quantity_after": quantity[security] if security else 0.0,
            "cost_basis_before_sar": cost_before,
            "cost_basis_change_sar": basis_change,
            "cost_basis_after_sar": cost_sar[security] if security else 0.0,
            "realized_pl_sar": realized,
            "cash_change_original": cash_change_original,
            "cash_change_sar": cash_change_original * fx,
        })
        applied.append(row)

    positions = {}
    for security in sorted(quantity):
        qty = quantity[security]
        if abs(qty) < 1e-10:
            continue
        positions[security] = {
            "quantity": qty,
            "cost_basis_sar": cost_sar[security],
            "average_cost_sar": cost_sar[security] / qty,
        }
    return {"rows": applied, "positions": positions, "cash": dict(cash)}
