"""yfinance data fetcher for the Saudi stock analysis workbook.

Everything here is defensive: Yahoo drops fields without warning (REITs have no
PE, some large caps return an empty income statement), so every access is
wrapped and any value that cannot be computed becomes ``None`` rather than an
exception. Only a ticker with no price history at all is treated as a failure.

Units, fixed once so the scoring bands in metrics.py line up:
  * returns / drawdowns / momentum -> fractions (0.12 == +12%)
  * roe and div_yield              -> percent (23.7 == 23.7%)
  * payout                         -> fraction (0.79 == 79%)
"""

import json
import os
import sys
import time
from datetime import timedelta

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from metrics import (  # noqa: E402
    annual_return,
    annualized_vol,
    max_drawdown,
    momentum_12_1,
    rsi14,
    sma200_flag,
    vol_regime,
)

FIRST_YEAR = 2015

ANNUAL_FIELDS = [
    "symbol", "year", "close", "ret", "vol", "maxdd",
    "divs", "div_yield", "momentum", "eps",
]
STATEMENT_FIELDS = [
    "symbol", "year", "revenue", "net_income", "eps", "roe", "de", "payout",
]

# Yahoo renames statement lines between sectors; try each label in order.
REVENUE_KEYS = ("Total Revenue", "Operating Revenue")
NET_INCOME_KEYS = (
    "Net Income",
    "Net Income Common Stockholders",
    "Net Income Continuous Operations",
    "Net Income Including Noncontrolling Interests",
)
EPS_KEYS = ("Basic EPS", "Diluted EPS")
SHARES_KEYS = ("Basic Average Shares", "Diluted Average Shares")
EQUITY_KEYS = (
    "Stockholders Equity",
    "Common Stock Equity",
    "Total Equity Gross Minority Interest",
)
DEBT_KEYS = ("Total Debt",)
SHARE_COUNT_KEYS = ("Ordinary Shares Number", "Share Issued")
DIVIDEND_PAID_KEYS = (
    "Cash Dividends Paid",
    "Common Stock Dividend Paid",
    "Dividends Paid",
    "Payments For Dividends",
)


# --------------------------------------------------------------------------
# tiny guards
# --------------------------------------------------------------------------

def _num(value):
    """Float or None -- also swallows NaN, which pandas hands back constantly."""
    if value is None:
        return None
    try:
        out = float(value)
    except (TypeError, ValueError):
        return None
    if out != out or out in (float("inf"), float("-inf")):
        return None
    return out


def _round(value, digits=6):
    value = _num(value)
    return None if value is None else round(value, digits)


def _cell(frame, keys, column):
    """First value found for any of `keys` in `column` of a statement frame."""
    if frame is None or column is None:
        return None
    try:
        if getattr(frame, "empty", True):
            return None
        for key in keys:
            if key in frame.index:
                return _num(frame.at[key, column])
    except Exception:
        return None
    return None


def _column_for_year(frame, year):
    """Statement column whose period ends in `year`, if any."""
    if frame is None:
        return None
    try:
        for col in frame.columns:
            if getattr(col, "year", None) == year:
                return col
    except Exception:
        return None
    return None


def _safe(fn, *args, **kwargs):
    """Call a metrics helper without ever propagating a failure."""
    try:
        return fn(*args, **kwargs)
    except Exception:
        return None


# --------------------------------------------------------------------------
# price helpers
# --------------------------------------------------------------------------

def _close_days_back(closes, days):
    """Close at or just before `days` calendar days ago; None when too short."""
    try:
        target = closes.index[-1] - timedelta(days=days)
        window = closes.loc[:target]
        if len(window) == 0:
            return None
        return _num(window.iloc[-1])
    except Exception:
        return None


def _trailing_return(closes, days):
    base = _close_days_back(closes, days)
    last = _num(closes.iloc[-1]) if len(closes) else None
    if base is None or base == 0 or last is None:
        return None
    return round(last / base - 1.0, 6)


def _year_slice(closes, year):
    try:
        return closes[closes.index.year == year]
    except Exception:
        return closes.iloc[0:0]


def _daily_returns(values):
    out = []
    for prev, cur in zip(values, values[1:]):
        if prev == 0:
            continue
        out.append(cur / prev - 1.0)
    return out


# --------------------------------------------------------------------------
# single ticker
# --------------------------------------------------------------------------

def fetch_one(code):
    """Fetch every workbook input for one 4-digit Tadawul code.

    Returns None only when the ticker has no usable price history; otherwise a
    dict whose unavailable fields are None.
    """
    import yfinance as yf

    code = str(code).strip()
    ticker_id = "%s.SR" % code

    try:
        ticker = yf.Ticker(ticker_id)
    except Exception:
        return None

    # --- prices (the one hard requirement) --------------------------------
    try:
        hist = ticker.history(period="max")
    except Exception:
        hist = None
    if hist is None or getattr(hist, "empty", True) or "Close" not in hist.columns:
        return None

    closes = hist["Close"].dropna()
    if len(closes) == 0:
        return None

    # --- everything else is best-effort -----------------------------------
    try:
        info = ticker.info or {}
    except Exception:
        info = {}
    try:
        divs = ticker.dividends
        if divs is None:
            divs = pd.Series(dtype="float64")
    except Exception:
        divs = pd.Series(dtype="float64")
    try:
        income = ticker.income_stmt
    except Exception:
        income = None
    try:
        balance = ticker.balance_sheet
    except Exception:
        balance = None
    try:
        cashflow = ticker.cashflow
    except Exception:
        cashflow = None

    close_list = [float(x) for x in closes.tolist()]
    last_date = closes.index[-1]

    two_years = _safe(lambda: closes.loc[last_date - timedelta(days=730):])
    dd_values = [float(x) for x in two_years.tolist()] if two_years is not None else []

    snapshot = {
        "code": code,
        "name_en": info.get("longName") or info.get("shortName") or None,
        "sector": info.get("sector") or None,
        "price": _round(closes.iloc[-1], 4),
        "ret_1w": _trailing_return(closes, 7),
        "ret_1m": _trailing_return(closes, 30),
        "ret_3m": _trailing_return(closes, 91),
        "ret_6m": _trailing_return(closes, 182),
        "ret_1y": _trailing_return(closes, 365),
        "rsi14": _round(_safe(rsi14, close_list), 2),
        "vol_regime": _safe(vol_regime, close_list),
        "sma200_flag": _safe(sma200_flag, close_list),
        "maxdd_2y": _round(_safe(max_drawdown, dd_values)),
    }

    roe = _num(info.get("returnOnEquity"))
    snapshot["info"] = {
        "pe": _round(info.get("trailingPE"), 4),
        "roe": None if roe is None else round(roe * 100.0, 4),     # percent
        "payout": _round(info.get("payoutRatio"), 4),              # fraction
        "div5y": _round(info.get("fiveYearAvgDividendYield"), 4),  # percent
    }

    snapshot["statement_rows"] = _statement_rows(code, income, balance, cashflow)
    snapshot["annual_rows"] = _annual_rows(
        code, closes, divs, snapshot["statement_rows"]
    )
    return snapshot


def _statement_rows(code, income, balance, cashflow):
    """One row per annual statement column Yahoo returns (2021..2025)."""
    rows = []
    try:
        columns = list(income.columns) if income is not None and not income.empty else []
    except Exception:
        columns = []

    for col in columns:
        year = getattr(col, "year", None)
        if year is None:
            continue

        revenue = _cell(income, REVENUE_KEYS, col)
        net_income = _cell(income, NET_INCOME_KEYS, col)

        eps = _cell(income, EPS_KEYS, col)
        if eps is None:
            shares = _cell(income, SHARES_KEYS, col)
            if shares is None or shares == 0:
                shares = _cell(balance, SHARE_COUNT_KEYS, _column_for_year(balance, year))
            if net_income is not None and shares is not None and shares != 0:
                eps = net_income / shares

        bs_col = _column_for_year(balance, year)
        equity = _cell(balance, EQUITY_KEYS, bs_col)
        debt = _cell(balance, DEBT_KEYS, bs_col)

        roe = None
        if net_income is not None and equity is not None and equity != 0:
            roe = net_income / equity * 100.0
        de = None
        if debt is not None and equity is not None and equity != 0:
            de = debt / equity

        paid = _cell(cashflow, DIVIDEND_PAID_KEYS, _column_for_year(cashflow, year))
        payout = None
        if paid is not None and net_income is not None and net_income != 0:
            payout = abs(paid) / net_income

        rows.append({
            "symbol": code,
            "year": int(year),
            "revenue": _round(revenue, 2),
            "net_income": _round(net_income, 2),
            "eps": _round(eps, 4),
            "roe": _round(roe, 4),
            "de": _round(de, 4),
            "payout": _round(payout, 4),
        })

    rows.sort(key=lambda r: r["year"])
    return rows


def _annual_rows(code, closes, divs, statement_rows):
    """Calendar-year metrics from FIRST_YEAR to the last year with prices."""
    try:
        years = sorted({int(y) for y in closes.index.year})
    except Exception:
        return []
    years = [y for y in years if y >= FIRST_YEAR]

    eps_by_year = {r["year"]: r["eps"] for r in statement_rows}
    rows = []

    for year in years:
        values = [float(x) for x in _year_slice(closes, year).tolist()]
        if not values:
            continue

        div_sum = None
        try:
            if divs is not None and len(divs) > 0:
                in_year = divs[divs.index.year == year]
                div_sum = float(in_year.sum()) if len(in_year) else 0.0
        except Exception:
            div_sum = None

        div_yield = None
        if div_sum is not None and values[0] != 0:
            div_yield = div_sum / values[0] * 100.0

        # momentum as it stood on the last bar of that year
        try:
            through_year = closes[closes.index.year <= year]
            momentum = _safe(momentum_12_1, [float(x) for x in through_year.tolist()])
        except Exception:
            momentum = None

        rows.append({
            "symbol": code,
            "year": year,
            "close": _round(values[-1], 4),
            "ret": _round(_safe(annual_return, values)),
            "vol": _round(_safe(annualized_vol, _daily_returns(values))),
            "maxdd": _round(_safe(max_drawdown, values)),
            "divs": _round(div_sum, 4),
            "div_yield": _round(div_yield, 4),
            "momentum": _round(momentum),
            "eps": eps_by_year.get(year),
        })
    return rows


# --------------------------------------------------------------------------
# persistence
# --------------------------------------------------------------------------

def _paths(data_dir):
    return (
        os.path.join(data_dir, "annual_metrics.csv"),
        os.path.join(data_dir, "statements.csv"),
        os.path.join(data_dir, "snapshot.json"),
    )


def _read_csv(path, fields):
    if not os.path.exists(path):
        return pd.DataFrame(columns=fields)
    try:
        frame = pd.read_csv(path, dtype={"symbol": str})
    except Exception:
        return pd.DataFrame(columns=fields)
    if "symbol" in frame.columns:
        frame["symbol"] = frame["symbol"].astype(str).str.zfill(4)
    for field in fields:
        if field not in frame.columns:
            frame[field] = None
    return frame[fields]


def _read_snapshot(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write_snapshot(path, data):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2, sort_keys=True)


def _merge_rows(existing, rows, fields):
    """Replace the touched symbols' rows, keep everyone else's."""
    if not rows:
        return existing
    new = pd.DataFrame(rows, columns=fields)
    new["symbol"] = new["symbol"].astype(str)
    touched = set(new["symbol"])
    kept = existing[~existing["symbol"].isin(touched)] if len(existing) else existing
    merged = pd.concat([kept, new], ignore_index=True)
    return merged.sort_values(["symbol", "year"]).reset_index(drop=True)


def _snapshot_only(record):
    """The snapshot view of a fetch_one result (no per-year tables)."""
    return {k: v for k, v in record.items()
            if k not in ("annual_rows", "statement_rows")}


def fetch_all(codes, sleep_s=1.0, data_dir="data"):
    """Fetch each code, merging results into the data_dir artefacts.

    Codes already present in annual_metrics.csv are skipped so an interrupted
    run can simply be re-run. Returns {code: snapshot} for every code handled,
    cached ones included.
    """
    os.makedirs(data_dir, exist_ok=True)
    annual_path, statements_path, snapshot_path = _paths(data_dir)

    annual = _read_csv(annual_path, ANNUAL_FIELDS)
    statements = _read_csv(statements_path, STATEMENT_FIELDS)
    snapshots = _read_snapshot(snapshot_path)

    cached = set(annual["symbol"]) if len(annual) else set()
    results = {}

    for index, raw in enumerate(codes):
        code = str(raw).strip()
        if code in cached:
            print("fetched %s skipped (cached)" % code)
            if code in snapshots:
                results[code] = snapshots[code]
            continue

        if index and sleep_s:
            time.sleep(sleep_s)

        try:
            record = fetch_one(code)
        except Exception as exc:  # fetch_one guards internally; belt and braces
            print("FAIL %s error (%s)" % (code, exc))
            continue

        if record is None:
            print("FAIL %s no data" % code)
            continue

        annual = _merge_rows(annual, record["annual_rows"], ANNUAL_FIELDS)
        statements = _merge_rows(statements, record["statement_rows"], STATEMENT_FIELDS)
        snapshots[code] = _snapshot_only(record)
        results[code] = snapshots[code]

        annual.to_csv(annual_path, index=False)
        statements.to_csv(statements_path, index=False)
        _write_snapshot(snapshot_path, snapshots)
        cached.add(code)
        print("fetched %s ok" % code)

    annual.to_csv(annual_path, index=False)
    statements.to_csv(statements_path, index=False)
    _write_snapshot(snapshot_path, snapshots)
    return results


def fetch_quick(codes, data_dir="data"):
    """Refresh only the snapshot entries for `codes`; the CSVs are left alone."""
    os.makedirs(data_dir, exist_ok=True)
    _, _, snapshot_path = _paths(data_dir)
    snapshots = _read_snapshot(snapshot_path)

    results = {}
    for raw in codes:
        code = str(raw).strip()
        try:
            record = fetch_one(code)
        except Exception as exc:
            print("FAIL %s error (%s)" % (code, exc))
            continue
        if record is None:
            print("FAIL %s no data" % code)
            continue
        snapshots[code] = _snapshot_only(record)
        results[code] = snapshots[code]
        print("fetched %s ok" % code)

    _write_snapshot(snapshot_path, snapshots)
    return results
