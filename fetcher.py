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
from datetime import datetime, timedelta, timezone

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from metrics import (  # noqa: E402
    annual_return,
    annualized_vol,
    max_drawdown,
    momentum_12_1,
    rsi14,
    sma200_flag,
    volatility_state,
    vol_regime,
)

FIRST_YEAR = 2015

ANNUAL_FIELDS = [
    "symbol", "year", "close", "ret", "vol", "maxdd",
    "divs", "div_yield", "momentum", "eps", "price_date",
    "return_basis", "dividend_unit", "source",
]
STATEMENT_FIELDS = [
    "symbol", "year", "period_end", "revenue", "net_income", "eps", "roe", "de", "payout",
]

MARKET_DEFAULTS = {
    "Tadawul": {
        "currency": "SAR", "calendar": "Saudi Exchange", "settlement": "T+2",
        "timezone": "Asia/Riyadh", "session": "Sunday–Thursday; core 10:00–15:00",
        "dst": "No daylight-saving shift", "quantity_rule": "Whole shares; minimum 1",
        "order_rule": "Broker/account capabilities must be confirmed before approval",
        "rule_source": "https://www.saudiexchange.sa/wps/portal/saudiexchange/trading/market-services/equities?locale=en",
    },
    "NYSE": {
        "currency": "USD", "calendar": "NYSE", "settlement": "T+1",
        "timezone": "America/New_York", "session": "Core session 09:30–16:00 ET",
        "dst": "America/New_York daylight-saving rules",
        "quantity_rule": "Whole/fractional shares depend on broker and security",
        "order_rule": "Broker/account capabilities must be confirmed before approval",
        "rule_source": "https://www.sec.gov/rules-regulations/2023/02/34-96930",
    },
    "NASDAQ": {
        "currency": "USD", "calendar": "NASDAQ", "settlement": "T+1",
        "timezone": "America/New_York", "session": "Core session 09:30–16:00 ET",
        "dst": "America/New_York daylight-saving rules",
        "quantity_rule": "Whole/fractional shares depend on broker and security",
        "order_rule": "Broker/account capabilities must be confirmed before approval",
        "rule_source": "https://www.sec.gov/rules-regulations/2023/02/34-96930",
    },
}

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

def normalize_security(raw):
    """Normalize legacy Saudi codes or a unified security mapping."""
    if isinstance(raw, dict):
        item = dict(raw)
        ticker = str(item.get("ticker") or item.get("code") or "").strip().upper()
        exchange_raw = str(item.get("exchange") or "").strip()
        exchange = {
            "TADAWUL": "Tadawul", "SAUDI EXCHANGE": "Tadawul",
            "NYSE": "NYSE", "NASDAQ": "NASDAQ",
        }.get(exchange_raw.upper(), exchange_raw) or (
            "Tadawul" if ticker.isdigit() and len(ticker) == 4 else "NASDAQ")
        currency = str(item.get("currency") or
                       MARKET_DEFAULTS.get(exchange, {}).get("currency") or "USD").upper()
        yahoo_symbol = str(item.get("yahoo_symbol") or
                           ((ticker + ".SR") if exchange == "Tadawul" else ticker)).strip()
        security_id = str(item.get("security_id") or
                          (("SA-" + ticker) if exchange == "Tadawul" else ("US-" + ticker))).strip()
    else:
        ticker = str(raw).strip().upper()
        exchange = "Tadawul" if ticker.isdigit() and len(ticker) == 4 else "NASDAQ"
        currency = "SAR" if exchange == "Tadawul" else "USD"
        yahoo_symbol = ticker + ".SR" if exchange == "Tadawul" else ticker
        security_id = "SA-" + ticker if exchange == "Tadawul" else "US-" + ticker
    defaults = MARKET_DEFAULTS.get(exchange, {})
    return {
        "security_id": security_id,
        "ticker": ticker,
        "code": ticker,
        "exchange": exchange,
        "currency": currency,
        "yahoo_symbol": yahoo_symbol,
        "calendar": defaults.get("calendar", "Pending confirmation"),
        "settlement": defaults.get("settlement", "Pending confirmation"),
        "timezone": defaults.get("timezone", "Pending confirmation"),
        "session": defaults.get("session", "Pending confirmation"),
        "dst": defaults.get("dst", "Pending confirmation"),
        "quantity_rule": defaults.get("quantity_rule", "Pending broker confirmation"),
        "order_rule": defaults.get("order_rule", "Pending broker confirmation"),
        "rule_source": defaults.get("rule_source"),
        "instrument_type": str(item.get("instrument_type") or "Equity")
        if isinstance(raw, dict) else "Equity",
    }


def fetch_security(raw):
    """Fetch one unified Saudi or US security using an explicit identifier map."""
    import yfinance as yf

    security = normalize_security(raw)
    code = security["ticker"]
    ticker_id = security["yahoo_symbol"]

    try:
        ticker = yf.Ticker(ticker_id)
    except Exception:
        return None

    # --- prices (the one hard requirement) --------------------------------
    try:
        # Explicitly use the total-return series.  yfinance's default has
        # changed historically, so never depend on an implicit setting.
        hist = ticker.history(period="max", auto_adjust=True, actions=True)
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

    vol = _safe(volatility_state, close_list) or {}
    fetched_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    try:
        price_as_of = last_date.date().isoformat()
    except Exception:
        price_as_of = str(last_date)
    try:
        price_observed_at = last_date.isoformat()
    except Exception:
        price_observed_at = str(last_date)
    snapshot = {
        "code": code,
        "security_id": security["security_id"],
        "ticker": security["ticker"],
        "exchange": security["exchange"],
        "currency": info.get("currency") or security["currency"],
        "yahoo_symbol": ticker_id,
        "calendar": security["calendar"],
        "settlement": security["settlement"],
        "market_timezone": security["timezone"],
        "market_session": security["session"],
        "dst_handling": security["dst"],
        "quantity_rule": security["quantity_rule"],
        "order_rule": security["order_rule"],
        "market_rule_source": security["rule_source"],
        "instrument_type": info.get("quoteType") or security["instrument_type"],
        "name_en": info.get("longName") or info.get("shortName") or None,
        "sector": info.get("sector") or None,
        "price": _round(closes.iloc[-1], 4),
        "price_as_of": price_as_of,
        "price_observed_at": price_observed_at,
        "fetched_at": fetched_at,
        "source": "Yahoo Finance",
        "source_url": "https://finance.yahoo.com/quote/%s" % ticker_id,
        "price_basis": "auto-adjusted Close (splits and cash dividends)",
        "return_basis": "total return from auto-adjusted Close",
        "dividend_unit": "%s per share" % (info.get("currency") or security["currency"]),
        "fundamentals_basis": "Annual statements; trailing provider ratios where labelled",
        "publication_date": None,
        "ret_1w": _trailing_return(closes, 7),
        "ret_1m": _trailing_return(closes, 30),
        "ret_3m": _trailing_return(closes, 91),
        "ret_6m": _trailing_return(closes, 182),
        "ret_1y": _trailing_return(closes, 365),
        "rsi14": _round(_safe(rsi14, close_list), 2),
        "vol_regime": vol.get("level") or _safe(vol_regime, close_list),
        "vol_level": vol.get("level"),
        "vol_trend": vol.get("trend"),
        "vol_short": _round(vol.get("short_vol")),
        "vol_long": _round(vol.get("long_vol")),
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


def fetch_one(code):
    """Backward-compatible Tadawul entry point for a four-digit code."""
    return fetch_security({
        "ticker": str(code).strip(), "exchange": "Tadawul", "currency": "SAR"
    })


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
            "period_end": getattr(col, "date", lambda: col)().isoformat()
            if hasattr(col, "date") else str(col),
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
            "price_date": _year_slice(closes, year).index[-1].date().isoformat(),
            "return_basis": "total return from auto-adjusted Close",
            "dividend_unit": "cash per share in listing currency",
            "source": "Yahoo Finance",
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


def fetch_all(codes, sleep_s=1.0, data_dir="data", force=False):
    """Fetch each code, merging results into the data_dir artefacts.

    Codes already present in annual_metrics.csv are skipped unless ``force`` is
    true, so an interrupted run can resume while a deliberate evidence refresh
    can replace stale records. Returns {code: snapshot} for every code handled.
    """
    os.makedirs(data_dir, exist_ok=True)
    annual_path, statements_path, snapshot_path = _paths(data_dir)

    annual = _read_csv(annual_path, ANNUAL_FIELDS)
    statements = _read_csv(statements_path, STATEMENT_FIELDS)
    snapshots = _read_snapshot(snapshot_path)

    cached = set(annual["symbol"]) if len(annual) else set()
    results = {}

    for index, raw in enumerate(codes):
        security = normalize_security(raw)
        code = security["ticker"]
        if code in cached and not force:
            print("fetched %s skipped (cached)" % code)
            if code in snapshots:
                results[code] = snapshots[code]
            continue

        if index and sleep_s:
            time.sleep(sleep_s)

        try:
            record = fetch_security(security)
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
