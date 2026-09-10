"""Build a price-history payload for the Musaffa-compliant US universe.

This is deliberately a market-data importer, not a Sharia adjudicator. The
classification JSON remains the source of the US universe and each generated
record keeps the raw classification ticker, company name and source status.
Fundamental fields are left missing when the source does not provide them;
missing data must remain visible in the workbook rather than becoming a fake
neutral score.
"""

from __future__ import annotations

import argparse
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime, timezone
import io
import json
import math
import os
from pathlib import Path
import time

import pandas as pd
import yfinance as yf

import metrics


def _number(value):
    try:
        value = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(value):
        return None
    return round(value, 8)


def _iso_date(value):
    try:
        return value.date().isoformat()
    except AttributeError:
        return str(value)[:10]


def _iso_datetime(value):
    try:
        return value.isoformat()
    except AttributeError:
        return str(value)


def load_classification(path):
    """Return unique US rows with the source status COMPLIANT/حلال."""
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    rows = payload.get("rows", payload) if isinstance(payload, dict) else payload
    out = []
    seen = set()
    for row in rows or []:
        if isinstance(row, dict):
            market = row.get("market") or row.get("السوق")
            ticker = row.get("ticker") or row.get("symbol") or row.get("الرمز")
            company = row.get("company") or row.get("name") or row.get("اسم الشركة")
            status = row.get("classification") or row.get("status") or row.get("التصنيف الشرعي")
            source = row.get("source") or row.get("المصدر")
            source_status = row.get("source_status") or row.get("original_status")
            notes = row.get("notes") or row.get("ملاحظات")
        else:
            market, ticker, company, status, source, source_status, notes = (list(row) + [""] * 7)[:7]
        ticker = str(ticker or "").strip()
        if market != "السوق الأمريكي" or not ticker:
            continue
        if status != "حلال" and source_status != "COMPLIANT":
            continue
        if ticker in seen:
            continue
        seen.add(ticker)
        out.append({
            "market": market,
            "ticker": ticker,
            "company": str(company or "").strip(),
            "classification": str(status or "").strip(),
            "source": str(source or "Musaffa").strip(),
            "source_status": str(source_status or "COMPLIANT").strip(),
            "notes": str(notes or "").strip(),
        })
    return out


def _columns(frame, field, ticker):
    if field not in frame:
        return pd.Series(dtype="float64")
    series = frame[field]
    if isinstance(series, pd.DataFrame):
        if ticker in series.columns:
            return series[ticker]
        return pd.Series(dtype="float64")
    return series


def _daily_returns(values):
    return [cur / prev - 1.0 for prev, cur in zip(values, values[1:]) if prev]


def _trailing_return(series, days):
    if series.empty:
        return None
    target = series.index[-1] - pd.Timedelta(days=days)
    prior = series.loc[:target]
    if prior.empty:
        return None
    base = _number(prior.iloc[-1])
    last = _number(series.iloc[-1])
    if base in (None, 0) or last is None:
        return None
    return _number(last / base - 1.0)


def _annual_rows(ticker, closes, dividends):
    rows = []
    try:
        years = sorted({int(year) for year in closes.index.year})
    except Exception:
        return rows
    for year in years:
        year_closes = closes[closes.index.year == year]
        values = [_number(x) for x in year_closes.tolist()]
        values = [x for x in values if x is not None]
        if not values:
            continue
        year_dividends = dividends[dividends.index.year == year] if not dividends.empty else pd.Series(dtype="float64")
        divs = _number(year_dividends.sum()) if not year_dividends.empty else 0.0
        div_yield = _number(divs / values[0] * 100.0) if values[0] else None
        through_year = closes[closes.index.year <= year]
        momentum = metrics.momentum_12_1([float(x) for x in through_year.tolist()])
        rows.append({
            "symbol": ticker,
            "year": year,
            "close": _number(values[-1]),
            "ret": _number(metrics.annual_return(values)),
            "vol": _number(metrics.annualized_vol(_daily_returns(values))),
            "maxdd": _number(metrics.max_drawdown(values)),
            "divs": divs,
            "div_yield": div_yield,
            "momentum": _number(momentum),
            "eps": None,
            "price_date": _iso_date(year_closes.index[-1]),
            "return_basis": "total return from auto-adjusted Close",
            "dividend_unit": "USD per share",
            "source": "Yahoo Finance",
        })
    return rows


def build_payload(classification_path, period="5y", chunk_size=75, pause=0.25):
    classification = load_classification(classification_path)
    tickers = [row["ticker"] for row in classification]
    by_ticker = {row["ticker"]: row for row in classification}
    records = {}
    failures = []
    started = datetime.now(timezone.utc).replace(microsecond=0).isoformat()

    for offset in range(0, len(tickers), chunk_size):
        chunk = tickers[offset:offset + chunk_size]
        # yfinance can emit one warning per stale/foreign/when-issued symbol;
        # keep the generated payload readable while retaining failures below.
        sink = io.StringIO()
        with redirect_stdout(sink), redirect_stderr(sink):
            try:
                frame = yf.download(
                    tickers=chunk,
                    period=period,
                    auto_adjust=True,
                    actions=True,
                    threads=True,
                    progress=False,
                    group_by="column",
                )
            except Exception as exc:  # keep other chunks usable
                frame = pd.DataFrame()
                failures.extend({"ticker": ticker, "error": str(exc)} for ticker in chunk)
        if isinstance(frame, pd.DataFrame) and not frame.empty and "Close" in frame:
            for ticker in chunk:
                closes = _columns(frame, "Close", ticker).dropna()
                if len(closes) < 60:
                    failures.append({"ticker": ticker, "error": "fewer than 60 usable closes"})
                    continue
                dividends = _columns(frame, "Dividends", ticker).fillna(0).dropna()
                close_values = [_number(x) for x in closes.tolist()]
                close_values = [x for x in close_values if x is not None]
                vol = metrics.volatility_state(close_values) or {}
                maxdd_values = close_values[-min(len(close_values), 504):]
                sma = metrics.sma200_flag(close_values)
                momentum = metrics.momentum_12_1(close_values)
                maxdd_2y = metrics.max_drawdown(maxdd_values)
                score = metrics.composite_score_with_coverage(
                    None, None, None, (sma, momentum), maxdd_2y
                )
                last_index = closes.index[-1]
                row = by_ticker[ticker]
                records[ticker] = {
                    **row,
                    "security_id": "US-" + ticker,
                    "exchange": "NYSE/NASDAQ pending",
                    "currency": "USD",
                    "yahoo_symbol": ticker,
                    "price": _number(closes.iloc[-1]),
                    "price_as_of": _iso_date(last_index),
                    "price_observed_at": _iso_datetime(last_index),
                    "source": "Yahoo Finance",
                    "source_url": "https://finance.yahoo.com/quote/" + ticker,
                    "price_basis": "auto-adjusted Close (splits and cash dividends)",
                    "return_basis": "total return from auto-adjusted Close",
                    "ret_1w": _trailing_return(closes, 7),
                    "ret_1m": _trailing_return(closes, 30),
                    "ret_3m": _trailing_return(closes, 91),
                    "ret_6m": _trailing_return(closes, 182),
                    "ret_1y": _trailing_return(closes, 365),
                    "rsi14": _number(metrics.rsi14(close_values)),
                    "vol_level": vol.get("level"),
                    "vol_trend": vol.get("trend"),
                    "vol_short": _number(vol.get("short_vol")),
                    "vol_long": _number(vol.get("long_vol")),
                    "sma200_flag": sma,
                    "maxdd_2y": _number(maxdd_2y),
                    "score": score["score"],
                    "completeness": score["completeness"],
                    "actionable": score["actionable"],
                    "annual_rows": _annual_rows(ticker, closes, dividends),
                    "statement_rows": [],
                }
        if pause:
            time.sleep(pause)

    return {
        "generated_at": started,
        "source": "Yahoo Finance",
        "period": period,
        "classification_source": str(classification_path),
        "records": [records[ticker] for ticker in tickers if ticker in records],
        "unresolved": [by_ticker[ticker] | {"error": "no usable Yahoo price history"}
                       for ticker in tickers if ticker not in records],
        "summary": {
            "classification_count": len(classification),
            "price_history_count": len(records),
            "unresolved_count": len(tickers) - len(records),
        },
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--classification-json", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--period", default="5y")
    parser.add_argument("--chunk-size", type=int, default=75)
    parser.add_argument("--pause", type=float, default=0.25)
    args = parser.parse_args()
    payload = build_payload(
        args.classification_json,
        period=args.period,
        chunk_size=args.chunk_size,
        pause=args.pause,
    )
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    Path(args.out).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload["summary"], ensure_ascii=False))


if __name__ == "__main__":
    main()
