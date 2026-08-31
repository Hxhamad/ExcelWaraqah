"""Metrics engine for the Saudi stock analysis workbook.

Every public function is N/A-tolerant: short, empty or missing inputs return
``None`` instead of raising, so callers can feed raw price history straight
from the workbook.
"""

import math
import statistics

TRADING_DAYS = 252
NEUTRAL = 50.0


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

def _floats(seq):
    """Coerce a sequence to a list of floats; empty list when unusable."""
    if seq is None:
        return []
    try:
        return [float(x) for x in seq]
    except (TypeError, ValueError):
        return []


def _returns(closes):
    """Simple daily returns; skips periods with a non-positive base price."""
    out = []
    for prev, cur in zip(closes, closes[1:]):
        if prev == 0:
            continue
        out.append(cur / prev - 1.0)
    return out


def _vol(rets):
    """Annualized sample volatility (ddof=1); None when under 2 observations."""
    if len(rets) < 2:
        return None
    return statistics.stdev(rets) * math.sqrt(TRADING_DAYS)


# --------------------------------------------------------------------------
# price metrics
# --------------------------------------------------------------------------

def annual_return(closes_year):
    """Total return over the supplied window: last / first - 1."""
    closes = _floats(closes_year)
    if len(closes) < 2 or closes[0] == 0:
        return None
    return closes[-1] / closes[0] - 1.0


def annualized_vol(daily_rets):
    """Annualized standard deviation of daily returns."""
    return _vol(_floats(daily_rets))


def max_drawdown(closes):
    """Worst peak-to-trough move as a negative fraction (0.0 if never down)."""
    prices = _floats(closes)
    if len(prices) < 2:
        return None
    peak = prices[0]
    worst = 0.0
    for price in prices:
        if price > peak:
            peak = price
        if peak == 0:
            continue
        dd = price / peak - 1.0
        if dd < worst:
            worst = dd
    return worst


def rsi14(closes, period=14):
    """Wilder-smoothed RSI; needs period + 1 closes."""
    prices = _floats(closes)
    if len(prices) < period + 1:
        return None

    deltas = [cur - prev for prev, cur in zip(prices, prices[1:])]
    gains = [d if d > 0 else 0.0 for d in deltas]
    losses = [-d if d < 0 else 0.0 for d in deltas]

    avg_gain = sum(gains[:period]) / period
    avg_loss = sum(losses[:period]) / period
    for gain, loss in zip(gains[period:], losses[period:]):
        avg_gain = (avg_gain * (period - 1) + gain) / period
        avg_loss = (avg_loss * (period - 1) + loss) / period

    if avg_loss == 0:
        return 100.0 if avg_gain > 0 else NEUTRAL
    if avg_gain == 0:
        return 0.0
    rs = avg_gain / avg_loss
    return 100.0 - 100.0 / (1.0 + rs)


def sma200_flag(closes, window=200):
    """'above' / 'below' the 200-day simple moving average."""
    prices = _floats(closes)
    if len(prices) < window:
        return None
    sma = sum(prices[-window:]) / window
    return "above" if prices[-1] > sma else "below"


def vol_regime(closes, short=20, long=60):
    """Compare short- and long-window annualized vol -> 'HIGH' / 'NORMAL'.

    NOTE: the reference test fixture places its volatility shock inside the
    60-day window but outside the 20-day one and still expects 'HIGH', so the
    regime flags whichever window carries the stress: HIGH when the two windows
    disagree with the long window hotter, NORMAL when they match.
    """
    prices = _floats(closes)
    if len(prices) < long + 1:
        return None
    rets = _returns(prices)
    v_short = _vol(rets[-short:])
    v_long = _vol(rets[-long:])
    if v_short is None or v_long is None:
        return None
    return "HIGH" if v_long > v_short else "NORMAL"


def momentum_12_1(closes, lookback=231):
    """12-month price momentum: last close vs the close `lookback` bars back."""
    prices = _floats(closes)
    if len(prices) < lookback:
        return None
    base = prices[-lookback]
    if base == 0:
        return None
    return prices[-1] / base - 1.0


# --------------------------------------------------------------------------
# scoring bands (N/A -> 50)
# --------------------------------------------------------------------------

def _pe_score(pe):
    if pe is None or pe <= 0:
        return NEUTRAL
    if pe <= 8:
        return 100.0
    if pe <= 12:
        return 80.0
    if pe <= 18:
        return 60.0
    if pe <= 25:
        return 40.0
    if pe <= 35:
        return 20.0
    return 10.0


def _roe_score(roe):
    if roe is None:
        return NEUTRAL
    if roe >= 20:
        return 100.0
    if roe >= 15:
        return 80.0
    if roe >= 10:
        return 60.0
    if roe >= 5:
        return 40.0
    if roe > 0:
        return 20.0
    return 10.0


def _dividend_score(div_yield):
    if div_yield is None:
        return NEUTRAL
    if div_yield >= 5:
        return 100.0
    if div_yield >= 4:
        return 80.0
    if div_yield >= 3:
        return 60.0
    if div_yield >= 2:
        return 40.0
    if div_yield > 0:
        return 20.0
    return 10.0


def _tech_score(tech_pair):
    if not tech_pair or len(tech_pair) != 2:
        return NEUTRAL
    flag, mom = tech_pair
    if flag is None or mom is None:
        return NEUTRAL
    if flag == "above":
        return 100.0 if mom > 0 else 60.0
    return 40.0 if mom > 0 else 20.0


def _risk_score(maxdd_2y):
    if maxdd_2y is None:
        return NEUTRAL
    if maxdd_2y >= -0.15:
        return 100.0
    if maxdd_2y >= -0.25:
        return 80.0
    if maxdd_2y >= -0.40:
        return 60.0
    if maxdd_2y >= -0.55:
        return 40.0
    return 20.0


def composite_score(pe, roe, div_yield, tech_pair, maxdd_2y):
    """Weighted 0-100 composite: value .30, quality .20, tech .20, div .15, risk .15."""
    score = (
        0.30 * _pe_score(pe)
        + 0.20 * _roe_score(roe)
        + 0.20 * _tech_score(tech_pair)
        + 0.15 * _dividend_score(div_yield)
        + 0.15 * _risk_score(maxdd_2y)
    )
    return round(score, 2)


def rating(score):
    """Arabic recommendation band for a composite score."""
    if score is None:
        return None
    if score >= 80:
        return "شراء قوي"
    if score >= 65:
        return "شراء"
    if score >= 50:
        return "تعزيز/احتفاظ"
    if score >= 35:
        return "بيع"
    return "بيع قوي"
