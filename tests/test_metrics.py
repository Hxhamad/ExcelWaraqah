import math, pytest
from metrics import (annual_return, annualized_vol, max_drawdown, rsi14,
                     sma200_flag, volatility_state, vol_regime, momentum_12_1,
                     composite_score, composite_score_with_coverage, rating)

def test_annual_return():
    assert annual_return([100.0] * 252) == pytest.approx(0.0)
    assert annual_return([100.0] + [200.0] * 251) == pytest.approx(1.0)

def test_max_drawdown():
    path = [100, 90, 80, 70, 60, 50, 60, 70, 80, 90, 100]
    assert max_drawdown(path) == pytest.approx(-0.5)
    assert max_drawdown([100.0] * 10) == pytest.approx(0.0)
    assert max_drawdown([100, 120, 90]) == pytest.approx(-0.25)

def test_annualized_vol():
    rets = [0.01, -0.01, 0.02, -0.02]
    expected = (math.sqrt((0.0001 + 0.0001 + 0.0004 + 0.0004) / 3)) * math.sqrt(252)
    assert annualized_vol(rets) == pytest.approx(expected)

def test_rsi14():
    assert rsi14([float(i) for i in range(20)]) == pytest.approx(100.0)
    assert rsi14([float(-i) for i in range(20)]) == pytest.approx(0.0)
    assert rsi14([100.0] * 5) is None

def test_sma200_flag():
    closes = [100.0] * 200 + [110.0]
    assert sma200_flag(closes) == "above"
    assert sma200_flag([110.0] + [100.0] * 199) == "below"
    assert sma200_flag([100.0] * 100) is None

def test_vol_regime():
    assert vol_regime([100.0] * 80) == "LOW"
    state = volatility_state([100.0] * 40 + [50.0] + [100.0] * 39)
    assert state["level"] == "HIGH"
    assert state["trend"] == "FALLING"


def test_volatility_rising_is_not_inverted():
    prices = [100.0] * 61
    for move in (110, 90, 112, 88, 115, 85, 118, 82, 120, 80,
                 121, 79, 122, 78, 123, 77, 124, 76, 125, 75):
        prices.append(float(move))
    assert volatility_state(prices)["trend"] == "RISING"

def test_momentum_12_1():
    flat_then_recent_jump = [100.0] * 232 + [110.0] * 21
    assert momentum_12_1(flat_then_recent_jump) == pytest.approx(0.0)
    twelve_to_one_gain = [100.0] + [110.0] * 231 + [200.0] * 21
    assert momentum_12_1(twelve_to_one_gain) == pytest.approx(0.1)
    assert momentum_12_1([100.0] * 100) is None

def test_composite_missing_is_not_neutralized():
    s = composite_score(None, None, None, (None, None), None)
    assert s is None
    assert rating(s) == "بيانات ناقصة"

def test_composite_full_marks():
    # PE 6->100, ROE 25->100, divY 6->100, tech (above, mom 0.2)->100, DD -0.10->100
    s = composite_score(6.0, 25.0, 6.0, ("above", 0.2), -0.10)
    assert s == 100
    assert rating(s) == "شراء قوي"

def test_composite_value_only():
    result = composite_score_with_coverage(12.0, None, None, (None, None), None)
    assert result["score"] == pytest.approx(80.0)
    assert result["completeness"] == pytest.approx(0.30)
    assert result["actionable"] is False
    assert rating(result["score"], result["completeness"]) == "بيانات ناقصة"
