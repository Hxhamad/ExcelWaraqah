import math, pytest
from metrics import (annual_return, annualized_vol, max_drawdown, rsi14,
                     sma200_flag, vol_regime, momentum_12_1, composite_score, rating)

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
    assert vol_regime([100.0] * 80) == "NORMAL"
    assert vol_regime([100.0] * 40 + [50.0] + [100.0] * 39) == "HIGH"

def test_momentum_12_1():
    assert momentum_12_1([100.0] * 231) == pytest.approx(0.0)
    assert momentum_12_1([100.0] * 230 + [110.0]) == pytest.approx(0.1)
    assert momentum_12_1([100.0] * 100) is None

def test_composite_neutral():
    # all N/A -> every sub-score 50 -> composite 50
    s = composite_score(None, None, None, (None, None), None)
    assert s == 50
    assert rating(s) == "تعزيز/احتفاظ"

def test_composite_full_marks():
    # PE 6->100, ROE 25->100, divY 6->100, tech (above, mom 0.2)->100, DD -0.10->100
    s = composite_score(6.0, 25.0, 6.0, ("above", 0.2), -0.10)
    assert s == 100
    assert rating(s) == "شراء قوي"

def test_composite_value_only():
    # only PE known (12 -> 80); others neutral 50
    # composite = 0.3*80 + 0.2*50 + 0.2*50 + 0.15*50 + 0.15*50 = 59
    s = composite_score(12.0, None, None, (None, None), None)
    assert s == pytest.approx(59.0)
