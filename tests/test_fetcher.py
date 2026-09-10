from fetcher import normalize_security


def test_normalize_us_default_is_auditable():
    security = normalize_security("AAPL")
    assert security["security_id"] == "US-AAPL"
    assert security["exchange"] == "NASDAQ"
    assert security["currency"] == "USD"
    assert security["settlement"] == "T+1"
    assert security["timezone"] == "America/New_York"
    assert "sec.gov" in security["rule_source"]
    assert "broker" in security["quantity_rule"].lower()


def test_normalize_explicit_us_exchange_and_instrument():
    security = normalize_security({
        "ticker": "msft", "exchange": "nyse", "instrument_type": "ETF",
    })
    assert security["ticker"] == "MSFT"
    assert security["security_id"] == "US-MSFT"
    assert security["exchange"] == "NYSE"
    assert security["instrument_type"] == "ETF"


def test_normalize_saudi_legacy_code():
    security = normalize_security("2222")
    assert security["security_id"] == "SA-2222"
    assert security["exchange"] == "Tadawul"
    assert security["currency"] == "SAR"
    assert security["settlement"] == "T+2"
    assert security["timezone"] == "Asia/Riyadh"
    assert "saudiexchange.sa" in security["rule_source"]
