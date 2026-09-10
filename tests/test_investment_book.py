import pytest

from investment_book import apply_ledger, stable_transaction_id, upsert_transactions


def test_stable_id_and_idempotent_import():
    row = {
        "account": "Main", "broker_reference": "ABC-123", "type": "Buy",
        "security_id": "SA-2222", "quantity": 10, "price": 25,
        "currency": "SAR", "fx_to_sar": 1,
    }
    assert stable_transaction_id(row) == stable_transaction_id(dict(row))
    merged = upsert_transactions([], [row, dict(row)])
    assert len(merged) == 1


def test_weighted_average_sale_and_realized_pl():
    rows = [
        {"transaction_id": "OPEN", "type": "Opening balance", "security_id": "SA-2222",
         "quantity": 10, "price": 20, "currency": "SAR", "fx_to_sar": 1},
        {"transaction_id": "BUY", "type": "Buy", "security_id": "SA-2222",
         "quantity": 10, "price": 30, "fees": 2, "currency": "SAR", "fx_to_sar": 1},
        {"transaction_id": "SELL", "type": "Sell", "security_id": "SA-2222",
         "quantity": 5, "price": 40, "fees": 1, "currency": "SAR", "fx_to_sar": 1},
    ]
    book = apply_ledger(rows)
    assert book["positions"]["SA-2222"]["quantity"] == pytest.approx(15)
    assert book["positions"]["SA-2222"]["average_cost_sar"] == pytest.approx(25.1)
    assert book["rows"][-1]["realized_pl_sar"] == pytest.approx(73.5)


def test_us_fx_is_retained_and_converted():
    book = apply_ledger([{
        "transaction_id": "US1", "type": "Buy", "security_id": "US-AAPL",
        "quantity": 2, "price": 100, "fees": 1, "currency": "USD", "fx_to_sar": 3.75,
    }])
    assert book["positions"]["US-AAPL"]["cost_basis_sar"] == pytest.approx(753.75)
    assert book["cash"]["USD"] == pytest.approx(-201)


def test_non_sar_requires_fx():
    with pytest.raises(ValueError, match="fx_to_sar"):
        apply_ledger([{
            "type": "Buy", "security_id": "US-AAPL", "quantity": 1,
            "price": 100, "currency": "USD",
        }])


def test_us_cash_flow_and_withholding_need_fx_and_keep_currency():
    rows = [
        {"transaction_id": "FX-IN", "type": "FX buy", "cash_amount": 1000,
         "currency": "USD", "fx_to_sar": 3.75},
        {"transaction_id": "DIV", "type": "Dividend", "cash_amount": 20,
         "tax": 3, "currency": "USD", "fx_to_sar": 3.76},
        {"transaction_id": "WH", "type": "Withholding", "cash_amount": 2,
         "currency": "USD", "fx_to_sar": 3.76},
        {"transaction_id": "FX-OUT", "type": "FX sell", "cash_amount": 100,
         "currency": "USD", "fx_to_sar": 3.75},
    ]
    book = apply_ledger(rows)
    assert book["cash"]["USD"] == pytest.approx(915)
    assert book["rows"][1]["cash_change_sar"] == pytest.approx(17 * 3.76)
    assert book["rows"][2]["cash_change_original"] == pytest.approx(-2)


def test_non_sar_cash_flow_requires_fx():
    with pytest.raises(ValueError, match="fx_to_sar"):
        apply_ledger([{
            "type": "Deposit", "cash_amount": 100, "currency": "USD",
        }])
