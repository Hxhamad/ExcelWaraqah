"""Workbook builder for the Saudi/US personal investment book, v3.

Builds a 12-sheet xlsx from the CSV/JSON artefacts produced by fetcher.py:

    Portfolio | Orders | Stock Lookup | Performance | Activity | Sharia |
    Risk & Horizons | DB | Statements | Symbols | Checks | Guide

Design notes that matter for anyone editing this file
-----------------------------------------------------
* openpyxl 3.1.5 quirks are load-bearing here: PatternFill is always built with
  keyword args (positional fill types blow up at save time), and every merged
  range gets its anchor cell written BEFORE the merge call.
* Every row number is a module-level constant; formulas are assembled with
  f-strings so a layout change never leaves a stale hardcoded row behind.
* Formulas favor broadly supported Excel/Google Sheets functions.  A small
  number of modern functions such as IFS are used where they materially reduce
  nested-formula risk; LibreOffice and live Google Sheets are both verified.
* Static analytics (scores, ratings, horizon verdicts, P/E, missing-data notes)
  are computed in Python and written as values; formulas are only used where the
  user's own inputs must flow through.

DB row ordering
---------------
DB holds one row per symbol-year, symbols ascending, and years DESCENDING
(latest year first). Descending is deliberate: a plain VLOOKUP returns the FIRST
matching row, so putting the latest year first is what makes any by-symbol
VLOOKUP into DB mean "the current year". Stock Lookup does not depend on the
ordering at all - it addresses DB through helper key columns (below), so it
stays correct under either order.

DB helper columns (all static strings, built here in Python)
    O  na key      "<symbol>|ناقص" when that row has missing inputs, else blank
    P  year key    "<symbol>|<year>"  -> per-year INDEX/MATCH from Stock Lookup
    Q  latest key  "<symbol>" on the symbol's latest-year row only, blank elsewhere
    R  payout      payout ratio, latest-year row only
Stock Lookup reads "the latest year for the typed symbol" as
MATCH($B$3,DB!$Q:$Q,0) - one exact text match, no embedded year literals, and it
keeps working when a symbol's history stops short of the newest year.
"""

import argparse
import csv
import json
import os
import sys
from datetime import date, datetime

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import metrics  # noqa: E402
from symbols import load_symbols  # noqa: E402

BOOK_BUILD_DATE = date.today().isoformat()
BOOK_BUILD_ID_DATE = BOOK_BUILD_DATE.replace("-", "")

# --------------------------------------------------------------------------
# layout constants
# --------------------------------------------------------------------------
SHEET_ORDER = [
    "Portfolio", "Orders", "Stock Lookup", "Performance", "Activity",
    "Sharia", "Risk & Horizons", "DB", "Statements", "Symbols", "Checks", "Guide",
]

# Portfolio
PF_TITLE_ROW = 2
PF_NOTE_ROW = 3
PF_HDR_ROW = 5
PF_FIRST_ROW = 6
PF_LAST_ROW = 25
PF_TOTAL_ROW = 26
PF_SUM_TITLE_ROW = 27
PF_SUM_COUNT_ROW = 28
PF_SUM_PL_ROW = 29
PF_SUM_MAXW_ROW = 30
PF_SUM_CONC_ROW = 31
PF_SUM_R1_ROW = 32   # شراء قوي
PF_SUM_R2_ROW = 33   # شراء
PF_SUM_R3_ROW = 34   # تعزيز/احتفاظ
PF_SUM_R4_ROW = 35   # بيع / بيع قوي
PF_SUM_SECTOR_ROW = 36

# Stock Lookup
SL_TITLE_ROW = 1
SL_INPUT_ROW = 3
SL_GUARD_ROW = 4
SL_MEMBER_ROW = 5
SL_PRICE_HDR_ROW = 7
SL_PRICE_FIRST_ROW = 8          # 7 items -> 8..14
SL_RISK_HDR_ROW = 16
SL_RISK_FIRST_ROW = 17          # 4 items -> 17..20
SL_FUND_HDR_ROW = 22
SL_FUND_FIRST_ROW = 23          # 4 items -> 23..26
SL_YEARS_HDR_ROW = 28
SL_YEARS_FIRST_ROW = 29         # 5 years -> 29..33
SL_VERDICT_HDR_ROW = 35
SL_VERDICT_FIRST_ROW = 36       # 5 items + note -> 36..41

# Risk & Horizons
RH_TITLE_ROW = 2
RH_HDR_ROW = 4
RH_FIRST_ROW = 5

# DB
DB_TITLE_ROW = 2
DB_HDR_ROW = 4
DB_FIRST_ROW = 5

# Statements
ST_TITLE_ROW = 1
ST_GROUP_ROW = 3
ST_HDR_ROW = 4
ST_FIRST_ROW = 5

# Symbols
SY_TITLE_ROW = 2
SY_HDR_ROW = 4
SY_FIRST_ROW = 5

# Guide
GD_FIRST_ROW = 2

STMT_YEARS = [2025, 2024, 2023, 2022, 2021]   # newest first, 3 columns each

DEMO_HOLDINGS = [
    ("2222", 500, 26.50),
    ("1120", 150, 89.00),
    ("2010", 200, 62.00),
    ("7010", 300, 42.00),
]

# --------------------------------------------------------------------------
# styles
# --------------------------------------------------------------------------
F_TITLE = Font(name="Arial", size=14, bold=True, color="1F3864")
F_NOTE = Font(name="Arial", size=9, italic=True, color="595959")
F_HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_HDR_DARK = Font(name="Arial", size=10, bold=True, color="1F3864")
F_SEC = Font(name="Arial", size=11, bold=True, color="1F3864")
F_LBL = Font(name="Arial", size=10)
F_LBL_B = Font(name="Arial", size=10, bold=True)
F_VAL = Font(name="Arial", size=10)
F_IN = Font(name="Arial", size=10, bold=True, color="0000FF")
F_STAT = Font(name="Arial", size=10, color="006100")
F_TOT = Font(name="Arial", size=10, bold=True)
F_WARN = Font(name="Arial", size=10, bold=True, color="C00000")
F_GUIDE = Font(name="Arial", size=10)
F_GUIDE_H = Font(name="Arial", size=12, bold=True, color="1F3864")

FILL_HDR = PatternFill(fill_type="solid", fgColor="1F3864")
FILL_SEC = PatternFill(fill_type="solid", fgColor="D9E2F3")
FILL_GRAY = PatternFill(fill_type="solid", fgColor="D9D9D9")
FILL_IN = PatternFill(fill_type="solid", fgColor="FFF2CC")
FILL_TOT = PatternFill(fill_type="solid", fgColor="EDEDED")

THIN = Side(style="thin", color="BFBFBF")
THICK = Side(style="medium", color="BF8F00")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
B_IN = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
WRAP = Alignment(horizontal="right", vertical="top", wrap_text=True)

FMT_MONEY = "#,##0.00"
FMT_INT = "#,##0"
FMT_PCT = "0.0%"
FMT_PCTV = '0.00"%"'      # value already scaled x100
FMT_NUM2 = "0.00"
FMT_NUM1 = "0.0"
FMT_TEXT = "@"

NA = "بيانات ناقصة"

# sector -> oil beta. Anything unmapped falls through to محايدة.
OIL_BETA_MAP = {
    "energy": "موجبة",
    "oil": "موجبة",
    "utilities": "موجبة",
    "basic materials": "موجبة",
    "materials": "موجبة",
    "petrochemical": "موجبة",
    "bank": "سالبة",
    "financial": "سالبة",
    "insurance": "سالبة",
    "consumer": "محايدة",
    "healthcare": "محايدة",
    "health care": "محايدة",
    "telecom": "محايدة",
    "communication": "محايدة",
    "real estate": "محايدة",
    "industrial": "محايدة",
    "technology": "محايدة",
}


def oil_beta(sector):
    """Sector -> oil sensitivity label; anything unmapped is محايدة."""
    key = (sector or "").strip().lower()
    if key in OIL_BETA_MAP:
        return OIL_BETA_MAP[key]
    for token, label in OIL_BETA_MAP.items():
        if token in key:
            return label
    return "محايدة"


# --------------------------------------------------------------------------
# data loading
# --------------------------------------------------------------------------

def _num(value):
    """CSV cell -> float or None (blank, 'nan' and junk all become None)."""
    if value is None:
        return None
    text = str(value).strip()
    if text == "" or text.lower() in ("nan", "none", "null"):
        return None
    try:
        out = float(text)
    except ValueError:
        return None
    if out != out:  # NaN
        return None
    return out


def _int(value):
    out = _num(value)
    return None if out is None else int(out)


def _read_csv(path):
    if not os.path.exists(path):
        print("WARNING: missing data file %s (continuing without it)" % path)
        return []
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def load_data(data_dir):
    """Read annual_metrics.csv / statements.csv / snapshot.json / symbols.csv.

    Missing files or missing fields degrade to empty structures and None rather
    than raising, so a partial refresh still produces a workbook.
    """
    annual = {}
    for row in _read_csv(os.path.join(data_dir, "annual_metrics.csv")):
        sym = (row.get("symbol") or "").strip()
        year = _int(row.get("year"))
        if not sym or year is None:
            continue
        annual.setdefault(sym, {})[year] = {
            "close": _num(row.get("close")),
            "ret": _num(row.get("ret")),
            "vol": _num(row.get("vol")),
            "maxdd": _num(row.get("maxdd")),
            "divs": _num(row.get("divs")),
            "div_yield": _num(row.get("div_yield")),
            "momentum": _num(row.get("momentum")),
            "eps": _num(row.get("eps")),
        }

    stmts = {}
    for row in _read_csv(os.path.join(data_dir, "statements.csv")):
        sym = (row.get("symbol") or "").strip()
        year = _int(row.get("year"))
        if not sym or year is None:
            continue
        stmts.setdefault(sym, {})[year] = {
            "revenue": _num(row.get("revenue")),
            "net_income": _num(row.get("net_income")),
            "eps": _num(row.get("eps")),
            "roe": _num(row.get("roe")),
            "de": _num(row.get("de")),
            "payout": _num(row.get("payout")),
        }

    snapshot = {}
    snap_path = os.path.join(data_dir, "snapshot.json")
    if os.path.exists(snap_path):
        with open(snap_path, "r", encoding="utf-8") as fh:
            snapshot = json.load(fh) or {}
    else:
        print("WARNING: missing data file %s (continuing without it)" % snap_path)

    names = {}
    sym_csv = os.path.join(HERE, "symbols.csv")
    if os.path.exists(sym_csv):
        try:
            for row in load_symbols(sym_csv):
                names[row["code"]] = row
        except (ValueError, FileNotFoundError) as exc:
            print("WARNING: symbols.csv unusable (%s); falling back to snapshot names" % exc)

    def entry(code):
        snap = snapshot.get(code) or {}
        meta = names.get(code) or {}
        name_ar = (meta.get("name_ar") or "").strip()
        name_en = (meta.get("name_en") or snap.get("name_en") or "").strip()
        return {
            "code": code,
            "name": name_ar or name_en or code,
            "name_ar": name_ar,
            "name_en": name_en,
            "sector": (meta.get("sector") or snap.get("sector") or "").strip(),
        }

    # symbols.csv is the full listed universe; a refresh may only have fetched
    # part of it. The analysis sheets cover the symbols that actually carry data
    # - a row of blanks is worse than no row - while the Symbols sheet lists
    # every listed code and marks the un-fetched ones with ✗.
    codes_all = sorted(set(list(annual) + list(stmts) + list(snapshot) + list(names)))
    codes_data = sorted(set(list(annual) + list(stmts) + list(snapshot)))
    if len(codes_data) < len(codes_all):
        print("NOTE: %d of %d listed symbols have data; the rest appear on the "
              "Symbols sheet marked ✗" % (len(codes_data), len(codes_all)))

    return {
        "annual": annual,
        "statements": stmts,
        "snapshot": snapshot,
        "universe": [entry(code) for code in codes_data],
        "all_symbols": [entry(code) for code in codes_all],
    }


# --------------------------------------------------------------------------
# analytics computed in python (written to the sheets as static values)
# --------------------------------------------------------------------------

def latest_year(years):
    return max(years) if years else None


def pe_value(close, eps):
    """P/E as a rounded number, or the Arabic missing-data marker."""
    if close is None or eps is None or eps <= 0:
        return NA
    return round(close / eps, 1)


def _verdict(points, present):
    """Map a signed point total to an Arabic verdict.

    `present` counts how many inputs were actually available: with none of them
    we say so explicitly instead of pretending the neutral band was earned.
    """
    if present == 0:
        return "محايد (بيانات ناقصة)"
    if points >= 2:
        return "إيجابي"
    if points <= -2:
        return "سلبي"
    return "محايد"


def near_verdict(sma_flag, rsi, vol_level, vol_trend, ret_3m, ret_6m):
    """Six-month view: medium trend, momentum, RSI, and volatility direction."""
    pts = 0
    present = 0
    if sma_flag:
        present += 1
        pts += 1 if sma_flag == "above" else -1
    if rsi is not None:
        present += 1
        if rsi > 70:
            pts -= 1
        elif rsi < 30:
            pts += 1
        elif rsi >= 40:
            pts += 1
    if vol_level:
        present += 1
        if vol_level == "HIGH":
            pts -= 1
    if vol_trend:
        present += 1
        pts += 1 if vol_trend == "FALLING" else (-1 if vol_trend == "RISING" else 0)
    if ret_3m is not None:
        present += 1
        pts += 1 if ret_3m > 0 else -1
    if ret_6m is not None:
        present += 1
        pts += 1 if ret_6m > 0 else -1
    return _verdict(pts, present)


def mid_verdict(ret_1y, momentum, sma_flag, pe, maxdd_2y):
    """Two-year view: return, 12-1 momentum, trend, value and drawdown."""
    pts = 0
    present = 0
    if ret_1y is not None:
        present += 1
        pts += 1 if ret_1y > 0 else -1
    if momentum is not None:
        present += 1
        pts += 1 if momentum > 0 else -1
    if sma_flag:
        present += 1
        pts += 1 if sma_flag == "above" else -1
    if pe is not None and pe > 0:
        present += 1
        if pe <= 18:
            pts += 1
        elif pe > 25:
            pts -= 1
    if maxdd_2y is not None:
        present += 1
        pts += 1 if maxdd_2y >= -0.25 else -1
    return _verdict(pts, present)


def far_verdict(roe, div_yield, payout, pe):
    """Five-year view: quality, income sustainability and valuation."""
    pts = 0
    present = 0
    if roe is not None:
        present += 1
        if roe >= 20:
            pts += 2
        elif roe >= 15:
            pts += 1
        elif roe < 5:
            pts -= 1
    if div_yield is not None:
        present += 1
        if div_yield >= 3:
            pts += 1
    if payout is not None:
        present += 1
        if payout > 1.0:
            pts -= 1
        elif payout <= 0.8:
            pts += 1
    if pe is not None and pe > 0:
        present += 1
        if pe <= 18:
            pts += 1
        elif pe > 25:
            pts -= 1
    return _verdict(pts, present)


def build_rows(data):
    """One analytic record per symbol, shared by Risk & Horizons / DB / Lookup."""
    rows = []
    for entry in data["universe"]:
        code = entry["code"]
        snap = data["snapshot"].get(code) or {}
        info = snap.get("info") or {}
        years = data["annual"].get(code) or {}
        last = years.get(latest_year(years)) or {}
        stmt_years = data["statements"].get(code) or {}
        stmt_last = (stmt_years.get(max(stmt_years)) if stmt_years else None) or {}

        price = snap.get("price")
        if price is None:
            price = last.get("close")
        pe = info.get("pe")
        roe = info.get("roe")
        if roe is None:
            roe = stmt_last.get("roe")
        payout = info.get("payout")
        if payout is None:
            payout = stmt_last.get("payout")
        div_yield = info.get("div5y")
        if div_yield is None:
            div_yield = last.get("div_yield")
        maxdd_2y = snap.get("maxdd_2y")
        if maxdd_2y is None:
            maxdd_2y = last.get("maxdd")
        sma_flag = snap.get("sma200_flag")
        vol_level = snap.get("vol_level") or snap.get("vol_regime")
        vol_trend = snap.get("vol_trend")
        rsi = snap.get("rsi14")
        momentum = last.get("momentum")

        score_meta = metrics.composite_score_with_coverage(
            pe, roe, div_yield, (sma_flag, momentum), maxdd_2y
        )
        score = score_meta["score"]
        # Flag only the inputs that actually feed the score and the verdicts -
        # this is what "الحكم محايد للأجزاء الناقصة" refers to. Gaps in the old
        # annual history are reported separately (Statements col "ملاحظة").
        missing_inputs = any(v is None for v in
                             (pe, roe, div_yield, sma_flag, momentum, maxdd_2y))
        data_as_of = snap.get("price_as_of")
        freshness = "Unknown — refresh required"
        if data_as_of:
            try:
                age_days = (date.today() - datetime.fromisoformat(
                    str(data_as_of).replace("Z", "+00:00")
                ).date()).days
                freshness = "Fresh" if age_days <= 5 else "Stale"
            except (TypeError, ValueError):
                freshness = "Unknown — invalid date"
        actionable = bool(score_meta["actionable"] and freshness == "Fresh")
        rows.append({
            "code": code,
            "name": entry["name"],
            "name_ar": entry["name_ar"],
            "name_en": entry["name_en"],
            "sector": entry["sector"],
            "price": price,
            "ret_1w": snap.get("ret_1w"),
            "ret_1m": snap.get("ret_1m"),
            "ret_3m": snap.get("ret_3m"),
            "ret_6m": snap.get("ret_6m"),
            "ret_1y": snap.get("ret_1y"),
            "momentum": momentum,
            "rsi": rsi,
            "vol_regime": ({"HIGH": "مرتفع", "NORMAL": "عادي", "LOW": "منخفض"}
                           .get(vol_level, NA)),
            "vol_trend": ({"RISING": "صاعد", "STABLE": "مستقر", "FALLING": "هابط"}
                          .get(vol_trend, NA)),
            "sma200": "فوق" if sma_flag == "above" else ("تحت" if sma_flag == "below" else NA),
            "maxdd_2y": maxdd_2y,
            "oil_beta": oil_beta(entry["sector"]),
            "near": near_verdict(sma_flag, rsi, vol_level, vol_trend,
                                 snap.get("ret_3m"), snap.get("ret_6m")),
            "mid": mid_verdict(snap.get("ret_1y"), momentum, sma_flag, pe, maxdd_2y),
            "far": far_verdict(roe, div_yield, payout, pe),
            "score": score,
            "rating": metrics.rating(score, score_meta["completeness"]),
            "completeness": score_meta["completeness"],
            "actionable": actionable,
            "data_as_of": data_as_of,
            "freshness": freshness,
            "fetched_at": snap.get("fetched_at"),
            "source": snap.get("source") or "Yahoo Finance (legacy snapshot)",
            "source_url": snap.get("source_url"),
            "security_id": snap.get("security_id") or ("SA-" + code),
            "exchange": snap.get("exchange") or "Tadawul",
            "currency": snap.get("currency") or "SAR",
            "fx_to_sar": 1.0 if (snap.get("currency") or "SAR") == "SAR" else None,
            "market_calendar": snap.get("calendar") or (
                "Saudi Exchange" if (snap.get("exchange") or "Tadawul") == "Tadawul"
                else "US market"
            ),
            "settlement": snap.get("settlement") or (
                "T+2" if (snap.get("exchange") or "Tadawul") == "Tadawul" else "T+1"
            ),
            "fundamentals_period": str(max(stmt_years)) if stmt_years else None,
            "instrument_type": snap.get("instrument_type") or "Equity",
            "market_timezone": snap.get("market_timezone") or (
                "Asia/Riyadh" if (snap.get("exchange") or "Tadawul") == "Tadawul"
                else "America/New_York"
            ),
            "market_session": snap.get("market_session") or (
                "Sunday–Thursday; core 10:00–15:00" if
                (snap.get("exchange") or "Tadawul") == "Tadawul"
                else "Core session 09:30–16:00 ET"
            ),
            "dst_handling": snap.get("dst_handling") or (
                "No daylight-saving shift" if
                (snap.get("exchange") or "Tadawul") == "Tadawul"
                else "America/New_York daylight-saving rules"
            ),
            "quantity_rule": snap.get("quantity_rule") or (
                "Whole shares; minimum 1" if
                (snap.get("exchange") or "Tadawul") == "Tadawul"
                else "Whole/fractional shares depend on broker and security"
            ),
            "order_rule": snap.get("order_rule") or
                "Broker/account capabilities must be confirmed before approval",
            "fundamentals_basis": snap.get("fundamentals_basis") or
                "Annual history; trailing ratios only where labelled",
            "publication_date": snap.get("publication_date"),
            "observation_time": snap.get("price_observed_at") or data_as_of,
            "retrieval_time": snap.get("fetched_at"),
            "price_basis": snap.get("price_basis") or "Adjusted close for return series",
            "return_basis": snap.get("return_basis") or
                "Total return from adjusted close where available",
            "dividend_unit": snap.get("dividend_unit") or
                "%s per share" % (snap.get("currency") or "SAR"),
            "market_rule_source": snap.get("market_rule_source") or (
                "https://www.saudiexchange.sa/wps/portal/saudiexchange/trading/market-services/equities?locale=en"
                if (snap.get("exchange") or "Tadawul") == "Tadawul" else
                "https://www.sec.gov/rules-regulations/2023/02/34-96930"
            ),
            "pe": pe,
            "roe": roe,
            "payout": payout,
            "div_yield": div_yield,
            "missing_inputs": missing_inputs,
        })
    return rows


# --------------------------------------------------------------------------
# small sheet helpers
# --------------------------------------------------------------------------

def put(ws, row, col, value, font=None, fill=None, fmt=None, align=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font is not None:
        cell.font = font
    if fill is not None:
        cell.fill = fill
    if fmt is not None:
        cell.number_format = fmt
    if align is not None:
        cell.alignment = align
    if border is not None:
        cell.border = border
    return cell


def merge_band(ws, row, first_col, last_col, text, font=F_SEC, fill=FILL_SEC, align=CENTER):
    """Write the anchor cell FIRST, then merge: openpyxl silently drops writes
    to a cell that is already inside a merged range."""
    put(ws, row, first_col, text, font=font, fill=fill, align=align)
    for col in range(first_col + 1, last_col + 1):
        ws.cell(row=row, column=col).fill = fill
    ws.merge_cells(start_row=row, start_column=first_col,
                   end_row=row, end_column=last_col)


def header_row(ws, row, first_col, labels, font=F_HDR, fill=FILL_HDR):
    for i, label in enumerate(labels):
        put(ws, row, first_col + i, label, font=font, fill=fill,
            align=CENTER, border=B_ALL)
    ws.row_dimensions[row].height = 30


def set_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def set_filter(ws, hdr_row, last_row, last_col):
    """Put an auto-filter over the header row + its data rows.

    Silently does nothing when the table has no data rows, because Excel
    rejects a filter whose range is the header alone.
    """
    if last_row < hdr_row + 1:
        return
    ws.auto_filter.ref = "A%d:%s%d" % (hdr_row, get_column_letter(last_col), last_row)


# --------------------------------------------------------------------------
# Portfolio
# --------------------------------------------------------------------------

def read_preserved(path):
    """Pull (symbol, shares, cost) from Portfolio!B/E/F rows 6-25 of an old build.

    Returns [] (and prints why) when the file is unreadable or has no Portfolio
    sheet, so a bad --preserve never kills the build.
    """
    if not path:
        return []
    if not os.path.exists(path):
        print("WARNING: --preserve file not found: %s (using demo holdings)" % path)
        return []
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as exc:                    # noqa: BLE001 - any xlsx defect
        print("WARNING: could not read --preserve workbook (%s); using demo holdings" % exc)
        return []
    if "Portfolio" not in wb.sheetnames:
        print("WARNING: --preserve workbook has no Portfolio sheet; using demo holdings")
        wb.close()
        return []
    ws = wb["Portfolio"]
    out = []
    for row in range(PF_FIRST_ROW, PF_LAST_ROW + 1):
        sym = ws.cell(row=row, column=2).value
        shares = ws.cell(row=row, column=5).value
        cost = ws.cell(row=row, column=6).value
        if sym is None or str(sym).strip() == "":
            continue
        if isinstance(sym, float) and sym.is_integer():
            sym = int(sym)
        out.append((str(sym).strip(), _num(shares), _num(cost)))
    wb.close()
    print("preserved %d holding row(s) from %s" % (len(out), path))
    return out


def restore_manual_records(wb, path):
    """Restore durable user inputs/history without copying calculated cells.

    This makes refreshes non-destructive for the transaction ledger, Sharia
    evidence, proposal decisions, performance history, owner settings and
    review log.  Formula/helper columns are regenerated from the current code.
    """
    if not path or not os.path.exists(path):
        return
    try:
        old = load_workbook(path, data_only=False)
    except Exception as exc:  # noqa: BLE001
        print("WARNING: could not restore manual book records (%s)" % exc)
        return

    # Activity is intentionally positional: its schema is stable and formula
    # columns are excluded.  The other durable records are restored by header
    # name below so schema upgrades do not silently move owner inputs.
    specs = {
        "Activity": (6, 505, list(range(1, 18)) + [25, 26, 28]),
    }
    for name, (first, last, columns) in specs.items():
        if name not in old.sheetnames or name not in wb.sheetnames:
            continue
        src, dst = old[name], wb[name]
        for row in range(first, min(last, src.max_row) + 1):
            for col in columns:
                value = src.cell(row=row, column=col).value
                if value is not None:
                    dst.cell(row=row, column=col).value = value
    def headers(ws, row=5):
        return {str(ws.cell(row=row, column=col).value).strip(): col
                for col in range(1, ws.max_column + 1)
                if ws.cell(row=row, column=col).value is not None}

    def keyed_rows(ws, key_col, first=6, last=505):
        result = {}
        for row in range(first, min(last, ws.max_row) + 1):
            key = ws.cell(row=row, column=key_col).value
            if key is not None and str(key).strip():
                result[str(key).strip()] = row
        return result

    def first_blank_key_row(ws, key_col, first=6, last=505):
        for row in range(first, last + 1):
            value = ws.cell(row=row, column=key_col).value
            if value is None or not str(value).strip():
                return row
        return None

    # Sharia: preserve evidence and owner review fields, but regenerate the
    # eligibility/holding-review formulas.  Legacy v2 labels are mapped to the
    # richer v3 evidence schema without pretending provider == methodology.
    if "Sharia" in old.sheetnames and "Sharia" in wb.sheetnames:
        src, dst = old["Sharia"], wb["Sharia"]
        sh, dh = headers(src), headers(dst)
        aliases = {
            "Methodology / Provider": "Authority / Provider",
            "Business Activity": "Business Activity Evidence",
            "Financial Ratios": "Financial Ratio Evidence",
            "Purification Notes": "Purification Method",
        }
        allowed = {
            "Security ID", "Ticker", "Company", "Exchange", "Status",
            "Authority / Provider", "Methodology", "Methodology Version",
            "Evidence URL", "Reporting Period", "Screen Date", "Next Review",
            "Business Activity Evidence", "Financial Ratio Evidence",
            "Purification Method", "Purification Due SAR", "Purification Paid SAR",
            "Payment Date", "Change Since Prior Review", "Reviewer", "Notes",
        }
        src_key = sh.get("Security ID")
        dst_key = dh.get("Security ID")
        if src_key and dst_key:
            destinations = keyed_rows(dst, dst_key)
            for src_row in range(6, min(505, src.max_row) + 1):
                key = src.cell(src_row, src_key).value
                if key is None or not str(key).strip():
                    continue
                key = str(key).strip()
                dst_row = destinations.get(key) or first_blank_key_row(dst, dst_key)
                if dst_row is None:
                    break
                destinations[key] = dst_row
                for old_name, src_col in sh.items():
                    new_name = aliases.get(old_name, old_name)
                    if new_name not in allowed or new_name not in dh:
                        continue
                    value = src.cell(src_row, src_col).value
                    if value is not None and not (isinstance(value, str) and value.startswith("=")):
                        dst.cell(dst_row, dh[new_name]).value = value

    # Orders: preserve proposal identity, owner decisions and broker evidence;
    # regenerate research summaries, risk math and control formulas every run.
    if "Orders" in old.sheetnames and "Orders" in wb.sheetnames:
        src, dst = old["Orders"], wb["Orders"]
        sh, dh = headers(src), headers(dst)
        aliases = {
            "Quantity": "Proposed Quantity", "Stop Price": "Stop / Review Price",
            "Expiry": "Valid Until", "Risk / Invalidation": "Thesis Invalidation",
            "Post-Trade Position %": "Proposed Position %",
            "Post-Trade Sector %": "Proposed Sector %",
        }
        allowed = {
            "Proposal ID", "Version", "Rank", "Created At", "Reviewed At", "Horizon",
            "Security ID", "Ticker", "Company", "Exchange", "Broker", "Account",
            "Currency", "Action", "Order Type", "Entry Condition", "Proposed Quantity",
            "Limit Price", "Stop / Review Price", "Valid Until", "Evidence URLs",
            "Why I Own It", "Thesis Invalidation", "Confidence", "Status",
            "User Decision", "Decision Date", "Broker Order ID", "Filled Quantity",
            "Average Fill Price", "Execution Transaction ID", "Review ID",
            "Last Checked", "Notes",
        }
        action_map = {"Sell": "Exit"}
        status_map = {
            "Draft": "Proposed", "Rejected": "Cancelled", "Executed": "Filled",
        }
        src_key = sh.get("Proposal ID")
        dst_key = dh.get("Proposal ID")
        if src_key and dst_key:
            destinations = keyed_rows(dst, dst_key)
            for src_row in range(6, min(505, src.max_row) + 1):
                key = src.cell(src_row, src_key).value
                if key is None or not str(key).strip():
                    continue
                key = str(key).strip()
                dst_row = destinations.get(key) or first_blank_key_row(dst, dst_key)
                if dst_row is None:
                    break
                destinations[key] = dst_row
                for old_name, src_col in sh.items():
                    new_name = aliases.get(old_name, old_name)
                    if new_name not in allowed or new_name not in dh:
                        continue
                    value = src.cell(src_row, src_col).value
                    if value is None or (isinstance(value, str) and value.startswith("=")):
                        continue
                    if (new_name == "Thesis Invalidation" and isinstance(value, str) and
                            value.startswith("Unknown Sharia result")):
                        # This was v2 generated blocker text, not an owner thesis.
                        continue
                    if new_name == "Action":
                        value = action_map.get(str(value), value)
                    elif new_name == "Status":
                        value = status_map.get(str(value), value)
                    dst.cell(dst_row, dh[new_name]).value = value

    # Performance history is located by its Date header because v2 used row 20
    # and v3 uses row 30.  Formula-generated opening values are not imported as
    # owner facts; manually entered values and evidence are.
    if "Performance" in old.sheetnames and "Performance" in wb.sheetnames:
        src, dst = old["Performance"], wb["Performance"]
        src_header = next((r for r in range(1, min(100, src.max_row) + 1)
                           if src.cell(r, 1).value == "Date"), None)
        if src_header:
            dst_row = 31
            for src_row in range(src_header + 1, min(505, src.max_row) + 1):
                if all(src.cell(src_row, col).value is None for col in range(1, 10)):
                    continue
                for col in (1, 2, 3, 6, 7, 8, 9):
                    value = src.cell(src_row, col).value
                    if (col == 9 and isinstance(value, str) and
                            value.startswith("Opening snapshot; historical cash flows")):
                        continue
                    if value is not None and not (isinstance(value, str) and value.startswith("=")):
                        dst.cell(dst_row, col).value = value
                dst_row += 1

    # Settings and review history are matched by labels, not row numbers.
    if "Checks" in old.sheetnames and "Checks" in wb.sheetnames:
        src, dst = old["Checks"], wb["Checks"]
        dst_settings = {str(dst.cell(r, 1).value).strip(): r for r in range(5, 21)
                        if dst.cell(r, 1).value is not None}
        for src_row in range(5, min(30, src.max_row) + 1):
            key = src.cell(src_row, 1).value
            if key is None or str(key).strip() not in dst_settings:
                continue
            value = src.cell(src_row, 2).value
            if value is not None and not (isinstance(value, str) and value.startswith("=")):
                dst.cell(dst_settings[str(key).strip()], 2).value = value
        src_review_header = next((r for r in range(30, min(100, src.max_row) + 1)
                                  if src.cell(r, 1).value == "Review ID"), None)
        if src_review_header:
            dst_row = 45
            for src_row in range(src_review_header + 1, min(505, src.max_row) + 1):
                if src.cell(src_row, 1).value is None:
                    continue
                for col in range(1, 7):
                    value = src.cell(src_row, col).value
                    if value is not None:
                        dst.cell(dst_row, col).value = value
                dst_row += 1
    old.close()


def sheet_portfolio(wb, holdings):
    ws = wb.create_sheet("Portfolio")
    ws.sheet_view.showGridLines = False

    put(ws, PF_TITLE_ROW, 2, "محفظة Waraqah — السعودية والولايات المتحدة", font=F_TITLE)
    put(ws, PF_NOTE_ROW, 2,
        "المراكز والتكلفة مشتقة من Activity. عدّل سجل النشاط فقط؛ التحليل يقرأ من "
        "Risk & Horizons والضوابط من Sharia وChecks.",
        font=F_NOTE)

    headers = [
        "الرمز", "الشركة", "السعر", "عدد الأسهم", "متوسط التكلفة",
        "القيمة السوقية", "التكلفة الإجمالية", "الربح/الخسارة", "نسبة الربح/الخسارة",
        "الوزن", "النتيجة", "التقييم", "الإشارة", "Security ID", "السوق",
        "العملة", "FX إلى SAR", "تاريخ السعر", "اكتمال البيانات", "الحالة الشرعية",
        "مراجعة شرعية قادمة", "6 أشهر", "سنتان", "5 سنوات", "مؤهل للشراء؟",
        "إجراء الوكيل", "العوائق", "رابط المصدر", "حالة الحداثة",
        "القيمة السوقية بالعملة الأصلية", "التكلفة الأصلية المعروفة",
        "الربح/الخسارة بالعملة الأصلية", "أثر FX بالريال", "نوع الأداة", "القطاع",
    ]
    header_row(ws, PF_HDR_ROW, 2, headers, font=F_HDR_DARK, fill=FILL_GRAY)

    for idx in range(PF_LAST_ROW - PF_FIRST_ROW + 1):
        r = PF_FIRST_ROW + idx
        held = holdings[idx] if idx < len(holdings) else None
        put(ws, r, 2, held[0] if held else None, font=F_IN, fmt=FMT_TEXT, border=B_ALL,
            align=CENTER)
        put(ws, r, 3, f"=IFERROR(VLOOKUP(B{r},'Risk & Horizons'!$A:$B,2,0),\"\")",
            font=F_VAL, border=B_ALL)
        put(ws, r, 4, f"=IFERROR(VLOOKUP(B{r},'Risk & Horizons'!$A:$D,4,0),\"\")",
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 5, f'=IF(B{r}="","",SUMIF(Activity!$G$6:$G$505,B{r},Activity!$R$6:$R$505))',
            font=F_VAL, fmt=FMT_INT, border=B_ALL)
        put(ws, r, 6, f'=IFERROR(SUMIF(Activity!$G$6:$G$505,B{r},Activity!$U$6:$U$505)/E{r},"")',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 7, f'=IF(OR(E{r}="",D{r}="",R{r}=""),"",D{r}*E{r}*R{r})',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 8, f'=IF(B{r}="","",SUMIF(Activity!$G$6:$G$505,B{r},Activity!$U$6:$U$505))',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 9, f'=IF(OR(G{r}="",H{r}=""),"",G{r}-H{r})',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 10, f'=IFERROR(I{r}/H{r},"")',
            font=F_VAL, fmt=FMT_PCT, border=B_ALL)
        put(ws, r, 11, f'=IF(OR(G{r}="",$G${PF_TOTAL_ROW}=""),"",G{r}/$G${PF_TOTAL_ROW})',
            font=F_VAL, fmt=FMT_PCT, border=B_ALL)
        # Score/rating come from Risk & Horizons (exactly one row per symbol,
        # cols R=18 and S=19). DB carries the score on the latest-year row only,
        # so a by-symbol VLOOKUP into DB would be hostage to row ordering.
        put(ws, r, 12, f"=IFERROR(VLOOKUP(B{r},'Risk & Horizons'!$A:$S,18,0),\"\")",
            font=F_VAL, fmt=FMT_NUM1, border=B_ALL)
        put(ws, r, 13, f"=IFERROR(VLOOKUP(B{r},'Risk & Horizons'!$A:$S,19,0),\"\")",
            font=F_VAL, border=B_ALL)
        put(ws, r, 14, f'=IF(AA{r}="","",AA{r})', font=F_VAL, border=B_ALL)
        lookups = {
            15: (26, None), 16: (27, None), 17: (28, None), 18: (29, FMT_NUM2),
            19: (23, None), 20: (21, FMT_PCT), 23: (15, None), 24: (16, None),
            25: (17, None), 29: (25, None), 30: (30, None),
        }
        for col, (source_col, fmt) in lookups.items():
            put(ws, r, col,
                f'=IFERROR(VLOOKUP(B{r},\'Risk & Horizons\'!$A:$AD,{source_col},FALSE),"")',
                font=F_VAL, fmt=fmt, border=B_ALL)
        put(ws, r, 21, f'=IFERROR(VLOOKUP(O{r},Sharia!$A:$W,5,FALSE),"Uncertain")',
            font=F_VAL, border=B_ALL)
        put(ws, r, 22, f'=IFERROR(VLOOKUP(O{r},Sharia!$A:$W,12,FALSE),"")',
            font=F_VAL, border=B_ALL)
        put(ws, r, 26, f'=IFERROR(VLOOKUP(O{r},Sharia!$A:$W,22,FALSE),"NO")',
            font=F_VAL, border=B_ALL, align=CENTER)
        put(ws, r, 27,
            f'=IF(B{r}="","",IFS(U{r}="Non-compliant","Review Sharia status",'
            f'U{r}<>"Compliant","Wait — Sharia review",Z{r}<>"YES","Wait — Sharia evidence/expiry",'
            f'AD{r}<>"Fresh","Wait — stale/unknown data",T{r}<Checks!$B$7,"Wait — incomplete data",'
            f'OR(Checks!$B$9="",Checks!$B$10="",Checks!$B$11="",'
            f'IF(Q{r}="SAR",Checks!$B$13,Checks!$B$18)=""),'
            f'"Review only — limits/cash pending",AND(Checks!$B$10<>"",K{r}>=Checks!$B$10),'
            f'"Reduce concentration",AND(W{r}="إيجابي",X{r}="إيجابي",Y{r}="إيجابي"),'
            f'"Research add",AND(W{r}="سلبي",X{r}="سلبي",Y{r}="سلبي"),"Review thesis / risk",'
            f'TRUE,"Hold / review"))', font=F_VAL, border=B_ALL)
        put(ws, r, 28,
            f'=TEXTJOIN("; ",TRUE,IF(U{r}<>"Compliant","Sharia="&U{r},""),'
            f'IF(Z{r}<>"YES","Sharia evidence/review missing",""),IF(AD{r}<>"Fresh","Data stale/unknown",""),IF(T{r}<Checks!$B$7,"Data incomplete",""),'
            f'IF(OR(Checks!$B$9="",Checks!$B$10="",Checks!$B$11=""),"Risk limits pending",""),'
            f'IF(IF(Q{r}="SAR",Checks!$B$13,Checks!$B$18)="","Cash pending",""))',
            font=F_VAL, border=B_ALL)
        put(ws, r, 31, f'=IF(OR(D{r}="",E{r}=""),"",D{r}*E{r})',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 32, f'=IF(Q{r}="SAR",H{r},"")', font=F_VAL,
            fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 33, f'=IF(OR(AE{r}="",AF{r}=""),"",AE{r}-AF{r})',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 34,
            f'=IF(B{r}="","",IF(Q{r}="SAR",0,"Pending transaction FX history"))',
            font=F_VAL, border=B_ALL)
        put(ws, r, 35,
            f'=IFERROR(VLOOKUP(B{r},\'Risk & Horizons\'!$A:$AU,34,FALSE),"")',
            font=F_VAL, border=B_ALL)
        put(ws, r, 36,
            f'=IFERROR(VLOOKUP(B{r},\'Risk & Horizons\'!$A:$AU,3,FALSE),"")',
            font=F_VAL, border=B_ALL)

    # ---- TOTAL row
    put(ws, PF_TOTAL_ROW, 2, "الإجمالي", font=F_TOT, fill=FILL_TOT, border=B_ALL)
    for col in list(range(3, 7)) + list(range(12, 37)):
        put(ws, PF_TOTAL_ROW, col, None, fill=FILL_TOT, border=B_ALL)
    put(ws, PF_TOTAL_ROW, 7, f"=SUM(G{PF_FIRST_ROW}:G{PF_LAST_ROW})",
        font=F_TOT, fill=FILL_TOT, fmt=FMT_MONEY, border=B_ALL)
    put(ws, PF_TOTAL_ROW, 8, f"=SUM(H{PF_FIRST_ROW}:H{PF_LAST_ROW})",
        font=F_TOT, fill=FILL_TOT, fmt=FMT_MONEY, border=B_ALL)
    put(ws, PF_TOTAL_ROW, 9, f"=SUM(I{PF_FIRST_ROW}:I{PF_LAST_ROW})",
        font=F_TOT, fill=FILL_TOT, fmt=FMT_MONEY, border=B_ALL)
    put(ws, PF_TOTAL_ROW, 10, f'=IFERROR(I{PF_TOTAL_ROW}/H{PF_TOTAL_ROW},"")',
        font=F_TOT, fill=FILL_TOT, fmt=FMT_PCT, border=B_ALL)
    put(ws, PF_TOTAL_ROW, 11, 1, font=F_TOT, fill=FILL_TOT, fmt=FMT_PCT, border=B_ALL)

    # ---- summary block: labels in B, values in E
    merge_band(ws, PF_SUM_TITLE_ROW, 2, 5, "ملخص المحفظة")
    first, last = PF_FIRST_ROW, PF_LAST_ROW
    summary = [
        (PF_SUM_COUNT_ROW, "عدد الأسهم", f"=COUNT(G{first}:G{last})", FMT_INT),
        (PF_SUM_PL_ROW, "إجمالي الربح/الخسارة", f"=I{PF_TOTAL_ROW}", FMT_MONEY),
        (PF_SUM_MAXW_ROW, "أعلى وزن سهم", f"=MAX(K{first}:K{last})", FMT_PCT),
        (PF_SUM_CONC_ROW, "حالة التركيز",
         f'=IF(OR(Checks!$B$9="",Checks!$B$10=""),"الحدود معلّقة",'
         f'IF(E{PF_SUM_MAXW_ROW}>=Checks!$B$10,"فوق الحد الصارم",'
         f'IF(E{PF_SUM_MAXW_ROW}>=Checks!$B$9,"فوق الحد المرن","ضمن الحدود")))', None),
        (PF_SUM_R1_ROW, "عدد «شراء قوي»", f'=COUNTIF(M{first}:M{last},"شراء قوي")', FMT_INT),
        (PF_SUM_R2_ROW, "عدد «شراء»", f'=COUNTIF(M{first}:M{last},"شراء")', FMT_INT),
        (PF_SUM_R3_ROW, "عدد «تعزيز/احتفاظ»",
         f'=COUNTIF(M{first}:M{last},"تعزيز/احتفاظ")', FMT_INT),
        (PF_SUM_R4_ROW, "عدد «بيع» و«بيع قوي»",
         f'=COUNTIF(M{first}:M{last},"بيع")+COUNTIF(M{first}:M{last},"بيع قوي")', FMT_INT),
    ]
    for row, label, formula, fmt in summary:
        put(ws, row, 2, label, font=F_LBL, align=RIGHT)
        put(ws, row, 5, formula, font=F_VAL, fmt=fmt)
    put(ws, PF_SUM_SECTOR_ROW, 2, "توزيع الصناعات", font=F_LBL, align=RIGHT)
    put(ws, PF_SUM_SECTOR_ROW, 5,
        "راجع عمودي «القطاع» و«بيتا النفط» في ورقة Risk & Horizons لموازنة القطاعات",
        font=F_NOTE)

    set_widths(ws, {
        "A": 3, "B": 11, "C": 30, "D": 11, "E": 12, "F": 14, "G": 16,
        "H": 16, "I": 15, "J": 13, "K": 10, "L": 10, "M": 14, "N": 22,
        "O": 15, "P": 12, "Q": 10, "R": 11, "S": 13, "T": 14, "U": 16,
        "V": 16, "W": 15, "X": 15, "Y": 15, "Z": 14, "AA": 25,
        "AB": 44, "AC": 40, "AD": 20,
        "AE": 24, "AF": 24, "AG": 24, "AH": 28, "AI": 16, "AJ": 22,
    })
    set_filter(ws, PF_HDR_ROW, PF_LAST_ROW, 1 + len(headers))
    ws.freeze_panes = "A6"      # rows 1-5 stay visible
    return ws


# --------------------------------------------------------------------------
# Stock Lookup
# --------------------------------------------------------------------------

def sheet_lookup(wb, rows):
    ws = wb.create_sheet("Stock Lookup")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.rightToLeft = True

    put(ws, SL_TITLE_ROW, 2, "بطاقة السهم — اكتب رمز السهم واقرأ التقرير", font=F_TITLE)

    put(ws, SL_INPUT_ROW, 1, "رمز السهم:", font=F_LBL_B, align=RIGHT)
    seed = rows[0]["code"] if rows else None
    put(ws, SL_INPUT_ROW, 2, seed, font=F_IN, fill=FILL_IN, fmt=FMT_TEXT,
        align=CENTER, border=B_IN)
    put(ws, SL_INPUT_ROW, 3,
        f"=IF($B${SL_INPUT_ROW}=\"\",\"\",IFERROR(VLOOKUP($B${SL_INPUT_ROW},"
        f"'Risk & Horizons'!$A:$C,2,0),\"\"))", font=F_LBL_B)
    put(ws, SL_INPUT_ROW, 4,
        f"=IF($B${SL_INPUT_ROW}=\"\",\"\",IFERROR(VLOOKUP($B${SL_INPUT_ROW},"
        f"'Risk & Horizons'!$A:$C,3,0),\"\"))", font=F_NOTE)

    put(ws, SL_GUARD_ROW, 2,
        f'=IF(LEN(B{SL_INPUT_ROW})=0,"",IF(LEN(B{SL_INPUT_ROW})<>4,'
        f'"أدخل رمزاً من 4 أرقام",IF(ISNA(MATCH(B{SL_INPUT_ROW},DB!$A:$A,0)),'
        f'"رمز غير موجود في قاعدة البيانات — حدّث القاعدة أولاً (refresh.bat)","")))',
        font=F_WARN)
    put(ws, SL_MEMBER_ROW, 2,
        f'=IF(COUNTIF(Portfolio!$B${PF_FIRST_ROW}:$B${PF_LAST_ROW},B{SL_INPUT_ROW})>0,'
        f'"✔ في محفظتك","غير موجود في محفظتك")', font=F_LBL_B)

    def rh(col_idx, fmt=None):
        """VLOOKUP into Risk & Horizons (A..S), 1-based column index."""
        return (f"=IF($B${SL_INPUT_ROW}=\"\",\"\",IFERROR(VLOOKUP($B${SL_INPUT_ROW},"
                f"'Risk & Horizons'!$A:$S,{col_idx},0),\"{NA}\"))"), fmt

    def db_latest(col_letter, fmt=None):
        """Latest-year DB row for the typed symbol, via the Q helper column.

        The INDEX is evaluated twice on purpose: INDEX into an EMPTY cell yields
        0, not "", so an unpopulated field would otherwise read as a confident
        zero. Testing the result first turns that into the missing-data marker.
        """
        idx = (f'INDEX(DB!${col_letter}:${col_letter},'
               f'MATCH($B${SL_INPUT_ROW},DB!$Q:$Q,0))')
        return (f'=IF($B${SL_INPUT_ROW}="","",IFERROR(IF({idx}="","{NA}",{idx}),"{NA}"))'), fmt

    def section(hdr_row, first_row, title, items):
        merge_band(ws, hdr_row, 2, 4, title)
        for i, (label, (formula, fmt)) in enumerate(items):
            r = first_row + i
            put(ws, r, 2, label, font=F_LBL, align=RIGHT, border=B_ALL)
            put(ws, r, 3, formula, font=F_VAL, fmt=fmt, align=CENTER, border=B_ALL)

    section(SL_PRICE_HDR_ROW, SL_PRICE_FIRST_ROW, "السعر والعائد", [
        ("السعر (ريال)", rh(4, FMT_MONEY)),
        ("عائد أسبوع", rh(5, FMT_PCT)),
        ("عائد شهر", rh(6, FMT_PCT)),
        ("عائد 3 أشهر", rh(7, FMT_PCT)),
        ("عائد 6 أشهر", rh(8, FMT_PCT)),
        ("عائد سنة", rh(9, FMT_PCT)),
        # YTD = the current-year row in DB (col D holds that year's return so far)
        ("عائد منذ بداية السنة", db_latest("D", FMT_PCT)),
    ])

    section(SL_RISK_HDR_ROW, SL_RISK_FIRST_ROW, "المخاطر", [
        ("مؤشر القوة النسبية RSI(14)", rh(10, FMT_NUM1)),
        ("نظام التقلب", rh(11)),
        ("السعر مقابل متوسط 200 يوم", rh(12)),
        ("أقصى تراجع خلال سنتين", rh(13, FMT_PCT)),
    ])

    section(SL_FUND_HDR_ROW, SL_FUND_FIRST_ROW, "الأساسيات", [
        ("مكرر الربحية P/E", db_latest("K", FMT_NUM1)),
        ("العائد على حقوق الملكية ROE", db_latest("L", FMT_PCTV)),
        # DB col H is the per-year series, so the latest row is the CURRENT
        # (part-)year yield. The score's dividend axis uses the 5-year average
        # instead, hence the explicit label - the two numbers differ on purpose.
        ("عائد التوزيعات (السنة الحالية)", db_latest("H", FMT_PCTV)),
        ("نسبة التوزيع", db_latest("R", FMT_PCT)),
    ])

    # ---- five-year table: revenue/NI/EPS from Statements, ROE per year from DB
    header_row(ws, SL_YEARS_HDR_ROW, 2,
               ["السنة", "الإيرادات (مليون ريال)", "صافي الربح (مليون ريال)",
                "ربحية السهم", "ROE"])
    for i, year in enumerate(sorted(STMT_YEARS)):
        r = SL_YEARS_FIRST_ROW + i
        base = 3 + STMT_YEARS.index(year) * 3    # Statements C/F/I/L/O per year group
        put(ws, r, 2, year, font=F_LBL, align=CENTER, border=B_ALL)
        for j in range(3):
            # Doubled VLOOKUP for the same reason as db_latest: a blank source
            # cell (year not filed) must show blank, not 0.
            vl = f'VLOOKUP($B${SL_INPUT_ROW},Statements!$A:$R,{base + j},0)'
            put(ws, r, 3 + j,
                f'=IF($B${SL_INPUT_ROW}="","",IFERROR(IF({vl}="","",{vl}),""))',
                font=F_VAL, fmt=(FMT_INT if j < 2 else FMT_NUM2),
                align=CENTER, border=B_ALL)
        put(ws, r, 6,
            f'=IF($B${SL_INPUT_ROW}="","",IFERROR(INDEX(DB!$L:$L,'
            f'MATCH($B${SL_INPUT_ROW}&"|{year}",DB!$P:$P,0)),"{NA}"))',
            font=F_VAL, fmt=FMT_PCTV, align=CENTER, border=B_ALL)

    section(SL_VERDICT_HDR_ROW, SL_VERDICT_FIRST_ROW, "الحكم", [
        ("النتيجة المركبة (0-100)", rh(18, FMT_NUM1)),
        ("التقييم", rh(19)),
        ("حكم قصير (6 أشهر)", rh(15)),
        ("حكم متوسط (سنتان)", rh(16)),
        ("حكم طويل (5 سنوات)", rh(17)),
    ])
    note_row = SL_VERDICT_FIRST_ROW + 5
    put(ws, note_row, 2, "ملاحظة", font=F_LBL, align=RIGHT, border=B_ALL)
    put(ws, note_row, 3,
        f'=IF(COUNTIF(DB!$O:$O,$B${SL_INPUT_ROW}&"|ناقص")>0,'
        f'"⚠ بعض البيانات ناقصة — الحكم محايد للأجزاء الناقصة","")',
        font=F_NOTE, border=B_ALL)

    set_widths(ws, {"A": 14, "B": 30, "C": 24, "D": 24, "E": 16, "F": 14})
    return ws


# --------------------------------------------------------------------------
# Risk & Horizons
# --------------------------------------------------------------------------

def sheet_risk(wb, rows):
    ws = wb.create_sheet("Risk & Horizons")
    ws.sheet_view.showGridLines = False
    put(ws, RH_TITLE_ROW, 1, "المخاطر والآفاق الزمنية — كل الرموز", font=F_TITLE)

    headers = [
        "الرمز", "الشركة", "القطاع", "السعر",
        "عائد أسبوع", "عائد شهر", "عائد 3 أشهر", "عائد 6 أشهر", "عائد سنة",
        "RSI(14)", "نظام التقلب", "مقابل SMA200", "أقصى تراجع سنتين",
        "بيتا النفط", "6 أشهر", "سنتان", "5 سنوات", "النتيجة", "التقييم",
        "اتجاه التقلب", "اكتمال البيانات", "قابل للإجراء", "تاريخ السعر",
        "المصدر", "رابط المصدر", "المعرف", "السوق", "العملة", "FX إلى SAR", "حالة الحداثة",
        "تقويم السوق", "التسوية", "فترة القوائم", "نوع الأداة", "توقيت السوق",
        "جلسة السوق", "معالجة التوقيت الصيفي", "قاعدة الكمية", "قواعد الأوامر",
        "أساس القوائم", "تاريخ النشر", "وقت الملاحظة", "وقت الجلب",
        "أساس السعر", "أساس العائد", "وحدة التوزيعات", "مصدر قواعد السوق",
    ]
    header_row(ws, RH_HDR_ROW, 1, headers)

    fmts = {
        4: FMT_MONEY, 5: FMT_PCT, 6: FMT_PCT, 7: FMT_PCT, 8: FMT_PCT, 9: FMT_PCT,
        10: FMT_NUM1, 13: FMT_PCT, 18: FMT_NUM1, 21: FMT_PCT, 29: FMT_NUM2,
    }
    for i, rec in enumerate(rows):
        r = RH_FIRST_ROW + i
        values = [
            rec["code"], rec["name"], rec["sector"] or NA, rec["price"],
            rec["ret_1w"], rec["ret_1m"], rec["ret_3m"], rec["ret_6m"], rec["ret_1y"],
            rec["rsi"], rec["vol_regime"], rec["sma200"], rec["maxdd_2y"],
            rec["oil_beta"], rec["near"], rec["mid"], rec["far"],
            rec["score"], rec["rating"], rec["vol_trend"], rec["completeness"],
            "نعم" if rec["actionable"] else "لا", rec["data_as_of"], rec["source"],
            rec["source_url"], rec["security_id"], rec["exchange"], rec["currency"],
            rec["fx_to_sar"], rec["freshness"], rec["market_calendar"],
            rec["settlement"], rec["fundamentals_period"], rec["instrument_type"],
            rec["market_timezone"], rec["market_session"], rec["dst_handling"],
            rec["quantity_rule"], rec["order_rule"], rec["fundamentals_basis"],
            rec["publication_date"], rec["observation_time"], rec["retrieval_time"],
            rec["price_basis"], rec["return_basis"], rec["dividend_unit"],
            rec["market_rule_source"],
        ]
        for col, value in enumerate(values, start=1):
            put(ws, r, col, value, font=F_STAT, border=B_ALL,
                fmt=(FMT_TEXT if col == 1 else fmts.get(col)),
                align=(LEFT if col == 2 else CENTER))

    set_widths(ws, {
        "A": 9, "B": 28, "C": 22, "D": 10, "E": 11, "F": 11, "G": 11, "H": 11,
        "I": 11, "J": 9, "K": 12, "L": 13, "M": 15, "N": 11, "O": 18, "P": 18,
        "Q": 18, "R": 9, "S": 14, "T": 12, "U": 14, "V": 12,
        "W": 13, "X": 18, "Y": 36, "Z": 14, "AA": 12, "AB": 10,
        "AC": 10, "AD": 20, "AE": 18, "AF": 12, "AG": 14, "AH": 14,
        "AI": 20, "AJ": 28, "AK": 28, "AL": 32, "AM": 36, "AN": 34,
        "AO": 18, "AP": 20, "AQ": 24, "AR": 34, "AS": 34, "AT": 22,
        "AU": 42,
    })
    set_filter(ws, RH_HDR_ROW, RH_FIRST_ROW + len(rows) - 1, len(headers))
    ws.freeze_panes = "A5"
    return ws


# --------------------------------------------------------------------------
# DB
# --------------------------------------------------------------------------

def sheet_db(wb, data, rows):
    ws = wb.create_sheet("DB")
    ws.sheet_view.showGridLines = False
    put(ws, DB_TITLE_ROW, 1,
        "قاعدة البيانات — صف لكل (رمز، سنة). الرموز تصاعدياً، والسنوات تنازلياً "
        "داخل كل رمز. العائد من إغلاق معدل والتوزيعات نقد/سهم بعملة الإدراج.", font=F_TITLE)

    headers = [
        "الرمز", "السنة", "الإغلاق المعدل", "العائد الإجمالي السنوي", "التقلب السنوي",
        "أقصى تراجع", "التوزيعات (عملة الإدراج/سهم)", "عائد التوزيعات", "زخم 12-1 (آخر شهر مستبعد)",
        "ربحية السهم", "مكرر الربحية", "ROE", "النتيجة", "التقييم",
        "مفتاح النقص", "مفتاح السنة", "مفتاح الأحدث", "نسبة التوزيع",
    ]
    header_row(ws, DB_HDR_ROW, 1, headers)

    by_code = {rec["code"]: rec for rec in rows}
    r = DB_FIRST_ROW
    for code in sorted(data["annual"]):
        years = data["annual"][code]
        rec = by_code.get(code) or {}
        stmt_years = data["statements"].get(code) or {}
        ly = latest_year(years)
        for year in sorted(years, reverse=True):   # newest first: see module docstring
            row = years[year]
            stmt = stmt_years.get(year) or {}
            is_latest = (year == ly)
            close = row["close"]
            eps = row["eps"] if row["eps"] is not None else stmt.get("eps")
            pe = pe_value(close, eps)
            roe = stmt.get("roe")
            payout = stmt.get("payout")
            if is_latest:
                # The latest row is "today", and the current year is usually still
                # running: no full-year EPS and no filed statements yet. Fall back
                # to the snapshot's trailing figures so Stock Lookup's الأساسيات
                # block shows current fundamentals instead of a hole.
                if pe == NA and rec.get("pe"):
                    pe = round(rec["pe"], 1)
                if roe is None:
                    roe = rec.get("roe")
                if payout is None:
                    payout = rec.get("payout")
            # The na key marks symbols whose SCORING inputs are incomplete, so it
            # lives on the latest row only; old years legitimately lack EPS/ROE
            # and must not raise the warning.
            missing = is_latest and rec.get("missing_inputs", False)
            put(ws, r, 1, code, font=F_STAT, fmt=FMT_TEXT, border=B_ALL, align=CENTER)
            put(ws, r, 2, year, font=F_STAT, fmt=FMT_INT, border=B_ALL, align=CENTER)
            put(ws, r, 3, close, font=F_STAT, fmt=FMT_MONEY, border=B_ALL)
            put(ws, r, 4, row["ret"], font=F_STAT, fmt=FMT_PCT, border=B_ALL)
            put(ws, r, 5, row["vol"], font=F_STAT, fmt=FMT_PCT, border=B_ALL)
            put(ws, r, 6, row["maxdd"], font=F_STAT, fmt=FMT_PCT, border=B_ALL)
            put(ws, r, 7, row["divs"], font=F_STAT, fmt=FMT_MONEY, border=B_ALL)
            put(ws, r, 8, row["div_yield"], font=F_STAT, fmt=FMT_PCTV, border=B_ALL)
            put(ws, r, 9, row["momentum"], font=F_STAT, fmt=FMT_PCT, border=B_ALL)
            put(ws, r, 10, eps, font=F_STAT, fmt=FMT_NUM2, border=B_ALL)
            put(ws, r, 11, pe, font=F_STAT, border=B_ALL,
                fmt=(None if isinstance(pe, str) else FMT_NUM1))
            put(ws, r, 12, NA if roe is None else roe, font=F_STAT, border=B_ALL,
                fmt=(None if roe is None else FMT_PCTV))
            put(ws, r, 13, rec.get("score") if is_latest else None,
                font=F_STAT, fmt=FMT_NUM1, border=B_ALL)
            put(ws, r, 14, rec.get("rating") if is_latest else None,
                font=F_STAT, border=B_ALL)
            put(ws, r, 15, ("%s|ناقص" % code) if missing else None,
                font=F_NOTE, fmt=FMT_TEXT, border=B_ALL)
            put(ws, r, 16, "%s|%d" % (code, year), font=F_NOTE, fmt=FMT_TEXT, border=B_ALL)
            put(ws, r, 17, code if is_latest else None,
                font=F_NOTE, fmt=FMT_TEXT, border=B_ALL)
            put(ws, r, 18, payout if is_latest else None,
                font=F_STAT, fmt=FMT_PCT, border=B_ALL)
            r += 1

    set_widths(ws, {
        "A": 9, "B": 8, "C": 11, "D": 12, "E": 12, "F": 12, "G": 14, "H": 14,
        "I": 11, "J": 11, "K": 13, "L": 12, "M": 9, "N": 14, "O": 14, "P": 14,
        "Q": 12, "R": 12,
    })
    set_filter(ws, DB_HDR_ROW, r - 1, len(headers))
    ws.freeze_panes = "A5"
    return ws


# --------------------------------------------------------------------------
# Statements
# --------------------------------------------------------------------------

def _stmt_value(value):
    """Statement cell value: an exact 0 means Yahoo had no data -> missing."""
    if value is None:
        return None
    try:
        if float(value) == 0.0:
            return None
    except (TypeError, ValueError):
        return None
    return value


def sheet_statements(wb, data, rows):
    ws = wb.create_sheet("Statements")
    ws.sheet_view.showGridLines = False
    put(ws, ST_TITLE_ROW, 1,
        "القوائم المالية — الإيرادات وصافي الربح (بالمليون ريال) وربحية السهم، "
        "%d-%d" % (min(STMT_YEARS), max(STMT_YEARS)), font=F_TITLE)

    for i, year in enumerate(STMT_YEARS):
        first_col = 3 + i * 3
        merge_band(ws, ST_GROUP_ROW, first_col, first_col + 2, str(year))

    headers = ["الرمز", "الشركة"]
    for _ in STMT_YEARS:
        headers += ["الإيرادات", "صافي الربح", "ربحية السهم"]
    headers.append("ملاحظة")
    header_row(ws, ST_HDR_ROW, 1, headers)

    note_col = 3 + len(STMT_YEARS) * 3
    for i, rec in enumerate(rows):
        r = ST_FIRST_ROW + i
        years = data["statements"].get(rec["code"]) or {}
        put(ws, r, 1, rec["code"], font=F_STAT, fmt=FMT_TEXT, border=B_ALL, align=CENTER)
        put(ws, r, 2, rec["name"], font=F_STAT, border=B_ALL, align=LEFT)
        missing_years = []
        for j, year in enumerate(STMT_YEARS):
            col = 3 + j * 3
            stmt = years.get(year) or {}
            # Yahoo reports "no data" as a hard 0 on some symbols, so an exact
            # 0 is treated as missing rather than written as a real zero. A
            # genuine negative (some investment/insurance firms) is kept.
            revenue = _stmt_value(stmt.get("revenue"))
            net_income = _stmt_value(stmt.get("net_income"))
            eps = _stmt_value(stmt.get("eps"))
            if revenue is None and net_income is None and eps is None:
                missing_years.append(str(year))
            put(ws, r, col, None if revenue is None else revenue / 1e6,
                font=F_STAT, fmt=FMT_INT, border=B_ALL)
            put(ws, r, col + 1, None if net_income is None else net_income / 1e6,
                font=F_STAT, fmt=FMT_INT, border=B_ALL)
            put(ws, r, col + 2, eps, font=F_STAT, fmt=FMT_NUM2, border=B_ALL)
        note = ("بيانات ناقصة: " + "، ".join(missing_years)) if missing_years else None
        put(ws, r, note_col, note, font=F_NOTE, border=B_ALL)

    widths = {"A": 9, "B": 28}
    for i in range(len(STMT_YEARS) * 3):
        widths[get_column_letter(3 + i)] = 14
    widths[get_column_letter(note_col)] = 28
    set_widths(ws, widths)
    set_filter(ws, ST_HDR_ROW, ST_FIRST_ROW + len(rows) - 1, note_col)
    ws.freeze_panes = "C5"
    return ws


# --------------------------------------------------------------------------
# Symbols
# --------------------------------------------------------------------------

def sheet_symbols(wb, rows):
    ws = wb.create_sheet("Symbols")
    ws.sheet_view.showGridLines = False
    put(ws, SY_TITLE_ROW, 1, "قائمة الرموز", font=F_TITLE)
    header_row(ws, SY_HDR_ROW, 1,
               ["الرمز", "الاسم بالعربية", "الاسم بالإنجليزية", "القطاع",
                "في قاعدة البيانات"])

    for i, rec in enumerate(sorted(rows, key=lambda x: x["code"])):
        r = SY_FIRST_ROW + i
        put(ws, r, 1, rec["code"], font=F_STAT, fmt=FMT_TEXT, border=B_ALL, align=CENTER)
        put(ws, r, 2, rec["name_ar"] or rec["name"], font=F_STAT, border=B_ALL)
        put(ws, r, 3, rec["name_en"] or "", font=F_STAT, border=B_ALL)
        put(ws, r, 4, rec["sector"] or NA, font=F_STAT, border=B_ALL)
        # ISNA(MATCH(...)) rather than COUNTIF: codes are stored as text and
        # COUNTIF's implicit text/number coercion is not portable across
        # Excel and LibreOffice Calc, while MATCH exact-match is.
        put(ws, r, 5, f'=IF(ISNA(MATCH(A{r},DB!$A:$A,0)),"✗","✔")',
            font=F_VAL, border=B_ALL, align=CENTER)

    set_widths(ws, {"A": 10, "B": 30, "C": 38, "D": 24, "E": 18})
    set_filter(ws, SY_HDR_ROW, SY_FIRST_ROW + len(rows) - 1, 5)
    ws.freeze_panes = "A5"
    return ws


# --------------------------------------------------------------------------
# Personal-book control sheets
# --------------------------------------------------------------------------

def _list_validation(ws, cell_range, values):
    validation = DataValidation(type="list", formula1='"%s"' % ",".join(values),
                                allow_blank=True)
    validation.error = "اختر قيمة من القائمة"
    validation.errorTitle = "قيمة غير صالحة"
    validation.prompt = "استخدم القائمة المنسدلة للحفاظ على سجل قابل للتدقيق"
    validation.promptTitle = "Waraqah"
    validation.showErrorMessage = True
    validation.showInputMessage = True
    ws.add_data_validation(validation)
    validation.add(cell_range)


def sheet_activity(wb, holdings):
    """Auditable transaction ledger; Portfolio positions derive from this sheet."""
    ws = wb.create_sheet("Activity")
    ws.sheet_view.showGridLines = False
    put(ws, 2, 1, "سجل النشاط — المصدر المحاسبي الوحيد للمراكز والنقد", font=F_TITLE)
    put(ws, 3, 1,
        "أدخل كل صف مرة واحدة بمعرّف ثابت. أرصدة الافتتاح تحفظ المراكز الحالية فقط؛ "
        "تواريخ الشراء الأصلية تظل معلّقة حتى استيراد كشف الوسيط.", font=F_NOTE)
    headers = [
        "Transaction ID", "Trade Date", "Settlement Date", "Account", "Broker",
        "Security ID", "Ticker", "Exchange", "Type", "Quantity", "Price",
        "Currency", "FX to SAR", "Fees", "Tax / Withholding", "Cash Amount", "Broker Reference",
        "Signed Qty", "Qty Before", "Cost Basis Before SAR", "Cost Basis Change SAR",
        "Realized P/L SAR", "Cash Change Original", "Cash Change SAR", "Source",
        "Imported At", "Duplicate?", "Notes",
    ]
    header_row(ws, 5, 1, headers)

    for idx, held in enumerate(holdings):
        r = 6 + idx
        code, shares, cost = held
        values = [
            "OPEN-SA-%s-20260909" % code, None, None, "Pending", "Pending",
            "SA-%s" % code, code, "Tadawul", "Opening balance", shares, cost,
            "SAR", 1.0, 0.0, 0.0, None, None,
        ]
        for col, value in enumerate(values, 1):
            put(ws, r, col, value, font=F_IN if col <= 17 else F_VAL,
                fill=FILL_IN if col <= 17 else None, border=B_ALL,
                fmt=FMT_TEXT if col in (1, 6, 7, 12, 17) else None,
                align=CENTER)
        put(ws, r, 25, "Original Portfolio opening balance; acquisition date unknown",
            font=F_NOTE, border=B_ALL)
        put(ws, r, 26, "2026-09-09", font=F_VAL, border=B_ALL)
        put(ws, r, 28, "Book opening snapshot — not an asserted trade date",
            font=F_NOTE, border=B_ALL)

    for r in range(6, 506):
        put(ws, r, 18,
            f'=IF(F{r}="","",IF(OR(I{r}="Opening balance",I{r}="Buy",I{r}="Transfer in"),J{r},'
            f'IF(OR(I{r}="Sell",I{r}="Transfer out"),-J{r},IF(I{r}="Split",S{r}*(J{r}-1),0))))',
            font=F_VAL, fmt=FMT_NUM2, border=B_ALL)
        before_qty = "0" if r == 6 else f'SUMIF($F$6:F{r-1},F{r},$R$6:R{r-1})'
        before_cost = "0" if r == 6 else f'SUMIF($F$6:F{r-1},F{r},$U$6:U{r-1})'
        put(ws, r, 19, f'=IF(F{r}="","",{before_qty})', font=F_VAL,
            fmt=FMT_NUM2, border=B_ALL)
        put(ws, r, 20, f'=IF(F{r}="","",{before_cost})', font=F_VAL,
            fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 21,
            f'=IF(F{r}="","",IF(OR(I{r}="Opening balance",I{r}="Buy",I{r}="Transfer in"),'
            f'(J{r}*K{r}+N{r}+O{r})*M{r},IF(OR(I{r}="Sell",I{r}="Transfer out"),'
            f'-IFERROR(J{r}*T{r}/S{r},0),0)))', font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 22,
            f'=IF(A{r}="","",IF(I{r}="Sell",(J{r}*K{r}-N{r}-O{r})*M{r}+U{r},0))',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 23,
            f'=IF(A{r}="","",IF(I{r}="Buy",-(J{r}*K{r}+N{r}+O{r}),'
            f'IF(I{r}="Sell",J{r}*K{r}-N{r}-O{r},IF(OR(I{r}="Deposit",I{r}="Dividend"),'
            f'P{r}-N{r}-O{r},IF(I{r}="FX buy",P{r}-N{r}-O{r},'
            f'IF(I{r}="Corporate action",P{r}-N{r}-O{r},'
            f'IF(OR(I{r}="Withdrawal",I{r}="Fee",I{r}="Tax",I{r}="Withholding",I{r}="FX sell"),'
            f'-(P{r}+N{r}+O{r}),0)))))))', font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 24, f'=IF(W{r}="","",W{r}*M{r})', font=F_VAL,
            fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 27, f'=IF(A{r}="","",IF(COUNTIF($A$6:$A$505,A{r})>1,"DUPLICATE","OK"))',
            font=F_VAL, border=B_ALL, align=CENTER)

    _list_validation(ws, "H6:H505", ["Tadawul", "NYSE", "NASDAQ", "Other"])
    _list_validation(ws, "I6:I505", [
        "Opening balance", "Buy", "Sell", "Dividend", "Fee", "Tax", "Withholding", "Deposit",
        "Withdrawal", "Transfer in", "Transfer out", "Split", "Corporate action",
        "FX buy", "FX sell",
    ])
    _list_validation(ws, "L6:L505", ["SAR", "USD"])
    widths = {get_column_letter(i): 14 for i in range(1, len(headers) + 1)}
    widths.update({"A": 24, "D": 16, "E": 16, "F": 15, "I": 18, "Q": 20,
                   "Y": 35, "Z": 18, "AB": 38})
    set_widths(ws, widths)
    set_filter(ws, 5, 505, len(headers))
    ws.freeze_panes = "A6"
    return ws


def sheet_sharia(wb, holdings, rows):
    """Evidence ledger and hard eligibility gate for new purchases."""
    ws = wb.create_sheet("Sharia")
    ws.sheet_view.showGridLines = False
    put(ws, 2, 1, "الضوابط الشرعية — لا شراء من دون دليل ساري", font=F_TITLE)
    put(ws, 3, 1,
        "المنهجية/الجهة ما زالت معلّقة بانتظار اختيار المالك. لذلك تبدأ جميع المراكز "
        "بحالة Uncertain ويُحظر أي اقتراح شراء تلقائياً.", font=F_NOTE)
    headers = [
        "Security ID", "Ticker", "Company", "Exchange", "Status",
        "Authority / Provider", "Methodology", "Methodology Version", "Evidence URL",
        "Reporting Period", "Screen Date", "Next Review", "Business Activity Evidence",
        "Financial Ratio Evidence", "Purification Method", "Purification Due SAR",
        "Purification Paid SAR", "Payment Date", "Change Since Prior Review", "Reviewer",
        "Notes", "Buy Eligible?", "Holding Review Flag",
    ]
    header_row(ws, 5, 1, headers)
    by_code = {row["code"]: row for row in rows}
    for idx, held in enumerate(holdings):
        r = 6 + idx
        code = held[0]
        rec = by_code.get(code) or {}
        values = [
            "SA-" + code, code, rec.get("name") or code, "Tadawul", "Uncertain",
            "Pending owner selection", "Pending owner selection",
            "Pending owner selection", None, None, None, None, "Pending evidence",
            "Pending evidence", "Pending methodology", None, None, None,
            "No prior verified screen", "Codex",
            "Initial control row; no Sharia conclusion made",
        ]
        for col, value in enumerate(values, 1):
            put(ws, r, col, value, font=F_IN if col in range(5, 22) else F_VAL,
                fill=FILL_IN if col in range(5, 22) else None, border=B_ALL,
                fmt=(FMT_TEXT if col in (1, 2, 9) else
                     (FMT_MONEY if col in (16, 17) else None)), align=CENTER)
    for r in range(6, 506):
        put(ws, r, 22,
            f'=IF(A{r}="","",IF(AND(E{r}="Compliant",F{r}<>"",G{r}<>"",H{r}<>"",'
            f'I{r}<>"",J{r}<>"",K{r}<>"",L{r}>=TODAY(),M{r}<>"",N{r}<>""),"YES","NO"))',
            font=F_VAL, border=B_ALL, align=CENTER)
        put(ws, r, 23,
            f'=IF(A{r}="","",IF(OR(E{r}="Non-compliant",E{r}="Review overdue"),'
            f'"Review required — no disposal instruction inferred",'
            f'IF(E{r}="Uncertain","Evidence required",'
            f'IF(AND(S{r}<>"",S{r}<>"No change"),"Change review required","Current"))))',
            font=F_VAL, border=B_ALL, align=CENTER)
    _list_validation(ws, "E6:E505", ["Compliant", "Non-compliant", "Uncertain", "Review overdue"])
    widths = {get_column_letter(i): 16 for i in range(1, len(headers) + 1)}
    widths.update({
        "A": 15, "C": 30, "F": 24, "G": 24, "H": 20, "I": 42, "J": 20,
        "M": 34, "N": 34, "O": 30, "S": 28, "U": 42, "W": 24,
    })
    set_widths(ws, widths)
    set_filter(ws, 5, 505, len(headers))
    ws.freeze_panes = "A6"
    return ws


def _metric_text(value, percent=False):
    if value is None:
        return "N/A"
    if percent:
        return "%.1f%%" % (float(value) * 100.0)
    return "%.2f" % float(value)


def _metric_percentage_points(value):
    """Format a metric already stored in percentage points without `N/A%`."""
    if value is None:
        return "N/A"
    return "%.2f%%" % float(value)


def _review_thesis(rec, horizon):
    """Auditable quantitative review text; intentionally makes no forecast."""
    if horizon == "6 months":
        return (
            "6-month view=%s; 3m return=%s; 6m return=%s; RSI=%s; "
            "volatility level=%s and trend=%s; SMA200=%s. Descriptive only; "
            "no trade until Sharia evidence and owner controls pass."
        ) % (rec.get("near") or "N/A", _metric_text(rec.get("ret_3m"), True),
             _metric_text(rec.get("ret_6m"), True), _metric_text(rec.get("rsi")),
             rec.get("vol_regime") or "N/A", rec.get("vol_trend") or "N/A",
             rec.get("sma200") or "N/A")
    if horizon == "2 years":
        return (
            "2-year view=%s; 1y total return=%s; 12-1 momentum=%s; "
            "max drawdown 2y=%s; P/E=%s; completeness=%s. Evidence refresh "
            "required before any non-Wait proposal."
        ) % (rec.get("mid") or "N/A", _metric_text(rec.get("ret_1y"), True),
             _metric_text(rec.get("momentum"), True), _metric_text(rec.get("maxdd_2y"), True),
             _metric_text(rec.get("pe")), _metric_text(rec.get("completeness"), True))
    return (
        "5-year view=%s; ROE=%s; dividend yield=%s; payout=%s; P/E=%s. "
        "These are latest/trailing inputs, not a five-year forecast; official "
        "filings and Sharia evidence are still required."
    ) % (rec.get("far") or "N/A", _metric_percentage_points(rec.get("roe")),
         _metric_percentage_points(rec.get("div_yield")), _metric_text(rec.get("payout")),
         _metric_text(rec.get("pe")))


def _horizon_framework(horizon):
    """Research lens text that states what remains to be evidenced."""
    if horizon == "6 months":
        return {
            "rationale": (
                "Review the 6-month trend, valuation, trading liquidity and event risk around "
                "verified earnings or corporate announcements; a favorable score alone is not an order."
            ),
            "entry": "Pending verified catalyst, valuation range, liquidity and owner-approved limits",
            "catalyst": "Next earnings or declared corporate action; event and date pending verification",
            "downside": "Bear case must test weak earnings, adverse event risk and a trend break; magnitude pending evidence",
            "triggers": "Re-review after earnings/corporate news, material valuation change, liquidity deterioration or trend break",
        }
    if horizon == "2 years":
        return {
            "rationale": (
                "Review earnings and cash-flow development, debt, execution milestones and sector-appropriate "
                "valuation scenarios over two years."
            ),
            "entry": "Pending verified financial path, debt capacity, milestones and owner-approved limits",
            "catalyst": "Earnings, cash-flow, debt and execution milestones pending verified filings",
            "downside": "Bear case must test missed milestones, weaker cash conversion and balance-sheet stress",
            "triggers": "Re-review on material guidance, debt, cash-flow or execution-milestone changes",
        }
    return {
        "rationale": (
            "Review competitive position, reinvestment, returns on capital, dilution, governance and "
            "long-term valuation sensitivity over five years."
        ),
        "entry": "Pending verified durable economics, governance review and long-term valuation sensitivity",
        "catalyst": "Competitive, reinvestment and governance evidence from future verified filings",
        "downside": "Bear case must test structural competition, poor reinvestment, dilution and governance deterioration",
        "triggers": "Re-review on structural moat, capital-allocation, dilution, governance or return-on-capital changes",
    }


def sheet_orders(wb, holdings, rows):
    """Ranked, versioned proposal queue.  It never sends orders to a broker."""
    ws = wb.create_sheet("Orders")
    ws.sheet_view.showGridLines = False
    put(ws, 2, 1, "اقتراحات الأوامر — للمراجعة البشرية فقط", font=F_TITLE)
    put(ws, 3, 1,
        "Codex يحدّث البحث والمقترحات فقط. لا توجد وصلة تنفيذ، ولمس السعر لا يثبت التعبئة؛ "
        "فقط تعبئة مؤكدة من الوسيط ومعرّف Activity تغيّر السجل الفعلي.", font=F_NOTE)
    headers = [
        "Proposal ID", "Version", "Rank", "Created At", "Reviewed At", "Horizon",
        "Security ID", "Ticker", "Company", "Exchange", "Broker", "Account", "Currency",
        "Action", "Order Type", "Entry Condition", "Proposed Quantity", "Limit Price",
        "Stop / Review Price", "Valid Until", "Thesis / Rationale", "Facts",
        "Estimates / Scenarios", "Agent Judgment", "Evidence URLs", "Catalyst",
        "Downside Scenario", "Review / Exit Triggers", "Why I Own It", "Thesis Invalidation",
        "Current Position %", "Proposed Position %", "Proposed Sector %",
        "Cash Impact Original", "Cash Impact SAR", "Cash After SAR", "Sizing Assumptions",
        "Horizon Reconciliation", "Sharia Status", "Sharia Evidence", "Data Freshness",
        "Blockers", "Confidence", "Status", "User Decision", "Decision Date",
        "Broker Order ID", "Filled Quantity", "Average Fill Price",
        "Execution Transaction ID", "Review ID", "Analysis Version", "Last Checked", "Notes",
        "Duplicate?", "Execution Control", "Horizon Conflict?",
    ]
    header_row(ws, 5, 1, headers)
    by_code = {row["code"]: row for row in rows}
    horizons = ["6 months", "2 years", "5 years"]
    current_row = 6
    rank = 1
    for held in holdings:
        code = held[0]
        rec = by_code.get(code) or {}
        for horizon in horizons:
            r = current_row
            framework = _horizon_framework(horizon)
            proposal = "REV-%s-SA-%s-%s" % (BOOK_BUILD_ID_DATE, code, horizon.split()[0])
            security_id = rec.get("security_id") or ("SA-" + code)
            exchange = rec.get("exchange") or "Tadawul"
            currency = rec.get("currency") or "SAR"
            values = [
                proposal, 1, rank, BOOK_BUILD_DATE, BOOK_BUILD_DATE, horizon,
                security_id, code, rec.get("name") or code, exchange, None, None, currency,
                "Wait", "No order", framework["entry"], None, None, None, None,
                framework["rationale"], _review_thesis(rec, horizon),
                "Bear/base/bull assumptions pending verified forward evidence; no target price invented",
                "Wait — Sharia method/evidence, cash, broker and owner risk limits are unresolved",
                rec.get("source_url"), framework["catalyst"], framework["downside"],
                framework["triggers"], "Pending owner thesis", "Pending owner invalidation criteria",
                None, None, None, None, None, None,
                "Quantity remains pending until cash, exposure, pending orders, broker rules and risk limits are known",
                "One shared position, three research lenses; at most one active executable order per ticker",
                None, None, None, None, "Low", "Proposed", None, None, None, None, None,
                None, "REVIEW-%s-INITIAL" % BOOK_BUILD_ID_DATE, "WARAQAH-3.0",
                BOOK_BUILD_DATE, "No trade instruction; controlled Wait outcome", None, None, None,
            ]
            for col, value in enumerate(values, 1):
                is_input = col in ({12} | set(range(14, 21)) | {29, 30} |
                                   set(range(44, 55)))
                put(ws, r, col, value, font=F_IN if is_input else F_VAL,
                    fill=FILL_IN if is_input else None,
                    border=B_ALL, align=CENTER,
                    fmt=(FMT_TEXT if col in (1, 7, 8, 25, 47, 50, 51, 52) else
                         (FMT_PCT if col in (31, 32, 33) else
                          (FMT_MONEY if col in (18, 19, 34, 35, 36, 49) else None))))
            current_row += 1
            rank += 1
    for r in range(6, 506):
        put(ws, r, 11,
            f'=IF(G{r}="","",IF(J{r}="Tadawul",Checks!$B$12,'
            f'IF(OR(J{r}="NYSE",J{r}="NASDAQ"),Checks!$B$14,"")))',
            font=F_VAL, border=B_ALL)
        put(ws, r, 31,
            f'=IF(H{r}="","",IFERROR(VLOOKUP(H{r},Portfolio!$B:$K,10,FALSE),""))',
            font=F_VAL, fmt=FMT_PCT, border=B_ALL)
        put(ws, r, 34,
            f'=IF(OR(N{r}="Hold",N{r}="Wait"),0,IF(OR(H{r}="",Q{r}=""),"",'
            f'IF(OR(N{r}="Buy",N{r}="Add"),-Q{r}*IF(R{r}<>"",R{r},'
            f'IFERROR(VLOOKUP(H{r},\'Risk & Horizons\'!$A:$D,4,FALSE),"")),'
            f'IF(OR(N{r}="Trim",N{r}="Exit"),Q{r}*IF(R{r}<>"",R{r},'
            f'IFERROR(VLOOKUP(H{r},\'Risk & Horizons\'!$A:$D,4,FALSE),"")),0))))',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 35,
            f'=IF(AH{r}="","",AH{r}*IFERROR(VLOOKUP(H{r},'
            f'\'Risk & Horizons\'!$A:$AC,29,FALSE),""))',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 32,
            f'=IF(H{r}="","",IF(OR(N{r}="Hold",N{r}="Wait"),AE{r},'
            f'IF(AI{r}="","",IFERROR((VLOOKUP(H{r},Portfolio!$B:$G,6,FALSE)-AI{r})/'
            f'Portfolio!$G$26,""))))',
            font=F_VAL, fmt=FMT_PCT, border=B_ALL)
        put(ws, r, 33,
            f'=IF(H{r}="","",IFERROR(SUMIF(Portfolio!$AJ$6:$AJ$25,'
            f'VLOOKUP(H{r},\'Risk & Horizons\'!$A:$C,3,FALSE),Portfolio!$G$6:$G$25)/'
            f'Portfolio!$G$26+IF(OR(N{r}="Hold",N{r}="Wait"),0,'
            f'IF(AI{r}="","",-AI{r}/Portfolio!$G$26)),""))',
            font=F_VAL, fmt=FMT_PCT, border=B_ALL)
        base_cash = (
            f'IF(M{r}="SAR",Checks!$B$13,IF(M{r}="USD",Checks!$B$18*'
            f'IFERROR(VLOOKUP(H{r},\'Risk & Horizons\'!$A:$AC,29,FALSE),""),""))'
        )
        put(ws, r, 36,
            f'=IF({base_cash}="","",{base_cash}+'
            f'SUMIFS($AI$6:$AI$505,$AR$6:$AR$505,"Approved")+'
            f'SUMIFS($AI$6:$AI$505,$AR$6:$AR$505,"Submitted")+'
            f'SUMIFS($AI$6:$AI$505,$AR$6:$AR$505,"Part-filled"))',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 39,
            f'=IF(G{r}="","",IFERROR(VLOOKUP(G{r},Sharia!$A:$W,5,FALSE),"Uncertain"))',
            font=F_VAL, border=B_ALL)
        put(ws, r, 40,
            f'=IF(G{r}="","",IFERROR(VLOOKUP(G{r},Sharia!$A:$W,9,FALSE),""))',
            font=F_VAL, border=B_ALL)
        put(ws, r, 41,
            f'=IF(H{r}="","",IFERROR(VLOOKUP(H{r},\'Risk & Horizons\'!$A:$AD,30,FALSE),"Unknown"))',
            font=F_VAL, border=B_ALL)
        executable = (
            f'OR(N{r}="Buy",N{r}="Add",N{r}="Trim",N{r}="Exit")'
        )
        potential_buy = f'OR(N{r}="Buy",N{r}="Add",N{r}="Wait")'
        put(ws, r, 42,
            f'=IF(G{r}="","",TEXTJOIN("; ",TRUE,'
            f'IF(AND({potential_buy},IFERROR(VLOOKUP(G{r},Sharia!$A:$W,22,FALSE),"NO")<>"YES"),"Sharia gate not passed",""),'
            f'IF(AND({potential_buy},AN{r}=""),"Sharia evidence missing",""),'
            f'IF(AO{r}<>"Fresh","Data stale/unknown",""),'
            f'IF(IFERROR(VLOOKUP(H{r},\'Risk & Horizons\'!$A:$U,21,FALSE),0)<Checks!$B$7,"Data incomplete",""),'
            f'IF(OR(Checks!$B$9="",Checks!$B$10="",Checks!$B$11=""),"Risk limits pending",""),'
            f'IF(AND(OR({executable},N{r}="Wait"),K{r}=""),"Broker pending",""),'
            f'IF(AND({executable},OR(Q{r}="",O{r}="",O{r}="No order")),"Executable terms incomplete",""),'
            f'IF(AND({potential_buy},IF(M{r}="SAR",Checks!$B$13,Checks!$B$18)=""),"Cash balance pending",""),'
            f'IF(AND(AJ{r}<>"",AJ{r}<0),"Negative cash after commitments",""),'
            f'IF(AND(OR(N{r}="Buy",N{r}="Add"),AF{r}<>"",Checks!$B$10<>"",AF{r}>Checks!$B$10),"Position hard limit breached",""),'
            f'IF(AND(OR(N{r}="Buy",N{r}="Add"),AG{r}<>"",Checks!$B$11<>"",AG{r}>Checks!$B$11),"Sector limit breached",""),'
            f'IF(OR(AC{r}="",LEFT(AC{r},7)="Pending"),"Owner thesis pending",""),'
            f'IF(OR(AD{r}="",LEFT(AD{r},7)="Pending"),"Invalidation criteria pending",""),'
            f'IF(AL{r}="","Horizon reconciliation missing","")))',
            font=F_VAL, border=B_ALL)
        put(ws, r, 55,
            f'=IF(A{r}="","",IF(COUNTIFS($A$6:$A$505,A{r},$B$6:$B$505,B{r})>1,'
            f'"DUPLICATE","OK"))', font=F_VAL, border=B_ALL, align=CENTER)
        put(ws, r, 56,
            f'=IF(A{r}="","",IF(OR(AR{r}="Filled",AR{r}="Part-filled"),'
            f'IF(AND(AU{r}<>"",AV{r}>0,AW{r}>0,AX{r}<>""),"OK","MISSING FILL EVIDENCE"),'
            f'IF(AR{r}="Submitted",IF(AU{r}<>"","OK","MISSING BROKER ORDER ID"),'
            f'IF(AX{r}<>"","INVALID ACTIVITY LINK","OK"))))',
            font=F_VAL, border=B_ALL, align=CENTER)
        put(ws, r, 57,
            f'=IF(H{r}="","",IF(COUNTIFS($H$6:$H$505,H{r},$AR$6:$AR$505,"Approved")+'
            f'COUNTIFS($H$6:$H$505,H{r},$AR$6:$AR$505,"Submitted")+'
            f'COUNTIFS($H$6:$H$505,H{r},$AR$6:$AR$505,"Part-filled")>1,'
            f'"CONFLICT","OK"))', font=F_VAL, border=B_ALL, align=CENTER)
    _list_validation(ws, "F6:F505", horizons)
    _list_validation(ws, "N6:N505", ["Buy", "Add", "Hold", "Trim", "Exit", "Wait"])
    _list_validation(ws, "O6:O505", ["Limit", "Stop", "Stop limit", "Market", "No order"])
    _list_validation(ws, "AQ6:AQ505", ["Low", "Medium", "High"])
    _list_validation(ws, "AR6:AR505", [
        "Proposed", "Approved", "Submitted", "Part-filled", "Filled",
        "Cancelled", "Expired", "Superseded",
    ])
    _list_validation(ws, "AS6:AS505", ["Approve", "Reject", "Modify", "Defer"])
    widths = {get_column_letter(i): 15 for i in range(1, len(headers) + 1)}
    widths.update({
        "A": 28, "D": 14, "E": 14, "I": 28, "P": 34, "U": 48, "V": 52,
        "W": 44, "X": 44, "Y": 40, "Z": 36, "AA": 42, "AB": 42,
        "AC": 32, "AD": 34, "AK": 46, "AL": 46, "AN": 40, "AP": 52,
        "AU": 24, "AX": 28, "AY": 25, "AZ": 18, "BB": 40, "BD": 28,
        "BE": 24,
    })
    set_widths(ws, widths)
    set_filter(ws, 5, 505, len(headers))
    ws.freeze_panes = "A6"
    return ws


def sheet_performance(wb):
    ws = wb.create_sheet("Performance")
    ws.sheet_view.showGridLines = False
    put(ws, 2, 1, "الأداء — العائد المحقق وغير المحقق والتدفقات", font=F_TITLE)
    put(ws, 3, 1,
        "الأداء الفعلي منفصل عن المقترحات والمحاكاة. لا تُحسب الإيداعات أرباحاً، ولا "
        "يُعرض TWR أو XIRR أو التراجع قبل اكتمال السجل المؤرخ.", font=F_NOTE)
    header_row(ws, 5, 1, ["Metric", "Value", "Notes"])
    metrics_rows = [
        ("Portfolio value SAR", "=Portfolio!G26", "Current open-position market value", FMT_MONEY),
        ("Cost basis SAR", "=Portfolio!H26", "Open-position cost basis including recorded purchase costs", FMT_MONEY),
        ("Unrealized P/L SAR", "=Portfolio!I26", "Current value minus open-position cost basis", FMT_MONEY),
        ("Realized P/L SAR", "=SUM(Activity!V6:V505)", "Weighted-average method; trade costs already included once", FMT_MONEY),
        ("Net dividends SAR", '=SUMIFS(Activity!X6:X505,Activity!I6:I505,"Dividend")',
         "Dividend cash net of fees and withholding entered on the dividend row", FMT_MONEY),
        ("Standalone fees / tax cash SAR",
         '=SUMIFS(Activity!X6:X505,Activity!I6:I505,"Fee")+SUMIFS(Activity!X6:X505,Activity!I6:I505,"Tax")+SUMIFS(Activity!X6:X505,Activity!I6:I505,"Withholding")',
         "Negative standalone cash entries only; avoids double-counting trade costs", FMT_MONEY),
        ("Total actual P/L SAR", "=B8+B9+B10+B11",
         "Unrealized + realized + net dividends + standalone costs", FMT_MONEY),
        ("Fees paid SAR",
         '=SUMPRODUCT(Activity!N6:N505,Activity!M6:M505)+SUMPRODUCT((Activity!I6:I505="Fee")*Activity!P6:P505*Activity!M6:M505)',
         "Positive information line; already reflected in actual P/L where applicable", FMT_MONEY),
        ("Taxes paid SAR",
         '=SUMPRODUCT((Activity!I6:I505<>"Dividend")*Activity!O6:O505*Activity!M6:M505)+SUMPRODUCT((Activity!I6:I505="Tax")*Activity!P6:P505*Activity!M6:M505)',
         "Excludes dividend withholding shown separately", FMT_MONEY),
        ("Dividend withholding SAR",
         '=SUMPRODUCT((Activity!I6:I505="Dividend")*Activity!O6:O505*Activity!M6:M505)+SUMPRODUCT((Activity!I6:I505="Withholding")*Activity!P6:P505*Activity!M6:M505)',
         "Positive information line governed by the selected Sharia/tax method", FMT_MONEY),
        ("FX effect SAR", "Pending complete non-SAR transaction FX history",
         "Price return and currency translation remain separate; no fabricated FX attribution", None),
        ("Net external flows SAR",
         '=SUMIFS(Activity!X6:X505,Activity!I6:I505,"Deposit")+SUMIFS(Activity!X6:X505,Activity!I6:I505,"Withdrawal")',
         "Deposits positive and withdrawals negative; excluded from profit", FMT_MONEY),
        ("Simple return", '=IFERROR(B12/B7,"")',
         "Diagnostic only; use TWR/XIRR when dated history is sufficient", FMT_PCT),
        ("Money-weighted return (XIRR)", "Pending complete dated cash-flow history",
         "Requires broker-confirmed external flows and valuations", None),
        ("Actual drawdown", "Pending sufficient dated valuation history",
         "No historical path is manufactured from the opening snapshot", None),
        ("Concentration", "=Portfolio!E30", "Largest current position weight", FMT_PCT),
        ("Turnover since opening",
         '=IFERROR(SUMPRODUCT(((Activity!I6:I505="Buy")+(Activity!I6:I505="Sell"))*Activity!J6:J505*Activity!K6:K505*Activity!M6:M505)/AVERAGE(B6,B7),0)',
         "Gross recorded buys and sells divided by average of value and cost basis", FMT_PCT),
        ("Cash balance SAR (ledger)", '=SUMIFS(Activity!W6:W505,Activity!L6:L505,"SAR")',
         "Opening cash remains pending until broker evidence is supplied", FMT_MONEY),
        ("Cash balance USD (ledger)", '=SUMIFS(Activity!W6:W505,Activity!L6:L505,"USD")',
         "Retained in original currency; opening cash remains pending", FMT_MONEY),
        ("Saudi benchmark", "TASI — configurable; history pending",
         "Use a Sharia-compatible comparison selected by the owner where required", None),
        ("US benchmark", "S&P 500 total return — configurable; history pending",
         "Use an owner-approved Sharia-compatible alternative where required", None),
        ("Simulated proposal outcomes", "Excluded from actual P/L",
         "Research scenarios never change actual performance or Activity", None),
    ]
    for idx, (label, value, note, fmt) in enumerate(metrics_rows, 6):
        put(ws, idx, 1, label, font=F_LBL, border=B_ALL)
        put(ws, idx, 2, value, font=F_VAL, border=B_ALL, fmt=fmt)
        put(ws, idx, 3, note, font=F_NOTE, border=B_ALL)
    header_row(ws, 30, 1, ["Date", "Net External Flow SAR", "Ending Value SAR", "Period Return",
                           "TWR Index", "Benchmark", "Benchmark Return", "Review ID", "Notes"])
    put(ws, 31, 1, BOOK_BUILD_DATE, font=F_IN, fill=FILL_IN, border=B_ALL)
    put(ws, 31, 2, None, font=F_IN, fill=FILL_IN, fmt=FMT_MONEY, border=B_ALL)
    put(ws, 31, 3, "=B6", font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
    put(ws, 31, 4, None, font=F_VAL, fmt=FMT_PCT, border=B_ALL)
    put(ws, 31, 5, 1, font=F_VAL, fmt=FMT_NUM2, border=B_ALL)
    put(ws, 31, 6, "Pending", font=F_IN, fill=FILL_IN, border=B_ALL)
    put(ws, 31, 8, "REVIEW-%s-INITIAL" % BOOK_BUILD_ID_DATE,
        font=F_IN, fill=FILL_IN, border=B_ALL)
    put(ws, 31, 9, "Opening valuation snapshot only; no historical return asserted",
        font=F_IN, fill=FILL_IN, border=B_ALL)
    for r in range(32, 506):
        put(ws, r, 4, f'=IF(OR(A{r}="",C{r}="",C{r-1}=""),"",(C{r}-B{r})/C{r-1}-1)',
            font=F_VAL, fmt=FMT_PCT, border=B_ALL)
        put(ws, r, 5, f'=IF(D{r}="","",E{r-1}*(1+D{r}))', font=F_VAL,
            fmt=FMT_NUM2, border=B_ALL)
    set_widths(ws, {"A": 25, "B": 18, "C": 45, "D": 15, "E": 14,
                    "F": 20, "G": 18, "H": 25, "I": 45})
    set_filter(ws, 30, 505, 9)
    ws.freeze_panes = "A6"
    return ws


def sheet_checks(wb):
    ws = wb.create_sheet("Checks")
    ws.sheet_view.showGridLines = False
    put(ws, 2, 1, "الإعدادات وفحوص السلامة", font=F_TITLE)
    header_row(ws, 4, 1, ["Setting", "Value", "Purpose"])
    settings = [
        ("Reporting currency", "SAR", "All portfolio totals"),
        ("Timezone", "Asia/Riyadh", "Review and refresh timestamps"),
        ("Minimum data completeness", 0.70, "Below this, scores are not actionable"),
        ("Sharia authority", None, "Pending owner selection"),
        ("Max position soft", None, "Pending owner limit"),
        ("Max position hard", None, "Pending owner limit; blocks buys"),
        ("Max sector", None, "Pending owner limit; blocks buys"),
        ("Saudi broker", None, "Pending owner input"),
        ("Cash balance SAR", None, "Pending broker statement / owner input"),
        ("US broker", None, "Pending owner input"),
        ("Review cadence", None, "Prepared but not scheduled until approved"),
        ("Price stale after days", 5, "Calendar-day warning; market calendars recorded separately"),
        ("Sharia methodology / version", None, "Pending owner selection; never silently inferred"),
        ("Cash balance USD", None, "Pending broker statement / owner input"),
        ("Saudi benchmark", None, "Pending owner-approved Sharia-compatible comparison"),
        ("US benchmark", None, "Pending owner-approved Sharia-compatible comparison"),
    ]
    for idx, (key, value, note) in enumerate(settings, 5):
        put(ws, idx, 1, key, font=F_LBL, border=B_ALL)
        put(ws, idx, 2, value, font=F_IN, fill=FILL_IN, border=B_ALL,
            fmt=FMT_PCT if "Max " in key or "completeness" in key else None)
        put(ws, idx, 3, note, font=F_NOTE, border=B_ALL)
    header_row(ws, 23, 1, ["Check", "Result", "Severity", "Meaning"])
    checks = [
        ("Duplicate transaction IDs", '=IF(COUNTIF(Activity!AA6:AA505,"DUPLICATE")=0,"PASS","FAIL")', "ERROR", "Repeated imports must be idempotent"),
        ("Missing FX", '=IF(COUNTIFS(Activity!L6:L505,"<>SAR",Activity!A6:A505,"<>",Activity!M6:M505,"<=0")=0,"PASS","FAIL")', "ERROR", "No non-SAR accounting without an FX rate"),
        ("Negative positions", '=IF(COUNTIF(Portfolio!E6:E25,"<0")=0,"PASS","FAIL")', "ERROR", "A sale cannot exceed holdings"),
        ("Sharia buy/add gate", '=IF(COUNTIFS(Orders!N6:N505,"Buy",Orders!AM6:AM505,"<>Compliant")+COUNTIFS(Orders!N6:N505,"Add",Orders!AM6:AM505,"<>Compliant")=0,"PASS","FAIL")', "ERROR", "No Buy/Add unless the current screen is compliant"),
        ("Sharia evidence", '=IF(COUNTIFS(Orders!N6:N505,"Buy",Orders!AN6:AN505,"")+COUNTIFS(Orders!N6:N505,"Add",Orders!AN6:AN505,"")=0,"PASS","FAIL")', "ERROR", "Every Buy/Add needs linked evidence"),
        ("Risk limits configured", '=IF(AND(B9<>"",B10<>"",B11<>""),"PASS","WARN")', "WARNING", "Orders stay blocked while limits are pending"),
        ("Cash configured SAR/USD", '=IF(AND(B13<>"",B18<>""),"PASS","WARN")', "WARNING", "Sizing requires market-currency cash"),
        ("Review cadence configured", '=IF(B15<>"","PASS","WARN")', "WARNING", "No automation is activated yet"),
        ("Opening history complete", '=IF(COUNTBLANK(Activity!B6:B10)=0,"PASS","WARN")', "WARNING", "Acquisition dates await broker history"),
        ("Five current holdings", '=IF(COUNTIF(Portfolio!E6:E25,">0")=5,"PASS","FAIL")', "ERROR", "Baseline preservation check"),
        ("Owned price freshness", '=IF(COUNTIFS(Portfolio!B6:B25,"<>",Portfolio!AD6:AD25,"<>Fresh")=0,"PASS","FAIL")', "ERROR", "Owned securities require a current dated source"),
        ("Filled proposals reconciled", '=IF(COUNTIFS(Orders!AR6:AR505,"Filled",Orders!BD6:BD505,"<>OK")+COUNTIFS(Orders!AR6:AR505,"Part-filled",Orders!BD6:BD505,"<>OK")=0,"PASS","FAIL")', "ERROR", "Fills need broker evidence and an Activity transaction ID"),
        ("Submitted broker references", '=IF(COUNTIFS(Orders!AR6:AR505,"Submitted",Orders!BD6:BD505,"<>OK")=0,"PASS","FAIL")', "ERROR", "Submitted orders need broker order IDs"),
        ("Approved/submitted unblocked", '=IF(COUNTIFS(Orders!AR6:AR505,"Approved",Orders!AP6:AP505,"<>")+COUNTIFS(Orders!AR6:AR505,"Submitted",Orders!AP6:AP505,"<>")+COUNTIFS(Orders!AR6:AR505,"Part-filled",Orders!AP6:AP505,"<>")=0,"PASS","FAIL")', "ERROR", "No active order while a hard blocker remains"),
        ("Lifecycle statuses valid", '=IF(COUNTA(Orders!A6:A505)=COUNTIF(Orders!AR6:AR505,"Proposed")+COUNTIF(Orders!AR6:AR505,"Approved")+COUNTIF(Orders!AR6:AR505,"Submitted")+COUNTIF(Orders!AR6:AR505,"Part-filled")+COUNTIF(Orders!AR6:AR505,"Filled")+COUNTIF(Orders!AR6:AR505,"Cancelled")+COUNTIF(Orders!AR6:AR505,"Expired")+COUNTIF(Orders!AR6:AR505,"Superseded"),"PASS","FAIL")', "ERROR", "Every proposal uses the exact controlled lifecycle"),
        ("Duplicate proposal versions", '=IF(COUNTIF(Orders!BC6:BC505,"DUPLICATE")=0,"PASS","FAIL")', "ERROR", "Proposal ID plus version must be unique"),
        ("Active horizon conflicts", '=IF(COUNTIF(Orders!BE6:BE505,"CONFLICT")=0,"PASS","FAIL")', "ERROR", "At most one active executable order per ticker"),
        ("Executable terms complete", '=IF(COUNTIFS(Orders!N6:N505,"Buy",Orders!Q6:Q505,"")+COUNTIFS(Orders!N6:N505,"Add",Orders!Q6:Q505,"")+COUNTIFS(Orders!N6:N505,"Trim",Orders!Q6:Q505,"")+COUNTIFS(Orders!N6:N505,"Exit",Orders!Q6:Q505,"")+COUNTIFS(Orders!N6:N505,"Buy",Orders!O6:O505,"No order")+COUNTIFS(Orders!N6:N505,"Add",Orders!O6:O505,"No order")+COUNTIFS(Orders!N6:N505,"Trim",Orders!O6:O505,"No order")+COUNTIFS(Orders!N6:N505,"Exit",Orders!O6:O505,"No order")=0,"PASS","FAIL")', "ERROR", "Executable actions require quantity and order type"),
    ]
    for idx, (name, formula, severity, meaning) in enumerate(checks, 24):
        put(ws, idx, 1, name, font=F_LBL, border=B_ALL)
        put(ws, idx, 2, formula, font=F_VAL, border=B_ALL, align=CENTER)
        put(ws, idx, 3, severity, font=F_VAL, border=B_ALL, align=CENTER)
        put(ws, idx, 4, meaning, font=F_NOTE, border=B_ALL)
    header_row(ws, 44, 1, ["Review ID", "Run At", "Scope", "Outcome", "Evidence", "Notes"])
    review = ["REVIEW-%s-INITIAL" % BOOK_BUILD_ID_DATE, BOOK_BUILD_DATE,
              "Five existing Saudi holdings",
              "WAIT — Sharia authority/method/version, evidence, cash, brokers and owner limits pending",
              "Risk & Horizons sources; Sharia intentionally Uncertain; Orders Proposed/Wait",
              "Manual end-to-end control test; no executable quantity and no broker execution"]
    for col, value in enumerate(review, 1):
        put(ws, 45, col, value, font=F_VAL, border=B_ALL)
    set_widths(ws, {"A": 30, "B": 24, "C": 46, "D": 48, "E": 48, "F": 48})
    ws.freeze_panes = "A5"
    return ws


# --------------------------------------------------------------------------
# Guide
# --------------------------------------------------------------------------

GUIDE_CONTENT = [
    ("h", "دليل الاستخدام"),
    ("p", "هذا كتاب استثماري شخصي موحّد للسوقين السعودي والأمريكي. العملة الأصلية تبقى "
          "محفوظة، بينما تجمع تقارير المحفظة بالريال السعودي وفي توقيت Asia/Riyadh."),
    ("h", "ترتيب العمل"),
    ("p", "Activity هو السجل المحاسبي الوحيد: الشراء والبيع والتوزيعات والرسوم والضرائب "
          "والاستقطاع والإيداعات والسحوبات والتحويلات والانقسامات. Portfolio مشتق منه ولا يُعدل يدوياً."),
    ("p", "Sharia هو سجل الدليل: الجهة والمنهجية والإصدار وفترة التقرير وفحص النشاط والنسب "
          "والتواريخ والتطهير المستحق والمدفوع. الحالات: Compliant وNon-compliant وUncertain "
          "وReview overdue. لا شراء مع دليل ناقص أو منتهي، ولا يختار Codex منهجية دينية عن المالك."),
    ("p", "Orders طابور مقترحات مصنف وليس قناة تنفيذ. الإجراءات: Buy وAdd وHold وTrim وExit "
          "وWait. دورة الحالة: Proposed ثم Approved وSubmitted وPart-filled وFilled أو Cancelled "
          "أو Expired أو Superseded. لمس السعر لا يثبت التنفيذ، وفقط إثبات الوسيط يحدّث Activity."),
    ("h", "الحسابات"),
    ("p", "تكلفة المراكز بطريقة المتوسط المرجّح. البيع يخفض التكلفة بمتوسط الوحدة قبل "
          "الصفقة ويظهر الربح المحقق منفصلاً. كل صف يحمل Transaction ID لمنع الاستيراد المكرر."),
    ("p", "العائد السنوي مبني على إغلاق Yahoo Finance المعدّل للانقسامات والتوزيعات؛ لذلك "
          "هو عائد إجمالي. السعر والتوزيعات وFX والرسوم والاستقطاع والتدفقات الخارجية تبقى "
          "مفصولة، وPerformance لا يخلط الإيداعات بالأرباح أو المقترحات بالأداء الفعلي. "
          "مرجع الضبط: https://ranaroussi.github.io/yfinance/reference/yfinance.price_history.html"),
    ("p", "زخم 12-1 يقارن السعر قبل 12 شهراً بالسعر قبل شهر ويستبعد آخر 21 جلسة. مستوى "
          "التقلب (منخفض/عادي/مرتفع) منفصل عن اتجاهه (صاعد/مستقر/هابط)."),
    ("p", "النتيجة المركبة تعيد وزن العناصر المتاحة فقط، وتعرض اكتمال البيانات. دون 70% "
          "تظهر بيانات ناقصة ولا يجوز تحويل النتيجة إلى إجراء."),
    ("h", "الآفاق الثلاثة"),
    ("p", "6 أشهر: الاتجاه المتوسط والزخم القصير وRSI ومستوى/اتجاه التقلب. سنتان: العائد "
          "السنوي وزخم 12-1 وSMA200 والتقييم والتراجع. 5 سنوات: جودة الأرباح وROE "
          "والتوزيعات واستدامتها والتقييم والمخاطر الهيكلية."),
    ("p", "الآفاق الثلاثة عدسات بحث لمركز واحد وليست ثلاث حصص متعارضة. الإشارة ليست ترجمة "
          "آلية للنتيجة؛ تُحجب عند نقص البيانات أو الدليل الشرعي أو النقد أو حدود المخاطر، "
          "ولا يجوز وجود أكثر من أمر نشط واحد للأسهم نفسها."),
    ("h", "المخاطر والإعدادات"),
    ("p", "حدود المركز المرنة والصارمة وحد القطاع قابلة للضبط في Checks. بقيت معلّقة حتى "
          "يعتمدها المالك؛ أي اقتراح شراء يبقى Wait في هذه الأثناء."),
    ("p", "السعودية تستخدم SAR وتوقيت Asia/Riyadh وتسوية الأسهم T+2 بحسب Saudi Exchange: "
          "https://www.saudiexchange.sa/wps/portal/saudiexchange/trading/market-services/equities?locale=en . "
          "أمريكا تستخدم USD وتوقيت America/New_York مع DST وتسوية T+1 بحسب SEC: "
          "https://www.sec.gov/rules-regulations/2023/02/34-96930 . قواعد الكمية والأمر يؤكدها الوسيط."),
    ("h", "التحديث والمراجعة"),
    ("p", "التحديث يكتب فترة القوائم وتاريخ النشر ووقت الملاحظة ووقت الجلب وأساس السعر والعائد "
          "ووحدة التوزيعات والمصدر والرابط. عند الفشل يبقى آخر سجل صالح معلّماً stale ولا يُستبدل بتخمين."),
    ("p", "التحديثات تحفظ Activity وSharia وقرارات Orders وسجل Performance وإعدادات Checks "
          "وسجل المراجعات. الاستيراد المتكرر بالمعرف نفسه يحدّث الصف ولا يضاعف الحركة. "
          "XIRR لا يُستخدم قبل تدفقات مؤرخة كافية: https://support.google.com/docs/answer/3093266 "
          "ومرجع الأداء: https://www.cfainstitute.org/-/media/documents/code/gips/2020-gips-standards-asset-owners.pdf"),
    ("p", "سير المراجعة: اقرأ السياسة والمراكز؛ تحقق من البيانات والدليل الشرعي؛ ابحث التغييرات؛ "
          "حدّث الآفاق؛ قيّم ملاءمة المحفظة؛ حدّث المقترحات؛ وسجّل التغييرات والمصادر. الجدولة "
          "مُهيأة فقط ولا تُفعّل حتى يختار المالك الوتيرة، والإشعار للتغيير المهم أو الفشل فقط."),
    ("h", "حدود صريحة"),
    ("p", "الأسعار ليست لحظية، وYahoo Finance مصدر مجاني قد ينقصه بعض الحقول. حالة "
          "Uncertain لا تعني مخالفة شرعية؛ تعني أن الدليل المعتمد غير مكتمل."),
    ("p", "هذه أداة حفظ وتحليل ومراجعة وليست توصية أو تفويضاً بالتداول. سعر الوقف لا يضمن "
          "حداً أقصى للخسارة، ولا يُشترى اشتراك بيانات أو تُنشأ خدمة مدفوعة من دون موافقة المالك."),
]


def sheet_guide(wb):
    ws = wb.create_sheet("Guide")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.rightToLeft = True
    r = GD_FIRST_ROW
    for kind, text in GUIDE_CONTENT:
        if kind == "h":
            merge_band(ws, r, 2, 6, text, font=F_GUIDE_H, align=RIGHT)
            ws.row_dimensions[r].height = 24
            r += 2
        else:
            put(ws, r, 2, text, font=F_GUIDE, align=WRAP)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.row_dimensions[r].height = 36
            r += 1
    set_widths(ws, {"A": 3, "B": 40, "C": 30, "D": 30, "E": 30, "F": 30})
    return ws


# --------------------------------------------------------------------------
# build
# --------------------------------------------------------------------------

def build(out_path, data_dir="data", preserve_from=None):
    """Build the full auditable Saudi/US investment book from CSV/JSON data.

    preserve_from: path to a previous v2 xlsx — read Portfolio sheet columns
    B (symbol), E (shares), F (cost) rows 6-25 from it and restore into the new
    build. If None, seed Portfolio with demo rows.

    Returns the path actually written; it differs from `out_path` only when the
    target was locked and the build fell back to <name>_new.xlsx.
    """
    data = load_data(data_dir)
    rows = build_rows(data)
    holdings = read_preserved(preserve_from) or list(DEMO_HOLDINGS)

    wb = Workbook()
    wb.remove(wb.active)
    sheet_portfolio(wb, holdings)
    sheet_orders(wb, holdings, rows)
    sheet_lookup(wb, rows)
    sheet_performance(wb)
    sheet_activity(wb, holdings)
    sheet_sharia(wb, holdings, rows)
    sheet_risk(wb, rows)
    sheet_db(wb, data, rows)
    sheet_statements(wb, data, rows)
    sheet_symbols(wb, data["all_symbols"])
    sheet_checks(wb)
    sheet_guide(wb)
    restore_manual_records(wb, preserve_from)
    assert wb.sheetnames == SHEET_ORDER, wb.sheetnames
    wb.active = 0

    try:
        wb.save(out_path)
        written = out_path
    except PermissionError:
        base, ext = os.path.splitext(out_path)
        written = base + "_new" + ext
        print("WARNING: %s is locked (open in Excel?) — saving to %s instead"
              % (out_path, written))
        wb.save(written)

    print("built %s: %d symbols, %d DB rows"
          % (written, len(rows), sum(len(v) for v in data["annual"].values())))
    return written


def main():
    # Arabic console output must survive a non-UTF-8 (cp1252) console.
    import sys
    for _stream in (sys.stdout, sys.stderr):
        if _stream is not None and hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    default_out = os.path.join(HERE, "..", "Sahm_Portfolio_Analysis_v2.xlsx")
    parser = argparse.ArgumentParser(
        description="Build the Saudi stock analysis workbook.")
    parser.add_argument("--out", default=default_out, help="output xlsx path")
    parser.add_argument("--data-dir", default=os.path.join(HERE, "data"),
                        help="dir holding annual_metrics.csv, statements.csv, snapshot.json")
    parser.add_argument("--preserve", default=None,
                        help="previous v2 xlsx to restore Portfolio holdings from")
    args = parser.parse_args()
    build(args.out, data_dir=args.data_dir, preserve_from=args.preserve)


if __name__ == "__main__":
    main()
