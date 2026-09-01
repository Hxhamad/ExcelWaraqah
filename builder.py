"""Workbook builder for the Saudi (Tadawul) stock analysis tool, v2.

Builds a 7-sheet xlsx from the CSV/JSON artefacts produced by fetcher.py:

    Portfolio | Stock Lookup | Risk & Horizons | DB | Statements | Symbols | Guide

Design notes that matter for anyone editing this file
-----------------------------------------------------
* openpyxl 3.1.5 quirks are load-bearing here: PatternFill is always built with
  keyword args (positional fill types blow up at save time), and every merged
  range gets its anchor cell written BEFORE the merge call.
* Every row number is a module-level constant; formulas are assembled with
  f-strings so a layout change never leaves a stale hardcoded row behind.
* Formulas stay in the Excel-2007 function set (VLOOKUP / INDEX / MATCH / IF /
  OR / AND / COUNTIF / SUM / MAX / IFERROR / ISNA / LEN / ROUND / TEXT) so that
  LibreOffice's headless recalc evaluates all of them.
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

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import metrics  # noqa: E402
from symbols import load_symbols  # noqa: E402

# --------------------------------------------------------------------------
# layout constants
# --------------------------------------------------------------------------
SHEET_ORDER = [
    "Portfolio", "Stock Lookup", "Risk & Horizons",
    "DB", "Statements", "Symbols", "Guide",
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


def near_verdict(sma_flag, rsi, vol_reg, ret_1m):
    """1-3 months: trend filter + RSI extremes + volatility regime + 1M drift."""
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
    if vol_reg:
        present += 1
        if vol_reg == "HIGH":
            pts -= 1
    if ret_1m is not None:
        present += 1
        pts += 1 if ret_1m > 0 else -1
    return _verdict(pts, present)


def mid_verdict(ret_6m, ret_1y, sma_flag, pe, maxdd_2y):
    """6-18 months: medium momentum + trend + valuation + drawdown tolerance."""
    pts = 0
    present = 0
    if ret_6m is not None:
        present += 1
        pts += 1 if ret_6m > 0 else -1
    if ret_1y is not None:
        present += 1
        pts += 1 if ret_1y > 0 else -1
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
    """3+ years: return on equity, income, payout sustainability, valuation."""
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
        vol_reg = snap.get("vol_regime")
        rsi = snap.get("rsi14")
        momentum = last.get("momentum")

        score = metrics.composite_score(pe, roe, div_yield, (sma_flag, momentum), maxdd_2y)
        # Flag only the inputs that actually feed the score and the verdicts -
        # this is what "الحكم محايد للأجزاء الناقصة" refers to. Gaps in the old
        # annual history are reported separately (Statements col "ملاحظة").
        missing_inputs = any(v is None for v in
                             (pe, roe, div_yield, sma_flag, momentum, maxdd_2y))
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
            "rsi": rsi,
            "vol_regime": "مرتفع" if vol_reg == "HIGH" else ("عادي" if vol_reg else NA),
            "sma200": "فوق" if sma_flag == "above" else ("تحت" if sma_flag == "below" else NA),
            "maxdd_2y": maxdd_2y,
            "oil_beta": oil_beta(entry["sector"]),
            "near": near_verdict(sma_flag, rsi, vol_reg, snap.get("ret_1m")),
            "mid": mid_verdict(snap.get("ret_6m"), snap.get("ret_1y"), sma_flag, pe, maxdd_2y),
            "far": far_verdict(roe, div_yield, payout, pe),
            "score": score,
            "rating": metrics.rating(score),
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


def sheet_portfolio(wb, holdings):
    ws = wb.create_sheet("Portfolio")
    ws.sheet_view.showGridLines = False

    put(ws, PF_TITLE_ROW, 2, "محفظة الأسهم السعودية — التحليل وحجم المراكز", font=F_TITLE)
    put(ws, PF_NOTE_ROW, 2,
        "عدّل الخلايا الصفراء فقط (عدد الأسهم ومتوسط التكلفة). باقي الأعمدة معادلات "
        "تقرأ من ورقة Risk & Horizons. راجع ورقة Guide.",
        font=F_NOTE)

    headers = [
        "الرمز", "الشركة", "السعر", "عدد الأسهم", "متوسط التكلفة",
        "القيمة السوقية", "التكلفة الإجمالية", "الربح/الخسارة", "نسبة الربح/الخسارة",
        "الوزن", "النتيجة", "التقييم", "الإشارة",
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
        put(ws, r, 5, held[1] if held else None,
            font=F_IN, fill=FILL_IN, fmt=FMT_INT, border=B_ALL)
        put(ws, r, 6, held[2] if held else None,
            font=F_IN, fill=FILL_IN, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 7, f'=IF(OR(E{r}="",D{r}=""),"",D{r}*E{r})',
            font=F_VAL, fmt=FMT_MONEY, border=B_ALL)
        put(ws, r, 8, f'=IF(OR(E{r}="",F{r}=""),"",E{r}*F{r})',
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
        put(ws, r, 14,
            f'=IF(M{r}="","",IF(M{r}="شراء قوي","دخول قوي",IF(M{r}="شراء","دخول",'
            f'IF(M{r}="بيع قوي","خروج","احتفاظ/مراجعة"))))',
            font=F_VAL, border=B_ALL)

    # ---- TOTAL row
    put(ws, PF_TOTAL_ROW, 2, "الإجمالي", font=F_TOT, fill=FILL_TOT, border=B_ALL)
    for col in (3, 4, 5, 6, 12, 13, 14):
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
         f'=IF(E{PF_SUM_MAXW_ROW}>=0.4,"عالي","مقبول")', None),
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
        "H": 16, "I": 15, "J": 13, "K": 10, "L": 10, "M": 14, "N": 15,
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
        ("حكم قريب (1-3 أشهر)", rh(15)),
        ("حكم متوسط (6-18 شهر)", rh(16)),
        ("حكم بعيد (3 سنوات فأكثر)", rh(17)),
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
        "بيتا النفط", "حكم قريب", "حكم متوسط", "حكم بعيد", "النتيجة", "التقييم",
    ]
    header_row(ws, RH_HDR_ROW, 1, headers)

    fmts = {
        4: FMT_MONEY, 5: FMT_PCT, 6: FMT_PCT, 7: FMT_PCT, 8: FMT_PCT, 9: FMT_PCT,
        10: FMT_NUM1, 13: FMT_PCT, 18: FMT_NUM1,
    }
    for i, rec in enumerate(rows):
        r = RH_FIRST_ROW + i
        values = [
            rec["code"], rec["name"], rec["sector"] or NA, rec["price"],
            rec["ret_1w"], rec["ret_1m"], rec["ret_3m"], rec["ret_6m"], rec["ret_1y"],
            rec["rsi"], rec["vol_regime"], rec["sma200"], rec["maxdd_2y"],
            rec["oil_beta"], rec["near"], rec["mid"], rec["far"],
            rec["score"], rec["rating"],
        ]
        for col, value in enumerate(values, start=1):
            put(ws, r, col, value, font=F_STAT, border=B_ALL,
                fmt=(FMT_TEXT if col == 1 else fmts.get(col)),
                align=(LEFT if col == 2 else CENTER))

    set_widths(ws, {
        "A": 9, "B": 28, "C": 22, "D": 10, "E": 11, "F": 11, "G": 11, "H": 11,
        "I": 11, "J": 9, "K": 12, "L": 13, "M": 15, "N": 11, "O": 18, "P": 18,
        "Q": 18, "R": 9, "S": 14,
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
        "داخل كل رمز (الأحدث أولاً).", font=F_TITLE)

    headers = [
        "الرمز", "السنة", "الإغلاق", "العائد السنوي", "التقلب السنوي",
        "أقصى تراجع", "التوزيعات (ريال)", "عائد التوزيعات", "زخم 12-1",
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
# Guide
# --------------------------------------------------------------------------

GUIDE_CONTENT = [
    ("h", "دليل الاستخدام"),
    ("p", "هذا الملف أداة تحليل للأسهم السعودية (تداول). الخلايا الصفراء فقط قابلة "
          "للتعديل، وباقي الأرقام إما معادلات أو بيانات ثابتة من آخر تحديث."),
    ("h", "المفاهيم"),
    ("p", "النتيجة المركبة: رقم من 0 إلى 100 يجمع خمسة محاور بأوزان ثابتة، والأعلى أفضل."),
    ("p", "التقييم: ترجمة النتيجة إلى شريحة — شراء قوي (80 فأكثر)، شراء (65-79)، "
          "تعزيز/احتفاظ (50-64)، بيع (35-49)، بيع قوي (أقل من 35)."),
    ("p", "أقصى تراجع: أسوأ هبوط من قمة إلى قاع خلال الفترة، ويقاس بالسالب."),
    ("p", "نظام التقلب: مقارنة تقلب 20 يوماً بتقلب 60 يوماً. «مرتفع» يعني أن النافذة "
          "الأطول أكثر اضطراباً، أي أن السهم يمر بفترة ضغط."),
    ("p", "SMA200: المتوسط المتحرك لمئتي يوم. «فوق» يعني أن السعر أعلى من المتوسط، "
          "و«تحت» يعني أنه أدنى منه."),
    ("p", "بيتا النفط: تصنيف قطاعي تقريبي لاتجاه تأثر السهم بأسعار النفط "
          "(موجبة / سالبة / محايدة)، وليس معاملاً محسوباً."),
    ("h", "الأوزان"),
    ("p", "القيمة (مكرر الربحية) 30% — أوضح عامل مثبت في السوق السعودي: الشرائح الأرخص "
          "سبقت الشرائح الأغلى في العائد التاريخي بفارق أكبر من أي عامل آخر."),
    ("p", "الجودة (العائد على حقوق الملكية) 20% — ROE المرتفع والمستقر يميّز الشركات "
          "التي تعيد استثمار أرباحها بكفاءة، ويقلّل احتمال انهيار الأرباح فجأة."),
    ("p", "الفني (SMA200 مع الزخم) 20% — فلتر الاتجاه يمنع الدخول في الأسهم الهابطة، "
          "والزخم 12-1 من أكثر الإشارات ثباتاً عبر الأسواق الناشئة ومنها تداول."),
    ("p", "التوزيعات 15% — السوق السعودي سوق توزيعات: جزء كبير من العائد التاريخي جاء "
          "من الأرباح النقدية لا من ارتفاع السعر وحده."),
    ("p", "المخاطر (أقصى تراجع سنتين) 15% — الأسهم الأقل تراجعاً أسهل في الالتزام بها، "
          "ولا تجبر المستثمر على البيع في القاع."),
    ("p", "أي عنصر ناقص يأخذ 50 نقطة (محايد) بدلاً من صفر، حتى لا تُعاقب الشركة على "
          "نقص في مصدر البيانات."),
    ("h", "قواعد المخاطر"),
    ("p", "التنويع: من 10 إلى 15 سهماً. أقل من ذلك تركيز عالٍ، وأكثر منه صعب المتابعة."),
    ("p", "حجم المركز: سهم واحد لا يتجاوز 20% عملياً، و40% حد صارم لا يُخترق. خانة "
          "«حالة التركيز» في ورقة Portfolio تتحول إلى «عالي» عند 40% فأكثر."),
    ("p", "فلتر SMA200: لا تبنِ مركزاً في سهم تحت متوسط 200 يوم — هذا الفلتر وحده يقلّص "
          "عمق التراجع إلى النصف تقريباً مقارنة بالشراء دون فلتر."),
    ("p", "نظام التقلب: قلّل التعرض عندما يكون «مرتفع»، وزده تدريجياً عندما يعود «عادي»."),
    ("p", "موازنة بيتا النفط: لا تجعل المحفظة كلها في اتجاه واحد. الطاقة والبتروكيماويات "
          "والمرافق بيتا نفط موجبة، والبنوك والمالية سالبة، والصحة والاتصالات "
          "والاستهلاكي محايدة. اخلط الثلاثة."),
    ("h", "الآفاق الثلاثة"),
    ("p", "كل حكم يجمع نقاطاً موجبة وسالبة ثم: مجموع +2 فأكثر → «إيجابي»، ومجموع -2 فأقل "
          "→ «سلبي»، وما بينهما → «محايد». أي عنصر ناقص لا يضيف ولا يخصم، وإذا غابت كل "
          "العناصر يصبح الحكم «محايد (بيانات ناقصة)»."),
    ("p", "حكم قريب (1-3 أشهر): فوق SMA200 (+1) أو تحته (-1)؛ RSI بين 40 و70 (+1)، أو "
          "فوق 70 تشبع شرائي (-1)، أو تحت 30 تشبع بيعي وارتداد محتمل (+1)؛ نظام تقلب "
          "مرتفع (-1)؛ عائد الشهر موجب (+1) أو سالب (-1)."),
    ("p", "حكم متوسط (6-18 شهراً): عائد 6 أشهر موجب (+1) وإلا (-1)؛ عائد سنة موجب (+1) "
          "وإلا (-1)؛ فوق SMA200 (+1) وإلا (-1)؛ مكرر ربحية 18 فأقل (+1) أو أكثر من 25 "
          "(-1)؛ أقصى تراجع سنتين أفضل من -25% (+1) وإلا (-1)."),
    ("p", "حكم بعيد (3 سنوات فأكثر): ROE 20% فأكثر (+2) أو 15% فأكثر (+1) أو أقل من 5% "
          "(-1)؛ عائد توزيعات 3% فأكثر (+1)؛ نسبة توزيع 80% فأقل (+1) أو أكثر من 100% "
          "(-1)؛ مكرر ربحية 18 فأقل (+1) أو أكثر من 25 (-1)."),
    ("p", "كيف تُستخدم: الحكم البعيد يقرر «هل أملك هذا السهم أصلاً»، والمتوسط يقرر «هل "
          "أزيد أم أخفّف»، والقريب يقرر «متى أنفّذ». تعارض الأحكام إشارة انتظار، "
          "وليس إشارة بيع."),
    ("h", "كيفية التحديث"),
    ("p", "التحديث السريع (refresh.bat quick): يحدّث الأسعار والعوائد والمؤشرات الفنية "
          "لكل الرموز فقط — دقائق معدودة، ولا يلمس القوائم المالية."),
    ("p", "التحديث الكامل (refresh.bat full): يعيد سحب التاريخ السنوي والقوائم المالية "
          "كاملة ثم يعيد بناء الملف — أبطأ بكثير، ويُستخدم كل ربع سنة أو بعد إعلان النتائج."),
    ("p", "مراكزك محفوظة بعد أي تحديث: البناء يقرأ الرمز وعدد الأسهم ومتوسط التكلفة من "
          "الملف السابق ويعيد كتابتها في الملف الجديد."),
    ("h", "حدود صادقة"),
    ("p", "لا يوجد مؤشر تاسي عبر هذا المصدر، لذلك لا توجد بيتا حقيقية مقابل السوق ولا "
          "أداء نسبي مقابل المؤشر."),
    ("p", "القوائم المالية متاحة لخمس سنوات فقط (2021-2025)، وما قبلها غير متوفر."),
    ("p", "الأسعار مؤخرة وليست لحظية، ولا تصلح للتداول اليومي."),
    ("p", "«بيانات ناقصة» تعني أن المصدر لم يعطِ القيمة؛ يُحتسب العنصر محايداً ولا "
          "يُفسَّر كإشارة سلبية."),
    ("p", "هذه الأداة مساعدة على القرار وليست توصية استثمارية."),
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
    """Build the full 7-sheet workbook from CSV/JSON data.

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
    sheet_lookup(wb, rows)
    sheet_risk(wb, rows)
    sheet_db(wb, data, rows)
    sheet_statements(wb, data, rows)
    sheet_symbols(wb, data["all_symbols"])
    sheet_guide(wb)
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
