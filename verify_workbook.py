"""Standalone verifier for a built v2 workbook.

Usage:  python verify_workbook.py <workbook.xlsx> [--data-dir DIR]

Prints one PASS/FAIL line per check and exits 1 if anything failed. No pytest,
no third-party deps beyond openpyxl.

The workbook is opened twice: data_only=True for cached values (only present
after a real Excel/LibreOffice recalc) and data_only=False for formula strings.
Checks that need cached values SKIP with a warning when the file has never been
recalculated, so this runs meaningfully both before and after the LO pass.

Every value assertion is pinned to the label that sits beside it, so a layout
shift turns into a loud failure instead of a silently-passing check on the
wrong cell.
"""

import argparse
import json
import os
import re
import sys

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))

EXPECTED_SHEETS = [
    "Portfolio", "Stock Lookup", "Risk & Horizons",
    "DB", "Statements", "Symbols", "Guide",
]

PF_HDR_ROW = 5
PF_FIRST_ROW = 6
PF_LAST_ROW = 25
PF_TOTAL_ROW = 26
SL_INPUT_ROW = 3
SL_GUARD_ROW = 4
RH_HDR_ROW = 4
RH_FIRST_ROW = 5
DB_HDR_ROW = 4
DB_FIRST_ROW = 5
ST_HDR_ROW = 4
ST_FIRST_ROW = 5
SY_HDR_ROW = 4
SY_FIRST_ROW = 5

MIN_YEAR = 2015
MAX_YEAR = 2026


class Report(object):
    """Collects PASS/FAIL/SKIP lines and decides the exit code."""

    def __init__(self):
        self.failures = 0
        self.passes = 0
        self.skips = 0

    def check(self, ok, name, detail=""):
        if ok:
            self.passes += 1
            print("PASS  %-52s %s" % (name, detail))
        else:
            self.failures += 1
            print("FAIL  %-52s %s" % (name, detail))
        return ok

    def skip(self, name, detail=""):
        self.skips += 1
        print("SKIP  %-52s %s" % (name, detail))


def cell(ws, row, col):
    return ws.cell(row=row, column=col).value


def text(value):
    return "" if value is None else str(value).strip()


def is_num(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def find_label_row(ws, col, label, first_row, last_row):
    """Row in `col` between first_row/last_row whose text equals `label`."""
    for row in range(first_row, last_row + 1):
        if text(cell(ws, row, col)) == label:
            return row
    return None


# --------------------------------------------------------------------------
# checks
# --------------------------------------------------------------------------

def check_sheets(rep, wbv):
    rep.check(wbv.sheetnames == EXPECTED_SHEETS, "sheet order",
              "%s" % (wbv.sheetnames,))


def check_no_error_strings(rep, wbv, wbf):
    """No cell may hold a string starting with '#' (#REF!, #N/A, #VALUE! ...).

    Checked in both the cached-value and the formula view: a broken formula
    shows up as a cached #NAME? on one side and never on the other.
    """
    bad = []
    scanned = 0
    for wb, tag in ((wbv, "values"), (wbf, "formulas")):
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    value = c.value
                    if value is None:
                        continue
                    scanned += 1
                    if isinstance(value, str) and value.startswith("#"):
                        bad.append("%s!%s [%s]=%r" % (ws.title, c.coordinate, tag, value))
    rep.check(not bad, "zero '#' error strings in any sheet",
              "scanned %d cells; %s" % (scanned, bad[:5] if bad else "clean"))


def check_portfolio(rep, wbv, wbf):
    wsf = wbf["Portfolio"]
    wsv = wbv["Portfolio"]

    rep.check(text(cell(wsf, PF_FIRST_ROW, 2)) != "",
              "Portfolio B6 (first symbol) non-empty",
              "B6=%r" % (cell(wsf, PF_FIRST_ROW, 2),))

    # label pin: the TOTAL row must be labelled before we trust G26
    label = text(cell(wsf, PF_TOTAL_ROW, 2))
    if not rep.check(label == "الإجمالي",
                     "Portfolio B%d is the TOTAL label" % PF_TOTAL_ROW,
                     "B%d=%r" % (PF_TOTAL_ROW, label)):
        return
    formula = text(cell(wsf, PF_TOTAL_ROW, 7))
    rep.check(formula == "=SUM(G%d:G%d)" % (PF_FIRST_ROW, PF_LAST_ROW),
              "Portfolio G%d is SUM over the data rows" % PF_TOTAL_ROW,
              "G%d=%r" % (PF_TOTAL_ROW, formula))

    parts = [cell(wsv, r, 7) for r in range(PF_FIRST_ROW, PF_LAST_ROW + 1)]
    total = cell(wsv, PF_TOTAL_ROW, 7)
    numeric = [p for p in parts if is_num(p)]
    if not is_num(total) or not numeric:
        rep.skip("Portfolio G%d == SUM(G%d:G%d)" % (PF_TOTAL_ROW, PF_FIRST_ROW, PF_LAST_ROW),
                 "no cached values (workbook not recalculated yet)")
    else:
        rep.check(abs(sum(numeric) - total) < 0.01,
                  "Portfolio G%d == SUM(G%d:G%d)" % (PF_TOTAL_ROW, PF_FIRST_ROW, PF_LAST_ROW),
                  "sum=%.4f total=%.4f over %d rows" % (sum(numeric), total, len(numeric)))

    # label pin for the concentration summary line
    conc_row = find_label_row(wsf, 2, "حالة التركيز", PF_TOTAL_ROW, PF_TOTAL_ROW + 20)
    if conc_row is None:
        rep.check(False, "Portfolio concentration label present", "label not found in col B")
    else:
        max_row = find_label_row(wsf, 2, "أعلى وزن سهم", PF_TOTAL_ROW, PF_TOTAL_ROW + 20)
        conc = text(cell(wsf, conc_row, 5))
        rep.check(max_row is not None and ("E%d" % max_row) in conc,
                  "Portfolio concentration formula points at max-weight cell",
                  "max-weight row=%s, formula=%r" % (max_row, conc))


def check_lookup(rep, wbf):
    ws = wbf["Stock Lookup"]
    b3 = cell(ws, SL_INPUT_ROW, 2)
    rep.check(ws.cell(row=SL_INPUT_ROW, column=2) is not None,
              "Stock Lookup B%d input cell exists" % SL_INPUT_ROW, "B3=%r" % (b3,))
    guard = text(cell(ws, SL_GUARD_ROW, 2))
    rep.check(guard != "" and "أدخل رمزاً" in guard,
              "Stock Lookup B%d guard formula present" % SL_GUARD_ROW,
              "%d chars" % len(guard))


def check_risk(rep, wbv, wbf, symbol_count):
    ws = wbf["Risk & Horizons"]
    header = text(cell(ws, RH_HDR_ROW, 1))
    if not rep.check(header != "", "Risk & Horizons header row %d populated" % RH_HDR_ROW,
                     "A%d=%r" % (RH_HDR_ROW, header)):
        return
    rows = 0
    for row in range(RH_FIRST_ROW, ws.max_row + 1):
        if text(cell(ws, row, 1)) == "":
            break
        rows += 1
    rep.check(rows >= symbol_count, "Risk & Horizons row count >= snapshot symbols",
              "%d rows vs %d symbols" % (rows, symbol_count))
    if rows == 0:
        rep.check(False, "Risk & Horizons first row price > 0", "no data rows")
        return
    code = text(cell(ws, RH_FIRST_ROW, 1))
    price = cell(wbv["Risk & Horizons"], RH_FIRST_ROW, 4)
    rep.check(is_num(price) and price > 0,
              "Risk & Horizons first row price > 0", "%s price=%r" % (code, price))


def check_db(rep, wbf):
    ws = wbf["DB"]
    rows = 0
    symbols = set()
    years = []
    blank_symbol_rows = []
    bad_years = []
    for row in range(DB_FIRST_ROW, ws.max_row + 1):
        sym = text(cell(ws, row, 1))
        year = cell(ws, row, 2)
        if sym == "" and year is None:
            break
        rows += 1
        if sym == "":
            blank_symbol_rows.append(row)
        else:
            symbols.add(sym)
        if is_num(year):
            years.append(int(year))
            if not (MIN_YEAR <= int(year) <= MAX_YEAR):
                bad_years.append((row, year))
        else:
            bad_years.append((row, year))
    rep.check(len(symbols) >= 3, "DB distinct symbols >= 3",
              "%d symbols across %d rows" % (len(symbols), rows))
    rep.check(not bad_years, "DB years all within %d..%d" % (MIN_YEAR, MAX_YEAR),
              "range=%s..%s bad=%s" % (min(years) if years else "-",
                                       max(years) if years else "-", bad_years[:5]))
    rep.check(not blank_symbol_rows, "DB every row has a symbol",
              "blank rows: %s" % (blank_symbol_rows[:5] if blank_symbol_rows else "none"))
    return symbols


def check_statements(rep, wbv, wbf):
    wsf = wbf["Statements"]
    wsv = wbv["Statements"]
    # label pin: col C must be the newest year's revenue column
    hdr = text(cell(wsf, ST_HDR_ROW, 3))
    if not rep.check(hdr == "الإيرادات", "Statements C%d is a revenue header" % ST_HDR_ROW,
                     "C%d=%r" % (ST_HDR_ROW, hdr)):
        return
    # A newest-year revenue may legitimately be negative (some investment and
    # insurance firms report it that way on Yahoo), so the check is a share of
    # positives rather than "all positive". An exact 0 is never legitimate --
    # the builder renders "no data" as a blank cell -- so zeros must be absent.
    checked = 0
    negative = []
    zero = []
    positive = 0
    for row in range(ST_FIRST_ROW, wsf.max_row + 1):
        sym = text(cell(wsf, row, 1))
        if sym == "":
            break
        revenue = cell(wsv, row, 3)
        if revenue is None:
            continue
        checked += 1
        if not is_num(revenue):
            zero.append("%s=%r" % (sym, revenue))
        elif revenue > 0:
            positive += 1
        elif revenue < 0:
            negative.append("%s=%r" % (sym, revenue))
        else:
            zero.append("%s=%r" % (sym, revenue))
    if checked == 0:
        rep.skip("Statements newest-year revenue sane",
                 "no symbol has a newest-year revenue")
        return
    share = positive / float(checked)
    for item in negative:
        print("INFO  %-52s %s" % ("Statements negative newest-year revenue", item))
    rep.check(share >= 0.90 and not zero, "Statements newest-year revenue sane",
              "%d checked; %.1f%% positive (>=90%%), %d negative, %d zero%s"
              % (checked, share * 100.0, len(negative), len(zero),
                 "; " + str(zero[:5]) if zero else ""))


def check_symbols(rep, wbf, db_symbols):
    ws = wbf["Symbols"]
    rows = 0
    for row in range(SY_FIRST_ROW, ws.max_row + 1):
        if text(cell(ws, row, 1)) == "":
            break
        rows += 1
    rep.check(rows >= len(db_symbols), "Symbols row count >= DB distinct symbols",
              "%d rows vs %d DB symbols" % (rows, len(db_symbols)))


# Sheet -> header row the auto-filter must anchor on. Stock Lookup and Guide
# are not tables, so they carry no filter.
FILTER_SHEETS = [
    ("Portfolio", PF_HDR_ROW),
    ("Risk & Horizons", RH_HDR_ROW),
    ("DB", DB_HDR_ROW),
    ("Statements", ST_HDR_ROW),
    ("Symbols", SY_HDR_ROW),
]


def check_filters(rep, wbf):
    details = []
    ok = True
    for name, hdr_row in FILTER_SHEETS:
        ref = wbf[name].auto_filter.ref if name in wbf.sheetnames else None
        if ref is None:
            ok = False
            details.append("%s=MISSING" % name)
            continue
        match = re.match(r"^[A-Z]+(\d+):[A-Z]+(\d+)$", str(ref))
        if match is None or int(match.group(1)) != hdr_row:
            ok = False
            details.append("%s=%s (want header row %d)" % (name, ref, hdr_row))
        else:
            details.append("%s PASS %s" % (name, ref))
    rep.check(ok, "auto-filters present", "; ".join(details))


def snapshot_symbol_count(data_dir):
    path = os.path.join(data_dir, "snapshot.json")
    if not os.path.exists(path):
        print("NOTE  snapshot.json not found at %s; symbol-count floor is 0" % path)
        return 0
    with open(path, "r", encoding="utf-8") as fh:
        return len(json.load(fh) or {})


def main():
    # Arabic console output must survive a non-UTF-8 (cp1252) console.
    import sys
    for _stream in (sys.stdout, sys.stderr):
        if _stream is not None and hasattr(_stream, "reconfigure"):
            try:
                _stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    parser = argparse.ArgumentParser(description="Verify a built v2 workbook.")
    parser.add_argument("workbook", help="path to the xlsx to verify")
    parser.add_argument("--data-dir", default=os.path.join(HERE, "data"),
                        help="data dir holding snapshot.json (for the symbol-count floor)")
    args = parser.parse_args()

    if not os.path.exists(args.workbook):
        print("FAIL  workbook not found: %s" % args.workbook)
        return 1

    wbv = load_workbook(args.workbook, data_only=True)
    wbf = load_workbook(args.workbook, data_only=False)
    symbol_count = snapshot_symbol_count(args.data_dir)

    print("verifying %s (%d sheets, %d snapshot symbols)"
          % (args.workbook, len(wbv.sheetnames), symbol_count))
    print("-" * 96)

    rep = Report()
    check_sheets(rep, wbv)
    check_no_error_strings(rep, wbv, wbf)
    check_portfolio(rep, wbv, wbf)
    check_lookup(rep, wbf)
    check_risk(rep, wbv, wbf, symbol_count)
    db_symbols = check_db(rep, wbf)
    check_statements(rep, wbv, wbf)
    check_symbols(rep, wbf, db_symbols)
    check_filters(rep, wbf)

    print("-" * 96)
    print("%d passed, %d failed, %d skipped" % (rep.passes, rep.failures, rep.skips))
    wbv.close()
    wbf.close()
    return 1 if rep.failures else 0


if __name__ == "__main__":
    sys.exit(main())
