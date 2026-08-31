"""Symbol universe loader for the Saudi stock analysis workbook.

The workbook is driven by a small CSV of Tadawul listings; this module reads it
and refuses to hand downstream code anything that is not a real 4-digit code.
"""

import csv
import os

FIELDS = ("code", "name_ar", "name_en", "sector")


def _is_code(value):
    """A Tadawul code is exactly four digits, e.g. '2222'."""
    return isinstance(value, str) and len(value) == 4 and value.isdigit()


def load_symbols(path="symbols.csv"):
    """Read `path` and return a list of {code, name_ar, name_en, sector} dicts.

    Codes are stripped before validation so a stray space or a spreadsheet's
    trailing newline is not treated as a bad row. Raises ValueError listing
    every offending row (1-based, matching the CSV's data rows) when any code
    is not four digits.
    """
    if not os.path.exists(path):
        raise FileNotFoundError("symbols file not found: %s" % path)

    rows = []
    bad = []
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        missing = [f for f in FIELDS if f not in (reader.fieldnames or [])]
        if missing:
            raise ValueError(
                "symbols file %s is missing columns: %s" % (path, ", ".join(missing))
            )
        for i, raw in enumerate(reader, start=1):
            row = {f: (raw.get(f) or "").strip() for f in FIELDS}
            if not any(row.values()):
                continue  # blank line
            if not _is_code(row["code"]):
                bad.append("row %d: code=%r" % (i, raw.get("code")))
                continue
            rows.append(row)

    if bad:
        raise ValueError(
            "invalid 4-digit codes in %s:\n  %s" % (path, "\n  ".join(bad))
        )
    return rows


def load_codes(path="symbols.csv"):
    """Convenience: just the codes, in file order."""
    return [r["code"] for r in load_symbols(path)]
