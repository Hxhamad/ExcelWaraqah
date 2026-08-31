"""Full refresh: fetch the whole listed universe, then rebuild (~20-40 min).

Same pipeline as refresh_quick.py - the only difference is step 2: every code
in symbols.csv is fetched instead of just the held ones. fetch_all skips codes
already present in annual_metrics.csv, so an interrupted run is resumed simply
by running this again.
"""

import csv
import os
import sys

from fetcher import fetch_all
from refresh_quick import DATA_DIR, HERE, pipeline
from symbols import load_symbols

SYMBOLS_CSV = os.path.join(HERE, "symbols.csv")


def db_rows():
    """Data rows in annual_metrics.csv - the DB sheet's row count."""
    path = os.path.join(DATA_DIR, "annual_metrics.csv")
    if not os.path.exists(path):
        return 0
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        return sum(1 for row in csv.DictReader(fh) if (row.get("symbol") or "").strip())


def main():
    codes = [row["code"] for row in load_symbols(SYMBOLS_CSV)]
    print("fetching %d symbol(s) from %s ..." % (len(codes), SYMBOLS_CSV))

    elapsed = pipeline(lambda: fetch_all(codes, data_dir=DATA_DIR))
    if elapsed is None:
        print("✖ فشل التحديث — لم يتم تعديل الملف")
        return 1

    print("✔ تم تحديث السوق كامل (%d رمز، %d صف في قاعدة البيانات، %.0f دقيقة)"
          % (len(codes), db_rows(), elapsed / 60.0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
