"""
Parses Toast's daily "Sales Summary" export -- a .zip of ~21 CSVs produced
by one export/download from Toast's reporting screen.

The real workflow (confirmed with River 2026-07-17, and confirmed against
the actual formulas in the real "SALES SUMMARY MASTER" workbook) takes
TWO pulls per day, not one:

- a "lunch cutoff" pull, taken ~4-5pm, used as the lunch numbers as-is
- a "final" pull, taken after everyone's clocked out and the day is closed

Dinner-only figures are the delta between the two. This is not a guess --
the real w 7.12.26.xlsx workbook has cells like
`Dining Room dinner net sales = 24885.1-9282.45`, i.e. final minus lunch,
hardcoded per revenue center.

This module only extracts what's literally in the CSVs. It does not
decide which Toast discount-reason lines count as "Comps" vs "Manager
Meals" vs anything else, and does not compute the lunch/dinner delta --
those are business-logic decisions being mapped cell-by-cell against the
real workbook with River, not guessed here. See the Daily Summary results
page for the current state of that mapping.
"""

import csv
import io
import zipfile

REVENUE_CENTERS_IN_ORDER = ["Dining Room", "Bar", "Patio", "Wine Villa", "Mezzanine", "Events"]

WANTED_FILES = [
    "Revenue center summary.csv",
    "Service Daypart summary.csv",
    "Net sales summary.csv",
    "Service mode summary.csv",
    "Menu Item Discounts.csv",
    "Check Discounts.csv",
    "Void summary.csv",
    "Service charge summary.csv",
    "Cash summary.csv",
    "Payments summary.csv",
    "Tip summary.csv",
    "Dining options summary.csv",
]


def _read_csv_rows(zf, name):
    raw = zf.read(name).decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(raw)))


def parse_sales_summary_zip(zip_path):
    """Returns {csv filename: list of row dicts} for every CSV in the zip
    that this app currently knows how to use. Rows are exactly as Toast
    wrote them -- string values, no type coercion -- so a format change
    shows up as a missing/odd field instead of a silently wrong number."""
    data = {}
    with zipfile.ZipFile(zip_path) as zf:
        names = set(zf.namelist())
        for fname in WANTED_FILES:
            if fname in names:
                data[fname] = _read_csv_rows(zf, fname)
    return data


def revenue_center_breakdown(sales_summary_data):
    """Revenue center summary.csv reordered to the Tip Sheet's fixed
    section order, with the 'Total' row and any 'No Revenue Center' stray
    -items row separated out (that row isn't part of the fixed list and
    needs a human look if it's non-zero, not a silent drop)."""
    rows = sales_summary_data.get("Revenue center summary.csv") or []
    by_name = {r["Revenue center"]: r for r in rows}
    ordered = [(name, by_name.get(name)) for name in REVENUE_CENTERS_IN_ORDER]
    return {
        "ordered": ordered,
        "total": by_name.get("Total"),
        "stray": by_name.get("No Revenue Center"),
    }
