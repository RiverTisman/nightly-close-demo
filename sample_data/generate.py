"""
Generates the fictional demo data for Nightly Close.

Everything here is invented: the restaurant ("Harbor & Vine"), every
employee, every EID, every dollar figure. The files match the real export
formats byte-for-byte in structure (Toast Shift Report CSV, HotSchedules
"Excel XP-2007" SpreadsheetML roster, ADP Punch Source Report CSV, and the
weekly Tip Sheet workbook) so the app runs its real parsing and checks on
them unchanged.

Each file deliberately contains a handful of the real-world problems the
app was built to catch -- an unmatched name, a call-in, a missing
clock-out, someone cut early, a wrong department -- so the demo shows the
checks firing, not just a clean happy path.

Run:  python sample_data/generate.py
"""

import csv
import json
import random
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DAY = datetime(2026, 9, 24)  # a Thursday
random.seed(24)

# (canonical, role, shift(s), shift type note, eid, cash_sales_number)
STAFF = [
    # Servers
    ("Maya Lindqvist", "Server", "AM", "OPENER", "300101", "2001"),
    ("Theo Barrington", "Server", "AM", "OPENER", "300102", "2002"),
    ("Priya Raman", "Server", "AM", "MID", "300103", "2003"),
    ("Jonah Feld", "Server", "AM", "MID", "300104", "2004"),
    ("Camila Ortiz", "Server", "AM", "CLOSER", "300105", "2005"),
    ("Dmitri Volkov", "Server", "PM", "OPENER", "300106", "2006"),
    ("Sasha Kimura", "Server", "PM", "OPENER", "300107", "2007"),
    ("Elena Marchetti", "Server", "PM", "MID", "300108", "2008"),
    ("Kwame Asante", "Server", "PM", "MID", "300109", "2009"),
    ("Lucia Ferreira", "Server", "PM", "CLOSER", "300110", "2010"),
    ("Noah Whitfield", "Server", "PM", "CLOSER", "300111", "2011"),
    ("Tomas Reyes", "Server", "PM", "CLOSER", "300112", None),
    # Bartenders
    ("Grace O'Connell", "Bar", "AM", "OPENER", "300201", None),
    ("Rafael Duarte", "Bar", "PM", "OPENER", "300202", None),
    ("Imani Brooks", "Bar", "PM", "CLOSER", "300203", None),
    # Runners
    ("Leo Castellano", "Runner", "AM", "OPENER", "300301", None),
    ("Ana Sofia Mendez", "Runner", "PM", "OPENER", "300302", None),
    ("Bilal Haddad", "Runner", "PM", "CLOSER", "300303", None),
    ("Wes Tanaka", "Runner", "PM", "CLOSER", "300304", None),
    # Bussers
    ("Jorge Alvarado", "Busser", "AM", "OPENER", "300401", None),
    ("Mateo Silva", "Busser", "PM", "OPENER", "300402", None),
    ("Rosa Delgado", "Busser", "PM", "CLOSER", "300403", None),
    ("Kenji Watanabe", "Busser", "PM", "CLOSER", "300404", None),
    ("Femi Adeyemi", "Busser", "PM", "CLOSER", "300405", None),
    # Barista / Barback
    ("Ines Laurent", "Barista", "AM", "OPENER", "300501", None),
    ("Omar Khalil", "Barback", "PM", "CLOSER", "300601", None),
    ("Dev Patel", "Barback", "PM", "OPENER", "300602", None),
    # Hosts (scheduled, not tipped out)
    ("Clara Nguyen", "Host", "PM", "OPENER", "300701", None),
]

# People in the mapping who don't work tonight, plus one terminated employee.
EXTRA = [
    ("Hannah Becker", "Server", "300113", "2012", "active"),
    ("Victor Lang", "Busser", "300406", None, "terminated"),
    ("Yusuf Demir", "Runner", "300305", None, "active"),
]


def tip_sheet_name(canonical):
    parts = canonical.replace("'", "'").split(" ")
    first, last = parts[0], " ".join(parts[1:])
    if canonical == "Ana Sofia Mendez":
        first, last = "ANA SOFIA", "MENDEZ"
    return f"{last.upper()}, {first.upper()}"


# A Toast or HotSchedules spelling that differs from the canonical name --
# the whole reason the mapping exists.
TOAST_ALIAS = {"Jonah Feld": "jonah feld", "Kwame Asante": "Kwame  Asante", "Noah Whitfield": "Noah Whitfield "}
HS_ALIAS = {"Ana Sofia Mendez": "AnaSofia Mendez", "Grace O'Connell": "Grace OConnell"}


def build_mapping():
    employees = []
    for canonical, role, _shift, _note, eid, csn in STAFF:
        e = {
            "canonical": canonical,
            "tip_sheet": tip_sheet_name(canonical),
            "toast": [TOAST_ALIAS.get(canonical, canonical)] if role in ("Server",) else [],
            "hotschedules": [HS_ALIAS.get(canonical, canonical)],
            "status": "active",
            "eid": eid,
        }
        if csn:
            e["cash_sales_number"] = csn
        employees.append(e)
    for canonical, role, eid, csn, status in EXTRA:
        e = {"canonical": canonical, "tip_sheet": tip_sheet_name(canonical),
             "toast": [canonical] if role == "Server" else [], "hotschedules": [canonical],
             "status": status, "eid": eid}
        if csn:
            e["cash_sales_number"] = csn
        employees.append(e)
    # Drift case for the dropdown check: the workbook's cell carries a
    # trailing space ("REYES, TOMAS ") and the mapping doesn't.
    lines = ",\n".join("    " + json.dumps(e, ensure_ascii=False) for e in employees)
    (ROOT / "employee_mapping.json").write_text('{\n  "employees": [\n' + lines + "\n  ]\n}\n", encoding="utf-8")
    return employees


SHIFT_TIMES = {
    ("AM", "OPENER"): ("10:00", "15:30"), ("AM", "MID"): ("11:00", "15:45"), ("AM", "CLOSER"): ("11:30", "16:30"),
    ("PM", "OPENER"): ("16:00", "22:30"), ("PM", "MID"): ("16:30", "23:00"), ("PM", "CLOSER"): ("17:00", "00:15"),
}


def at(hhmm, base=DAY):
    h, m = map(int, hhmm.split(":"))
    dt = base.replace(hour=h, minute=m)
    if h < 6:
        dt += timedelta(days=1)
    return dt


def jitter(dt, spread=12):
    return dt + timedelta(minutes=random.randint(-spread, spread))


# ---------------------------------------------------------------------------
# Toast Shift Report
# ---------------------------------------------------------------------------

TOAST_HEADER = ["", "Employee", "Job Title", "In Date", "Shift Closed Date", "Out Date", "Hours",
                "Cash Tips Decl.", "Cash on Hand", "Cash in Drawer", "Non-Cash Tips", "Cash Gratuity",
                "Non-Cash Gratuity", "Tips Withheld", "Non-Cash Sales", "Cash Sales", "Total Sales",
                "Cash Collected?", "Tips Paid?", "Open", "Paid", "Closed", "Tip Share Total"]


def toast_fmt(dt):
    return dt.strftime("%m/%d/%y %I:%M %p")


def build_toast():
    rows = []
    rid = 400000000000000000

    def add(name, job, tin, tout, cash_tips, credit_tips, sales, cash_sales, grat=0.0):
        nonlocal rid
        rid += random.randint(1000, 9999)
        hours = round((tout - tin).total_seconds() / 3600, 2)
        rows.append([str(rid), name, job, toast_fmt(tin), toast_fmt(tout), toast_fmt(tout), f"{hours:.2f}",
                     f"{cash_tips:.2f}", "0.00", "0.00", f"{credit_tips:.2f}", "0.00", f"{grat:.2f}", "0.00",
                     f"{sales - cash_sales:.2f}", f"{cash_sales:.2f}", f"{sales:.2f}", "yes", "no", "0", "0", "1", "0.00"])

    for canonical, role, shift, note, _eid, _csn in STAFF:
        if role != "Server":
            continue
        start, end = SHIFT_TIMES[(shift, note)]
        tin, tout = jitter(at(start)), jitter(at(end))
        name = TOAST_ALIAS.get(canonical, canonical)
        sales = round(random.uniform(1400, 3600) if shift == "PM" else random.uniform(700, 1900), 2)
        credit = round(sales * random.uniform(0.17, 0.21), 2)
        cash = round(random.choice([0, 0, 20, 35, 48, 60]), 2)
        cash_sales = round(random.choice([0, 0, 42.5, 88.49, 126.0, 213.5]), 2)
        grat = 0.0
        if canonical == "Elena Marchetti":
            grat = 684.00  # private party -- large auto-gratuity
        if canonical == "Tomas Reyes":
            cash_sales = 61.75  # has cash sales but no cash-sales number mapped
        if canonical == "Noah Whitfield":
            tin, tout = at("19:40"), at("22:55")  # under 4h on dinner
        if canonical == "Camila Ortiz":
            tin = at("12:47")  # near the 1pm AM/PM cutover
        add(name, "Server", tin, tout, cash, credit, sales, cash_sales, grat)

    # A server Toast spells in a way the mapping has never seen.
    add("Hannah Becker-Ross", "Server", at("16:05"), at("23:10"), 40.0, 512.37, 2711.90, 0.0)
    # Zero sales but hours on the clock -- forgot to clock out properly.
    add("Priya Raman", "Server", at("17:30"), at("21:30"), 0.0, 0.0, 0.0, 0.0)
    # Shared bar drawers (not people).
    add("BAR AM", "Bartender", at("10:30"), at("16:00"), 85.0, 642.18, 3180.40, 210.0)
    add("BAR PM", "Bartender", at("16:00"), at("00:30"), 214.0, 1888.62, 9412.75, 540.0)

    with open(HERE / "toast_shift_report.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(TOAST_HEADER)
        w.writerows(rows)


# ---------------------------------------------------------------------------
# HotSchedules daily roster, "Excel XP-2007" (SpreadsheetML XML)
# ---------------------------------------------------------------------------

def _xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def build_hotschedules():
    sections_order = ["Server", "Bar", "Runner", "Busser", "Barista", "Barback", "Host"]
    roster = {"AM": {s: [] for s in sections_order}, "PM": {s: [] for s in sections_order}}
    for canonical, role, shift, note, _eid, _csn in STAFF:
        hs_name = HS_ALIAS.get(canonical, canonical)
        first, _, last = hs_name.partition(" ")
        start = SHIFT_TIMES[(shift, note)][0]
        roster[shift][role].append(["", first, last, f"(917) 555-{random.randint(1000, 9999)}", note,
                                    datetime.strptime(start, "%H:%M").strftime("%I:%M %p").lstrip("0")])

    # The cases a manager has to rule on rather than the app guessing:
    roster["PM"]["Busser"].append(["", "Yusuf", "Demir", "(917) 555-0142", "Called In", "5:00 PM"])
    roster["PM"]["Runner"].append(["[House Shift]", "Victor", "Lang", "(917) 555-0187", "CLOSER", "5:00 PM"])
    roster["AM"]["Runner"].append(["", "Sam", "Okoro", "(917) 555-0199", "TRAINING - NOT LIVE", "11:00 AM"])
    roster["PM"]["Server"].append(["[D]", "Maya", "Lindqvist", "(917) 555-0101", "CLOSER", "5:00 PM"])
    roster["PM"]["Bar"].append(["[House Shift]", "", "", "", "", "5:00 PM"])

    out = ['\t<?xml version="1.0"?>',
           '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" '
           'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">']
    for shift in ("AM", "PM"):
        out.append(f'<Worksheet ss:Name="{shift}"><Table>')
        out.append(f'<Row><Cell ss:StyleID="sHeader"><Data ss:Type="String">{shift}</Data></Cell></Row>')
        for section in sections_order:
            people = roster[shift][section]
            if not people:
                continue
            out.append(f'<Row><Cell ss:StyleID="sHeader"><Data ss:Type="String">{section}</Data></Cell></Row>')
            for cells in people:
                out.append("<Row>" + "".join(
                    f'<Cell ss:StyleID="sData"><Data ss:Type="String">{_xml_escape(c)}</Data></Cell>' for c in cells
                ) + "</Row>")
        out.append("</Table></Worksheet>")
    out.append("</Workbook>")
    (HERE / "hotschedules_roster.xls").write_text("\n".join(out), encoding="utf-8")


# ---------------------------------------------------------------------------
# ADP Punch Source Report
# ---------------------------------------------------------------------------

ADP_HEADER = ["Clock In ID", "Job Code", "Last Name", "First Name", "Position ID", "Status", "Clock In ID",
              "Clock Out ID", "Department", "In Time", "Out Time", "Total", "Hours", "Pay Code"]
JOB_CODES = {"Server": "320", "Bar": "350", "Runner": "370", "Busser": "365", "Barista": "360",
             "Barback": "355", "Host": "310"}


def adp_fmt(dt):
    return dt.strftime("%m/%d/%Y %I:%M:%S %p") if dt else ""


def build_adp():
    rows = []

    def add(canonical, eid, job, tin, tout, status="Active", pay_code=""):
        first, _, last = canonical.partition(" ")
        hours = round((tout - tin).total_seconds() / 3600, 2) if tout else 0.0
        rows.append(["HARBOR", JOB_CODES[job].zfill(5), last, first, f"HV{eid}", status, "HARBOR", "HARBOR",
                     "FOH", adp_fmt(tin), adp_fmt(tout), f"{hours:.2f}", f"{hours:.2f}", pay_code])

    for canonical, role, shift, note, eid, _csn in STAFF:
        start, end = SHIFT_TIMES[(shift, note)]
        tin, tout = jitter(at(start), 6), jitter(at(end), 10)
        job = role
        if canonical == "Bilal Haddad":
            tout = None  # forgot to clock out
        if canonical == "Kenji Watanabe":
            tout = at("21:05")  # sent home early -- possible cut
        if canonical == "Mateo Silva":
            job = "Runner"  # scheduled Busser, clocked in as Runner
        if canonical == "Noah Whitfield":
            tin, tout = at("19:40"), at("22:55")
        if canonical == "Camila Ortiz":
            tin = at("12:47")
        add(canonical, eid, job, tin, tout)
    add("Maya Lindqvist", "300101", "Server", at("16:58"), at("23:40"))  # her double
    add("Yusuf Demir", "300305", "Busser", at("17:04"), at("23:55"))  # the call-in did show
    add("Sam Okoro", "300901", "Runner", at("11:02"), at("15:00"))  # trainee, not in the mapping
    add("Ines Laurent", "300501", "Barista", at("10:00"), at("14:00"), pay_code="TRAINING")

    with open(HERE / "adp_punch_report.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(ADP_HEADER)
        w.writerows(rows)


# ---------------------------------------------------------------------------
# Weekly Tip Sheet workbook (the manager's own Excel file)
# ---------------------------------------------------------------------------

ROLE_RANGES = {"Server": ("A", 8), "Bartender": ("A", 50), "Runner": ("A", 72), "Barback": ("J", 72),
               "Busser": ("A", 94), "Barista": ("J", 94)}


def build_tip_sheet(employees):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    week_start = DAY - timedelta(days=DAY.weekday())
    abbr = ["MON", "TUES", "WED", "THURS", "FRI", "SAT", "SUN"]
    for i, d in enumerate(abbr):
        day = week_start + timedelta(days=i)
        for shift in ("AM", "PM"):
            ws = wb.create_sheet(f"{d} {shift}")
            ws["A1"] = "HARBOR & VINE -- TIP SHEET"
            ws["A4"] = f"{day.month}.{day.day}.{day.year % 100}"
            for role, (col, r0) in ROLE_RANGES.items():
                ws[f"{col}{r0 - 2}"] = role.upper() + "S"

    # Fill Thursday the way the previous night's manager actually did --
    # mostly right, with the kind of slips the Punch Report Check exists
    # to catch.
    filled = {"AM": {r: [] for r in ROLE_RANGES}, "PM": {r: [] for r in ROLE_RANGES}}
    for canonical, role, shift, _note, _eid, _csn in STAFF:
        sheet_role = {"Bar": "Bartender"}.get(role, role)
        if sheet_role == "Host":
            continue
        if canonical == "Wes Tanaka":
            continue  # worked, punched in, left off the sheet
        if canonical == "Dev Patel":
            sheet_role = "Busser"  # wrong section
        filled[shift][sheet_role].append(tip_sheet_name(canonical))
    filled["PM"]["Server"].append(tip_sheet_name("Hannah Becker"))  # on the sheet, never punched in
    filled["PM"]["Server"].append(tip_sheet_name("Maya Lindqvist"))
    for shift, roles in filled.items():
        ws = wb[f"THURS {shift}"]
        for role, names in roles.items():
            col, r0 = ROLE_RANGES[role]
            for i, n in enumerate(names):
                ws[f"{col}{r0 + i}"] = n

    wt = wb.create_sheet("WEEKLY TOTALS")
    wt["A4"], wt["B4"] = "EE ID", "NAME"
    for i, e in enumerate(employees):
        cell_text = e["tip_sheet"]
        if e["canonical"] == "Tomas Reyes":
            cell_text = "REYES , TOMAS"  # workbook cell has a stray space before the comma; mapping does not
        if e["status"] == "terminated":
            continue
        wt[f"A{6 + i}"] = int(e["eid"])
        wt[f"B{6 + i}"] = cell_text
    wb.save(HERE / "tip_sheet_week.xlsx")


if __name__ == "__main__":
    employees = build_mapping()
    build_toast()
    build_hotschedules()
    build_adp()
    build_tip_sheet(employees)
    print("Sample data written to", HERE)
