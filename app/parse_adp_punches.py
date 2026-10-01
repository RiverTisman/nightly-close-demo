"""
ADP "Punch Source Report" CSV -> per-punch records.

The export has a real quirk: the header lists "Clock In ID" twice (columns
0 and 6), and a third similarly-named "Clock Out ID" (column 7) -- all three
turn out to just be the store/location name ("the store's location name", or
occasionally an internal code like "PTOGN"/"TCMGR" for non-regular punch
types), not per-punch identifiers. Reading this with csv.DictReader would
silently drop column 0's data (Python keeps the last value for a repeated
key), so this reads by position instead.

The real per-employee identifier is "Position ID" (format a company-code prefix + digits).
Confirmed against the real Tip Sheet's WEEKLY TOTALS sheet: this is exactly
the company code prefixed onto that sheet's "EE ID" column -- an exact join key, no
name-matching needed once employee_mapping.json has "eid" populated (see
parse_toast/hotsched name-matching for how identity resolution works when
there's no shared ID).
"""

import csv
from datetime import datetime

# Tipped-role job codes only -- confirmed against a real export with the
# restaurant manager. Every other code seen in real data (335, 345, 380,
# 905, 915, 920, 925, 930, 940, 945, 970, 975, 980, 990) is a kitchen role
# and not relevant to the tip pool.
JOB_CODE_ROLES = {
    "310": "Host",
    "312": "Maitre D",
    "315": "Server Captain",
    "320": "Server",
    "350": "Bartender",
    "355": "Barback",
    "360": "Barista",
    "365": "Busser",
    "370": "Runner",
}

_TIME_FMT = "%m/%d/%Y %I:%M:%S %p"


def _parse_time(s):
    s = s.strip()
    if not s:
        return None
    return datetime.strptime(s, _TIME_FMT)


def parse_adp_punches(csv_path):
    """Returns a flat list of punch dicts, one per row in the export.

    Fields: eid, last, first, job_code, job_role (None if not a tipped
    role), status (Active/Inactive per ADP, NOT the same concept as this
    app's own employee_mapping.json status -- don't conflate them), in_time
    and out_time (datetime or None if missing), hours (float), pay_code
    (blank for a regular worked shift; SICK/VACATION/UNPAID HOURS/TRAINING
    etc. otherwise), date (the punch's calendar date as "MM/DD/YYYY", or
    None if in_time is missing).
    """
    punches = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        next(reader)  # header -- read positionally, see module docstring
        for row in reader:
            if not row or len(row) < 14:
                continue
            eid = row[4].strip()
            # Strip the company-code letters ADP prefixes onto the number.
            eid = eid.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
            job_code = row[1].strip().lstrip("0") or "0"
            in_time = _parse_time(row[9])
            out_time = _parse_time(row[10])
            punches.append({
                "eid": eid,
                "last": row[2].strip(),
                "first": row[3].strip(),
                "job_code": job_code,
                "job_role": JOB_CODE_ROLES.get(job_code),
                "status": row[5].strip(),
                "in_time": in_time,
                "out_time": out_time,
                "hours": float(row[12]) if row[12].strip() else 0.0,
                "pay_code": row[13].strip(),
                "date": in_time.strftime("%m/%d/%Y") if in_time else None,
            })
    return punches
