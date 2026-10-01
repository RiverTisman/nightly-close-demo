"""
HotSchedules "Weekly Roster" CSV export -> the week's staff list.

Different export from the daily Roster Report that Daily Close uses
(parse_hotschedules_daily.py): one row per employee per posted schedule
line, one column group per weekday. It's a plain CSV, and its Employee
column carries names in exactly the spelling the daily export uses too --
which is the only reason this module exists.

The Add Employee page needs a list of real HotSchedules name strings to
offer for selection, so that a mapping's `hotschedules` alias is picked
from the export rather than typed from memory. Typing it from memory is
what produced "DANA KOWALSKY" in the real mapping -- a transposition
nobody spots by eye, which quietly stopped matching Dana on every close
until an unmatched-name flag finally surfaced it.

A weekly export is the better file to ask a manager for here: it covers
seven days rather than one, so a new hire who worked Tuesday is in it
even if you're setting them up on Friday.
"""

import csv

WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

# HotSchedules section/job name -> Tip Sheet role name. Same direction as
# parse_hotschedules_daily.ROLE_MAP, which maps its "Bar" section header;
# this export spells the job out per shift instead.
JOB_TO_ROLE = {
    "Server": "Server",
    "Bartender": "Bartender",
    "Bar": "Bartender",
    "Runner": "Runner",
    "Barback": "Barback",
    "Busser": "Busser",
    "Barista": "Barista",
}


class NotAWeeklyRoster(ValueError):
    pass


def looks_like_weekly_roster(path):
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            header = f.readline()
    except (OSError, UnicodeDecodeError):
        return False
    return "Employee" in header and "Mon Job" in header


def parse_weekly_roster(path):
    """[{"name", "jobs": [role, ...], "roles": [tip sheet role, ...]}, ...]
    one entry per distinct Employee string, sorted by name.

    Names are returned byte-for-byte as the export writes them -- no
    case folding, no whitespace collapsing, no "Last, First" rewriting.
    The whole point is to hand back a string that will match on a future
    upload, so anything this module tidied up would be a string that
    never appears in a real export again.
    """
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or "Employee" not in reader.fieldnames:
            raise NotAWeeklyRoster(
                "This CSV has no Employee column -- confirm it's the HotSchedules Weekly Roster "
                "export and not a different report."
            )
        seen = {}
        for row in reader:
            name = (row.get("Employee") or "").strip()
            if not name:
                continue
            jobs = seen.setdefault(name, set())
            for day in WEEKDAYS:
                job = (row.get(f"{day} Job") or "").strip()
                if job and job != "-":
                    jobs.add(job)

    out = []
    for name in sorted(seen):
        jobs = sorted(seen[name])
        roles = sorted({JOB_TO_ROLE[j] for j in jobs if j in JOB_TO_ROLE})
        out.append({"name": name, "jobs": jobs, "roles": roles})
    return out
