"""
The guided Add Employee flow: reconcile employee_mapping.json against
the two real files, then add someone using only values taken from them.

The problem this replaces
-------------------------
The old form asked a manager to hand-type four strings that each had to
match a different external system byte-for-byte -- the Tip Sheet cell,
the ADP EID, the HotSchedules name, the Toast name -- with no way to
check any of them at the moment of typing. That's not a form, it's a
memory test with a silent failure mode: get the Tip Sheet text wrong and
Excel's SUMIF returns $0 for that person's whole night, with no error
anywhere. Two live examples were sitting in the real data when this was
written:

  * "DANA KOWALSKY" in the mapping vs. "Dana Kowalski" in every
    HotSchedules export -- a transposed pair of letters, so Dana came up
    "unmatched" on the daily roster every single night.
  * "Marco R Delvin Abernathy" in HotSchedules vs. "DELGIN ABERNATHY,
    MARCO" in the workbook -- Delvin/Delgin. A person reads straight
    past that. A byte comparison doesn't.

Here, every value is either selected from one of the uploaded files or,
when the person genuinely isn't in the workbook yet, written INTO the
workbook by this app -- so the mapping and the spreadsheet are two
copies of one Python string rather than two independent typings of one
name. That's the one-to-one guarantee the whole page is built around.

Nothing in this module ever matches names approximately. Similarity is
used in exactly one place -- ordering the candidate list a human picks
from -- and a candidate is never preselected, so a wrong guess costs a
scroll, not a payout.
"""

import difflib
import re

from .logic import EmployeeMapping
from .parse_hotschedules_daily import parse_daily_roster, UnsupportedRosterFormat
from .parse_hotschedules_weekly import (
    looks_like_weekly_roster, parse_weekly_roster, NotAWeeklyRoster, JOB_TO_ROLE,
)
from .tip_sheet_writer import (
    ROLE_ORDER, TipSheetWriteError, audit_formula_rows, read_workbook_sections,
    suggested_tip_sheet_name,
)


def read_roster_names(path):
    """[{"name", "jobs", "roles"}, ...] from either HotSchedules export
    the app understands -- the Weekly Roster CSV or the daily Roster
    Report ("Excel XP-2007"). Which one it is, is sniffed from the file
    rather than from the filename, since HotSchedules saves both with
    unhelpful extensions."""
    if looks_like_weekly_roster(path):
        return parse_weekly_roster(path)
    rows = parse_daily_roster(path)  # raises UnsupportedRosterFormat with its own message
    seen = {}
    for r in rows:
        name = f"{r['first']} {r['last']}".strip()
        if name:
            seen.setdefault(name, set()).add(r["section"])
    out = []
    for name in sorted(seen):
        jobs = sorted(seen[name])
        out.append({"name": name, "jobs": jobs,
                    "roles": sorted({JOB_TO_ROLE[j] for j in jobs if j in JOB_TO_ROLE})})
    return out


def analyze(mapping_path, workbook_path=None, roster_path=None):
    """Everything the page needs to show, given whichever files were
    uploaded. Returns a dict with an `errors` list rather than raising --
    a bad roster file shouldn't hide a perfectly readable workbook's
    findings, or vice versa.

    The four reconciliation buckets, in the order they matter:

      workbook_unmapped   in the Tip Sheet, missing from the mapping.
                          The app can't see these people at all, so
                          every close flags them as unmatched. One click
                          each to fix, with the name copied out of the
                          sheet.
      roster_unmatched    scheduled in HotSchedules, matching no mapping
                          alias. Same visible symptom, but the fix is
                          usually attaching an alias to somebody who's
                          already mapped, not adding a new person.
      drift               mapped and in the sheet, but the two strings
                          differ. This is the dangerous one: it looks
                          fine on every screen and silently zeroes the
                          SUMIF.
      mapping_not_in_workbook  active in the mapping, absent from the
                          sheet. Usually someone who left and was never
                          marked terminated.
    """
    mapping = EmployeeMapping(mapping_path)
    out = {
        "errors": [],
        "has_workbook": False,
        "has_roster": False,
        "sections": {},
        "formula_audit": [],
        "workbook_unmapped": [],
        "roster_names": [],
        "roster_unmatched": [],
        "drift": [],
        "mapping_not_in_workbook": [],
        "role_order": ROLE_ORDER,
    }

    sections = None
    if workbook_path:
        try:
            sections = read_workbook_sections(workbook_path)
            out["formula_audit"] = audit_formula_rows(workbook_path)
            out["has_workbook"] = True
        except TipSheetWriteError as e:
            out["errors"].append(f"Tip Sheet workbook: {e}")
        except Exception as e:
            out["errors"].append(f"Tip Sheet workbook couldn't be read ({type(e).__name__}: {e}).")

    roster = []
    if roster_path:
        try:
            roster = read_roster_names(roster_path)
            out["roster_names"] = roster
            out["has_roster"] = True
        except (UnsupportedRosterFormat, NotAWeeklyRoster) as e:
            out["errors"].append(f"HotSchedules roster: {e}")
        except Exception as e:
            out["errors"].append(f"HotSchedules roster couldn't be read ({type(e).__name__}: {e}).")

    if sections:
        out["sections"] = {
            role: {
                "range": s["range"],
                "occupied_count": len(s["occupied"]),
                "free_count": len(s["free_rows"]),
                "blocked_count": len(s["broken_rows"]),
            }
            for role, s in sections.items()
        }

        # One person can hold two roles (one employee is a Server and a
        # Bartender), so the same EID legitimately appears in two
        # sections with the same name. Group by EID and keep the roles,
        # rather than reporting them as two separate people to add.
        by_eid = {}
        for role in ROLE_ORDER:
            for o in sections.get(role, {}).get("occupied", []):
                if not o["eid"]:
                    continue
                rec = by_eid.setdefault(
                    EmployeeMapping._norm_eid(o["eid"]),
                    {"eid": o["eid"], "names": [], "roles": [], "rows": []},
                )
                rec["roles"].append(role)
                rec["rows"].append(o["row"])
                if o["name"] not in rec["names"]:
                    rec["names"].append(o["name"])

        for eid_norm, rec in sorted(by_eid.items(), key=lambda kv: kv[1]["names"][0]):
            entry = mapping.eid_entry(rec["eid"])
            if entry is None:
                rec["roles"] = sorted(set(rec["roles"]), key=ROLE_ORDER.index)
                rec["suggested_canonical"] = canonical_from_tip_sheet(rec["names"][0])
                out["workbook_unmapped"].append(rec)
                continue
            # Same person written two different ways in two sections is
            # its own silent-SUMIF bug, so surface it as drift too.
            for sheet_name in rec["names"]:
                if entry.get("tip_sheet") != sheet_name:
                    out["drift"].append({
                        "canonical": entry["canonical"], "eid": rec["eid"],
                        "mapped": entry.get("tip_sheet"), "real": sheet_name,
                        "roles": sorted(set(rec["roles"]), key=ROLE_ORDER.index),
                    })

        for e in mapping.entries:
            if e.get("status", "active") != "active" or not e.get("eid"):
                continue
            if EmployeeMapping._norm_eid(e["eid"]) not in by_eid:
                out["mapping_not_in_workbook"].append({"canonical": e["canonical"], "eid": e["eid"],
                                                       "tip_sheet": e.get("tip_sheet")})

    if roster:
        # Candidates for an unmatched roster name are ordered by string
        # similarity purely so the right one tends to be near the top of
        # a 120-name list. Nothing is preselected and nothing is applied
        # without a person choosing it -- see the module docstring.
        unmapped_pool = [
            {"eid": rec["eid"], "name": rec["names"][0],
             "roles": sorted(set(rec["roles"]), key=ROLE_ORDER.index)}
            for rec in out["workbook_unmapped"]
        ]
        # The other half of "unmatched name" -- and the half the old form
        # had no answer for at all. Dana Kowalski is in the mapping,
        # in the workbook, and correct in both; his HotSchedules alias
        # was just typed as "DANA KOWALSKY". Re-adding her is refused
        # as a duplicate, so the only fix used to be hand-editing the
        # JSON. Offering existing employees as candidates turns it into
        # "attach this spelling to this person."
        existing_pool = [
            {"eid": e.get("eid"), "name": e["canonical"], "tip_sheet": e.get("tip_sheet"),
             "aliases": e.get("hotschedules", [])}
            for e in mapping.entries if e.get("status", "active") == "active"
        ]
        for r in roster:
            if mapping.hotsched_status(r["name"]) is not None:
                continue
            out["roster_unmatched"].append({
                "name": r["name"],
                "jobs": r["jobs"],
                "roles": r["roles"],
                "candidates": _rank_candidates(r["name"], unmapped_pool),
                "existing_candidates": _rank_candidates(
                    r["name"], existing_pool, key_fields=("name", "tip_sheet")
                ),
            })

    return out


def _rank_candidates(roster_name, pool, limit=5, key_fields=("name",)):
    """The pool entries most textually similar to a HotSchedules name,
    best first. An ordering aid for a human's dropdown -- never a match.

    Compared on sorted word sets so "Marco R Delvin Abernathy" and
    "DELGIN ABERNATHY, MARCO" score as the near-miss they are (0.91)
    instead of being pushed apart by word order and the comma. Scoring
    across several fields and keeping the best takes the same care of an
    entry whose canonical name and Tip Sheet text are formatted
    differently from each other.
    """
    def key(text):
        return " ".join(sorted(re.findall(r"[a-z]+", (text or "").lower())))

    target = key(roster_name)
    scored = []
    for c in pool:
        ratio = max(
            difflib.SequenceMatcher(None, target, key(c.get(f))).ratio() for f in key_fields
        )
        scored.append((ratio, c))
    scored.sort(key=lambda x: -x[0])
    return [dict(c, similarity=round(ratio, 3)) for ratio, c in scored[:limit]]


def canonical_from_tip_sheet(name):
    """A first draft of the app's display name from a workbook cell:
    "DELGIN ABERNATHY, MARCO" -> "Marco Delgin Abernathy".

    Purely cosmetic. `canonical` is the one field in an entry that
    doesn't have to match anything outside this app, so getting it
    slightly wrong costs nothing and it's shown in an editable box
    anyway. Every field that DOES have to match is taken from a file.
    """
    text = (name or "").strip()
    if not text:
        return ""
    last, _, first = text.partition(",")
    parts = [p for p in (first.strip(), last.strip()) if p]
    return " ".join(w.capitalize() for w in " ".join(parts).split())


def build_entry(canonical, eid, tip_sheet, hotschedules, toast, cash_sales_number):
    """One mapping entry in the file's established key order. The single
    place an entry is shaped, so the add and edit paths can't drift into
    writing different key orders for the same thing."""
    entry = {
        "canonical": canonical.strip(),
        # Never stripped: a real trailing space in the workbook cell is
        # significant, and this string has to stay identical to it.
        "tip_sheet": tip_sheet,
        "toast": list(dict.fromkeys(n.strip() for n in toast if n and n.strip())),
        "hotschedules": list(dict.fromkeys(n.strip() for n in hotschedules if n and n.strip())),
        "status": "active",
        "eid": str(eid).strip(),
    }
    if cash_sales_number and str(cash_sales_number).strip():
        entry["cash_sales_number"] = str(cash_sales_number).strip()
    return entry


def preview_new_workbook_name(full_name, workbook_path, role):
    """The "LAST, FIRST" text this app would write into the workbook for
    a typed-in name, shown for confirmation before anything is written.
    Falls back to a plain conversion if the workbook can't be read -- the
    manager confirms the final string either way."""
    existing = []
    try:
        sections = read_workbook_sections(workbook_path)
        existing = [o["name"] for o in sections.get(role, {}).get("occupied", [])]
    except Exception:
        pass
    return suggested_tip_sheet_name(full_name, existing)
