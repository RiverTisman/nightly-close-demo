"""
HotSchedules daily Roster Report -> per-employee, per-shift role assignments.

HotSchedules' "Roster Report" export comes in at least two different
formats, both saved with a `.xls` extension by the export tool, neither of
which is a real binary Excel file:

- Microsoft's SpreadsheetML XML format (openpyxl can't read it; no library
  needs to, the stdlib is enough). Has a stray tab/whitespace sequence
  before the XML declaration, which is invalid per the XML spec, so the raw
  text is stripped before parsing.
- Genuine HTML (a real `<html><style>...<table>` document) -- confirmed
  this is what HotSchedules' "Excel 97-2000" export option actually
  produces, despite the name. Same underlying data and the same
  role-section / [D] / [House Shift] conventions, just markup instead of a
  spreadsheet schema. Parsed with the stdlib `html.parser`, no new
  dependency needed.

parse_daily_roster() sniffs which one it got and dispatches accordingly --
both return the same row shape, so nothing downstream needs to care which
format a given upload was.

Each file covers ONE calendar day (AM and PM). The file itself contains no
date anywhere -- not in the metadata, not in any cell -- so the manager
supplies the date at upload time.
"""

import re
import xml.etree.ElementTree as ET
from html.parser import HTMLParser

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "{urn:schemas-microsoft-com:office:spreadsheet}"

ROLE_MAP = {"Bar": "Bartender"}  # HotSchedules section name -> tip-sheet role name
IGNORED_SECTIONS = {"Host"}      # not tipped out, never enters the cross-check

TRAINEE_RE = re.compile(r"TRAINING|NOT LIVE|\bTEST\b", re.IGNORECASE)
# Was just "CALL IN" -- missed the common past-tense phrasing "Called In"
# (the note describes something that already happened, so managers often
# write it that way), since "ED" breaks up the literal "CALL IN" substring.
# Confirmed this is why on-call confirmation silently didn't fire for some
# real rows. Also covers "CALL-IN"/"CALLIN"/"CALLED-IN" variants.
CALLIN_RE = re.compile(r"CALL(?:ED)?[\s-]*IN", re.IGNORECASE)
DOUBLE_TEXT_RE = re.compile(r"DOUBLE", re.IGNORECASE)

FLAG_RE = re.compile(r"^(\[D\]|\[House Shift\])")


def _cell_text(cell):
    d = cell.find("ss:Data", NS)
    if d is None or not d.text:
        return ""
    return d.text.strip()


def _parse_spreadsheetml_roster(text):
    root = ET.fromstring(text.lstrip())

    rows = []
    for ws in root.findall("ss:Worksheet", NS):
        shift = ws.get(f"{SS}Name")  # "AM" or "PM"
        table = ws.find("ss:Table", NS)
        if table is None:
            continue
        section = None
        for row in table.findall("ss:Row", NS):
            cells = row.findall("ss:Cell", NS)
            if not cells:
                continue
            style = cells[0].get(f"{SS}StyleID")
            if style == "sHeader" and len(cells) == 1:
                text_val = _cell_text(cells[0])
                if text_val not in ("AM", "PM"):
                    section = text_val
                continue
            if len(cells) < 6 or section is None:
                continue
            flag, first, last, phone, note, start = (_cell_text(c) for c in cells[:6])
            rows.append({
                "shift": shift, "section": section, "flag": flag,
                "first": first, "last": last, "phone": phone,
                "note": note, "start": start,
            })
    return rows


class _RosterHTMLParser(HTMLParser):
    """Walks the "Excel 97-2000" HTML export. Role sections are marked by
    a <td class="font-header">ROLE</td>; AM/PM by a
    <th class="font-header2">AM|PM</th>; each employee is a 3-<td> <tr>
    (name+phone, note, start time), with name and phone separated by a
    <br> inside the first cell."""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._shift = None
        self._section = None
        self._cell_class = None
        self._in_cell = False
        self._cell_buf = []
        self._row_cells = []
        # AM and PM are side-by-side nested <table class="column"> elements
        # inside an outer layout table, not sequential blocks -- a flat
        # (non-nesting-aware) <tr>/<td> handler cross-contaminates rows
        # right at the AM/PM boundary. Track a table stack so <tr>/<td>
        # events only count while the innermost table is the data table.
        self._table_stack = []

    def _in_data_table(self):
        return bool(self._table_stack) and self._table_stack[-1]

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "table":
            self._table_stack.append(attrs.get("class") == "column")
        elif tag == "tr" and self._in_data_table():
            self._row_cells = []
        elif tag in ("td", "th") and self._in_data_table():
            self._in_cell = True
            self._cell_buf = []
            self._cell_class = attrs.get("class", "")
        elif tag == "br" and self._in_cell:
            # Sentinel, not "\n" -- the source HTML's own indentation
            # already contains real newlines/tabs *inside* the name text
            # (between first and last name), so a plain "\n" can't be
            # trusted to mark only the <br>-inserted name/phone boundary.
            self._cell_buf.append("\x00")

    def handle_endtag(self, tag):
        if tag == "table":
            if self._table_stack:
                self._table_stack.pop()
        elif tag in ("td", "th") and self._in_cell:
            text = "".join(self._cell_buf)
            if self._cell_class == "font-header2":
                cleaned = re.sub(r"\s+", " ", text).strip()
                if cleaned in ("AM", "PM"):
                    self._shift = cleaned
            elif self._cell_class == "font-header":
                self._section = re.sub(r"\s+", " ", text).strip()
            else:
                self._row_cells.append(text)
            self._in_cell = False
        elif tag == "tr" and self._in_data_table():
            if len(self._row_cells) == 3 and self._section and self._shift:
                self.rows.append((self._shift, self._section, self._row_cells))
            self._row_cells = []

    def handle_data(self, data):
        if self._in_cell:
            self._cell_buf.append(data)


def _parse_html_roster(text):
    parser = _RosterHTMLParser()
    parser.feed(text)

    rows = []
    for shift, section, (name_phone, note, start) in parser.rows:
        name_part, _, phone_part = name_phone.partition("\x00")
        name_part = re.sub(r"\s+", " ", name_part).strip()
        phone = re.sub(r"[^\d]", "", phone_part)
        if len(phone) == 10:
            phone = f"({phone[:3]}) {phone[3:6]}-{phone[6:]}"

        m = FLAG_RE.match(name_part)
        flag = m.group(1) if m else ""
        name = name_part[m.end():].strip() if m else name_part

        first, _, last = name.partition(" ")
        rows.append({
            "shift": shift, "section": section, "flag": flag,
            "first": first, "last": last, "phone": phone,
            "note": re.sub(r"\s+", " ", note).strip(),
            "start": re.sub(r"\s+", " ", start).strip(),
        })
    return rows


class UnsupportedRosterFormat(ValueError):
    pass


def parse_daily_roster(xml_path):
    """Returns a flat list of raw rows: one dict per employee/section row
    across both the AM and PM sheets. No filtering, no role mapping -- just
    what's literally in the file.

    Only the SpreadsheetML export ("Excel XP-2007" in HotSchedules' export
    menu) is accepted. The HTML export ("Excel 97-2000") is fully parseable
    (see _parse_html_roster/_RosterHTMLParser above) and is exercised by
    tests, but it's only ever been run against one real file and needed two
    real structural bugs fixed to get there (nested AM/PM tables, name/phone
    splitting) -- not enough real-world mileage to trust as a silent
    fallback yet. Rejecting it here rather than risking a wrong parse a
    manager wouldn't notice. Re-enable by returning _parse_html_roster(text)
    instead of raising, once that path has more real data behind it.
    """
    with open(xml_path, encoding="utf-8") as f:
        text = f.read()
    stripped = text.lstrip()
    if stripped.startswith("<?xml") or "<Workbook" in stripped[:2000]:
        return _parse_spreadsheetml_roster(text)
    raise UnsupportedRosterFormat(
        "This doesn't look like an Excel XP-2007 export. Please re-export the HotSchedules "
        "Roster Report as Excel XP-2007 (not Excel 97-2000, PDF, Web, or Word) and upload that file."
    )


def build_daily_roster_review(rows, mapping):
    """Turns raw rows into a review list, deciding per-row whether it's a
    clean auto-verified assignment or something that needs a manager's eyes.
    Never silently resolves an ambiguous flag/note/status conflict -- when in
    doubt this flags rather than guesses, per the project's design rule."""
    review = []
    for r in rows:
        section = r["section"]
        if section in IGNORED_SECTIONS:
            continue

        name = f"{r['first']} {r['last']}".strip()
        flag = r["flag"]
        note = r["note"]

        # Fully blank, unclaimed house shift: no effect on our systems, per
        # design -- nobody worked it, nothing to check.
        if flag == "[House Shift]" and not name:
            continue

        role = ROLE_MAP.get(section, section)
        entry = {
            "shift": r["shift"], "role": role, "employee": name,
            "flag": flag, "note": note, "start": r["start"],
            "needs_review": False, "review_reasons": [],
        }

        if flag == "[House Shift]" and name:
            entry["needs_review"] = True
            entry["review_reasons"].append(
                "House Shift with a name attached -- confirm this person actually worked it before including in the tip pool"
            )

        if TRAINEE_RE.search(note):
            entry["needs_review"] = True
            entry["review_reasons"].append(
                f"Note suggests trainee/not-live/test entry ({note!r}) -- confirm before including in or excluding from tip pool"
            )

        # A plain [D] with an ordinary note (e.g. "CLOSER") is the normal
        # case -- most doubles don't repeat "double" in the note, so that is
        # NOT treated as a mismatch. The real risk called out by the
        # manager is call-ins specifically: whether a call-in actually got
        # worked as a double isn't reliable from the export alone.
        if CALLIN_RE.search(note):
            entry["needs_review"] = True
            entry["review_reasons"].append(
                f"Call-in note ({note!r}) -- confirm whether this shift was actually worked before counting it in the tip pool"
            )
        if DOUBLE_TEXT_RE.search(note) and flag != "[D]":
            entry["needs_review"] = True
            entry["review_reasons"].append(
                f"Note says double ({note!r}) but the [D] flag isn't set -- confirm shift count"
            )

        if name:
            status = mapping.hotsched_status(name)
            if status == "terminated":
                entry["needs_review"] = True
                entry["review_reasons"].append(
                    f"TERMINATED EMPLOYEE APPEARS IN DATA: {name} -- do not include in tip pool without manager confirmation"
                )
            elif status is None:
                entry["needs_review"] = True
                entry["review_reasons"].append(
                    f"Unmatched name: {name!r} -- not found in employee_mapping.json, add once confirmed"
                )

        review.append(entry)
    return review
