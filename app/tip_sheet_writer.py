"""
Writes a new employee into the real Tip Sheet workbook's WEEKLY TOTALS
sheet, and reads back what's already there.

Why this exists
---------------
Every other name-matching bug in this project traces back to the same
shape: the workbook says one thing, employee_mapping.json says another,
and Excel's SUMIF silently returns $0 for the difference. The mapping's
`tip_sheet` field was being hand-typed to match a cell nobody could see
while typing it -- "DANA KOWALSKY" vs. the sheet's "KOWALSKI, DANA"
sat in the real mapping for weeks, and "Delvin"/"Delgin" would have been
next.

So this module inverts the direction: instead of typing a name and
hoping it matches the workbook, the app *writes the name into the
workbook itself* and stores that exact same string in the mapping. Both
sides come from one Python string, so they cannot disagree. That's the
one-to-one guarantee -- structural, not a copy/paste discipline.

How the write stays safe
------------------------
Two properties of the real workbook (verified against
NEW_MASTER_TIP_SHEET.xlsx) make this a two-cell edit rather than a
surgery:

1. Each role's dropdown pulls from a fixed WEEKLY TOTALS column-B range
   that is deliberately longer than the roster -- e.g. Servers occupy
   rows 6-49 inside a B6:B82 source range. Those trailing rows are empty
   *slots*, already built.
2. Every slot row already carries its full set of per-day SUMIF
   formulas, styles and borders, pointing at its own B cell. Row 50 is a
   complete server row that happens to have no name in it.

So adding someone is: put the EID in column A, the name in column B, of
the first free slot in that role's section. No inserted rows, no shifted
formulas, no touched validation ranges, no re-sorted sections. The name
appears in every day tab's dropdown the moment the file opens.

Everything is done on the raw XML inside the .xlsx zip, NOT through
openpyxl. openpyxl cannot round-trip this workbook: its dropdowns are
x14 extension data validations (they reference another sheet), and
openpyxl announces on load that it drops them -- "Data Validation
extension is not supported and will be removed". Saving through openpyxl
would hand the manager back a workbook with every dropdown silently
gone. Here, every zip entry except the two we edit is copied through
byte-for-byte.
"""

import re
import zipfile
from io import BytesIO

from .parse_tip_sheet_daily import ROLE_RANGES

WEEKLY_TOTALS_SHEET = "WEEKLY TOTALS"

# Display/section order, matching the workbook's own top-to-bottom layout.
ROLE_ORDER = ["Server", "Bartender", "Runner", "Barback", "Busser", "Barista"]

# A day tab's dropdown cell block, keyed by the (column, first row) its
# sqref starts at -- how a discovered x14 validation is identified as
# "this is the Busser name column." Derived from ROLE_RANGES so the two
# can't drift apart.
_ROLE_BY_SQREF_START = {(col, start): role for role, (col, start, _end) in ROLE_RANGES.items()}


class TipSheetWriteError(ValueError):
    """Anything that would make writing into this workbook unsafe or
    ambiguous. Always raised instead of guessing -- same rule as the rest
    of the app."""


# ---------------------------------------------------------------------------
# Reading the workbook's own structure
# ---------------------------------------------------------------------------

def _sheet_part_name(zf, sheet_name):
    """Locates the sheetN.xml part backing a named sheet, by following
    workbook.xml's r:id into workbook.xml.rels -- the sheet's position in
    the tab bar is NOT its file number (WEEKLY TOTALS is the 15th tab but
    sheet15.xml only by coincidence), so this never guesses from order."""
    workbook = zf.read("xl/workbook.xml").decode("utf-8")
    m = re.search(
        r'<sheet[^>]*\bname="%s"[^>]*\br:id="([^"]+)"' % re.escape(sheet_name), workbook
    )
    if not m:
        m = re.search(
            r'<sheet[^>]*\br:id="([^"]+)"[^>]*\bname="%s"' % re.escape(sheet_name), workbook
        )
    if not m:
        raise TipSheetWriteError(
            f'No "{sheet_name}" sheet in this workbook -- confirm this is the real Tip Sheet '
            "master workbook, not another file."
        )
    rel_id = m.group(1)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rm = re.search(r'<Relationship[^>]*\bId="%s"[^>]*\bTarget="([^"]+)"' % re.escape(rel_id), rels)
    if not rm:
        raise TipSheetWriteError(f'Workbook relationship {rel_id} for "{sheet_name}" is missing.')
    target = rm.group(1).lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def _shared_strings(zf):
    """[str, ...] indexed by shared-string index. Trailing spaces are
    significant here (several real names have one), so xml:space is
    irrelevant to reading -- the raw text between <t> and </t> is taken
    exactly as-is."""
    try:
        raw = zf.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for si in re.findall(r"<si>(.*?)</si>", raw, re.S):
        # A shared string can be split into several runs (<r><t>..</t></r>);
        # concatenating every <t> is correct for both shapes.
        out.append("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
    return [_unescape(t) for t in out]


def _unescape(text):
    return (text.replace("&lt;", "<").replace("&gt;", ">")
                .replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&"))


def _escape(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace('"', "&quot;"))


def discover_role_sections(zf):
    """{role: (first_row, last_row)} -- each role's dropdown source range
    in WEEKLY TOTALS column B, read out of the workbook's own data
    validations rather than hardcoded here.

    The day tabs declare these as x14 extension validations (a validation
    whose source list lives on another sheet has to be an extension), e.g.
    `<xm:f>'WEEKLY TOTALS'!$B$196:$B$262</xm:f>` against `<xm:sqref>A94:A124`
    -- and A94 is the Busser name block, so B196:B262 is the Busser
    section. Reading it this way means the workbook can grow a section
    without this code needing an edit, and a workbook whose layout has
    genuinely changed fails loudly here instead of writing a name into
    the wrong role."""
    for part in sorted(n for n in zf.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)):
        xml = zf.read(part).decode("utf-8")
        if "x14:dataValidation" not in xml:
            continue
        sections = {}
        for dv in re.findall(r"<x14:dataValidation\b.*?</x14:dataValidation>", xml, re.S):
            fm = re.search(
                r"<xm:f>'%s'!\$B\$(\d+):\$B\$(\d+)</xm:f>" % re.escape(WEEKLY_TOTALS_SHEET), dv
            )
            sm = re.search(r"<xm:sqref>([A-Z]+)(\d+)", dv)
            if not fm or not sm:
                continue
            role = _ROLE_BY_SQREF_START.get((sm.group(1), int(sm.group(2))))
            if role:
                sections[role] = (int(fm.group(1)), int(fm.group(2)))
        if len(sections) == len(ROLE_RANGES):
            return sections
    raise TipSheetWriteError(
        "Couldn't find this workbook's own name dropdowns on any day tab, so there's no safe "
        "way to tell which WEEKLY TOTALS rows belong to which role. Confirm this is the real "
        "Tip Sheet master workbook (the one whose Name columns are dropdowns)."
    )


_ROW_RE = re.compile(r'<row r="(\d+)"[^>]*(?:/>|>.*?</row>)', re.S)


def _rows_by_number(sheet_xml):
    return {int(m.group(1)): m.group(0) for m in _ROW_RE.finditer(sheet_xml)}


# Attributes are matched as explicit name="value" pairs rather than
# [^>]*, which would swallow the "/" of a self-closing <c .../> tag and
# then run on into the NEXT cell looking for a closing </c> -- silently
# reading a neighbouring cell's value as this one's. That bug read every
# empty slot row as occupied.
_CELL_RE_TMPL = r'<c r="%s"(?P<attrs>(?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>(?P<body>.*?)</c>)'


def _cell_match(row_xml, ref):
    return re.search(_CELL_RE_TMPL % re.escape(ref), row_xml, re.S)


def _cell_xml(row_xml, ref):
    m = _cell_match(row_xml, ref)
    return m.group(0) if m else None


def _cell_value(cell_xml, shared):
    """The cell's text, resolving a shared-string reference. Returns "" for
    an empty or absent cell."""
    if not cell_xml:
        return ""
    vm = re.search(r"<v>(.*?)</v>", cell_xml, re.S)
    if not vm:
        return ""
    raw = vm.group(1)
    if 't="s"' in cell_xml:
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return ""
    if 't="inlineStr"' in cell_xml:
        im = re.search(r"<t[^>]*>(.*?)</t>", cell_xml, re.S)
        return _unescape(im.group(1)) if im else ""
    return _unescape(raw)


def _formula_row_refs(row_xml):
    """Which WEEKLY TOTALS row numbers this row's SUMIF formulas actually
    look their name up in. For a correctly built row this is only its own
    row number."""
    return set(int(n) for n in re.findall(r"'%s'!B(\d+)" % re.escape(WEEKLY_TOTALS_SHEET), row_xml))


def read_workbook_sections(xlsx_path):
    """Everything the Add Employee page needs to know about the uploaded
    workbook, per role:

      {role: {"range": (first, last),
              "occupied": [{"row", "eid", "name"}, ...],
              "free_rows": [row, ...],
              "broken_rows": [{"row", "refs"}, ...]}}

    `free_rows` are slots that are empty AND whose formulas point at
    their own row -- the only rows this module will ever write into.
    `broken_rows` are empty slots disqualified because their formulas
    look up a different row's name; writing there would silently total
    somebody else's money into this person's line (see audit_formula_rows).
    """
    with zipfile.ZipFile(xlsx_path) as zf:
        sections = discover_role_sections(zf)
        part = _sheet_part_name(zf, WEEKLY_TOTALS_SHEET)
        sheet_xml = zf.read(part).decode("utf-8")
        shared = _shared_strings(zf)

    rows = _rows_by_number(sheet_xml)
    out = {}
    for role in ROLE_ORDER:
        if role not in sections:
            continue
        first, last = sections[role]
        occupied, free, broken = [], [], []
        for r in range(first, last + 1):
            row_xml = rows.get(r)
            if row_xml is None:
                continue
            name = _cell_value(_cell_xml(row_xml, f"B{r}"), shared)
            if name.strip():
                occupied.append({
                    "row": r,
                    "eid": _cell_value(_cell_xml(row_xml, f"A{r}"), shared),
                    "name": name,
                })
                continue
            refs = _formula_row_refs(row_xml)
            if not refs:
                # No SUMIF formulas at all -- a spacer/header row that
                # happens to fall inside the dropdown range (e.g. the
                # blank line under a section title). Writing a name here
                # would put it in the dropdown while totalling nothing.
                continue
            if refs != {r}:
                broken.append({"row": r, "refs": sorted(refs)})
            elif _cell_xml(row_xml, f"B{r}") is not None:
                free.append(r)
        out[role] = {"range": (first, last), "occupied": occupied,
                     "free_rows": free, "broken_rows": broken}
    return out


def audit_formula_rows(xlsx_path):
    """Flags every WEEKLY TOTALS row whose own SUMIF formulas look up a
    DIFFERENT row's name cell.

    This is the same failure the whole project exists to prevent -- a
    silently wrong SUMIF -- but coming from the workbook's own structure
    rather than from a name mismatch, so the existing dropdown check
    can't see it. A row like this pays out whatever the row above it
    earned, and the affected person's real tips total to nothing
    anywhere. Found for real in NEW_MASTER_TIP_SHEET.xlsx: the entire
    BARISTA section is shifted one row, so the first Barista reads a blank cell ($0)
    and every Barista under her reads the Barista above her.

    Returns [{"role", "row", "name", "refs"}, ...] -- occupied rows first,
    since those are actively wrong right now, then empty slots, which are
    only a trap for the next person added.
    """
    with zipfile.ZipFile(xlsx_path) as zf:
        sections = discover_role_sections(zf)
        sheet_xml = zf.read(_sheet_part_name(zf, WEEKLY_TOTALS_SHEET)).decode("utf-8")
        shared = _shared_strings(zf)

    rows = _rows_by_number(sheet_xml)

    def locate(r):
        """(role, is_inside) for a row -- the section it belongs to, or
        the nearest one above it if it has drifted past the end of a
        dropdown range."""
        for role, (first, last) in sections.items():
            if first <= r <= last:
                return role, True
        above = [(first, role) for role, (first, _last) in sections.items() if first <= r]
        return (max(above)[1] if above else None), False

    # Scanned across the WHOLE sheet, not just inside the dropdown
    # ranges. A wrong reference one row past the end of a range is just
    # as capable of totalling the wrong person's money, and being
    # outside the range is not protection -- a dropdown can't put a name
    # there, but typing can. Verified on the real workbook: widening
    # this found exactly one more broken row (285, immediately after the
    # Barista range) and zero false positives, because every correctly
    # built row on this sheet references only itself.
    findings = []
    for r in sorted(rows):
        refs = _formula_row_refs(rows[r])
        if not refs or refs == {r}:
            continue
        role, inside = locate(r)
        findings.append({
            "role": role, "row": r, "in_section": inside,
            "name": _cell_value(_cell_xml(rows[r], f"B{r}"), shared),
            "refs": sorted(refs),
        })
    findings.sort(key=lambda f: (not f["name"].strip(), f["row"]))
    return findings


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def _append_shared_string(shared_xml, text):
    """(new sharedStrings xml, index) for `text`, reusing an existing entry
    when the exact string is already in the table. Always writes
    xml:space="preserve" -- a real name in this workbook can end in a
    space (e.g. "REYES, TOMAS "), and without that attribute Excel
    strips it on the next save, which is precisely the trailing-space
    drift the mapping check keeps catching."""
    existing = []
    for si in re.findall(r"<si>(.*?)</si>", shared_xml, re.S):
        existing.append(_unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
    if text in existing:
        return shared_xml, existing.index(text)

    index = len(existing)
    si = f'<si><t xml:space="preserve">{_escape(text)}</t></si>'
    if "</sst>" not in shared_xml:
        raise TipSheetWriteError("This workbook's shared string table is in a format this app can't extend.")
    shared_xml = shared_xml.replace("</sst>", si + "</sst>", 1)

    # `count` is total cell references to strings, `uniqueCount` the size
    # of the table. One new cell pointing at one new string bumps both;
    # Excel repairs the file (with a warning dialog) if they're wrong.
    def bump(attr, xml, by):
        m = re.search(r'\b%s="(\d+)"' % attr, xml)
        return xml if not m else xml[:m.start()] + f'{attr}="{int(m.group(1)) + by}"' + xml[m.end():]

    shared_xml = bump("uniqueCount", shared_xml, 1)
    shared_xml = bump("count", shared_xml, 1)
    return shared_xml, index


def _bump_shared_count(shared_xml, by):
    m = re.search(r'\bcount="(\d+)"', shared_xml)
    if not m:
        return shared_xml
    return shared_xml[:m.start()] + f'count="{int(m.group(1)) + by}"' + shared_xml[m.end():]


def _write_cell(sheet_xml, ref, inner, as_shared_string):
    """Fills one existing, empty cell in place, keeping its style (`s`)
    attribute exactly as the workbook already had it -- that's what makes
    the new row look and behave identically to the rows around it. Only
    ever touches a cell that is currently empty."""
    m = _cell_match(sheet_xml, ref)
    if not m:
        raise TipSheetWriteError(f"Cell {ref} doesn't exist in this workbook's WEEKLY TOTALS sheet.")
    attrs, body = m.group("attrs"), m.group("body") or ""
    if "<v>" in body:
        raise TipSheetWriteError(f"Cell {ref} already has a value -- refusing to overwrite it.")
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    if as_shared_string:
        attrs += ' t="s"'
    return sheet_xml[:m.start()] + f'<c r="{ref}"{attrs}><v>{inner}</v></c>' + sheet_xml[m.end():]


def plan_additions(xlsx_path, additions):
    """Validates a set of {role, eid, tip_sheet_name} additions against
    the real workbook WITHOUT writing anything, and works out which slot
    row each one would land in.

    Returns (placements, errors). Every refusal is explicit -- a name or
    EID already in that section, a role with no free slots left, a
    section whose slots are structurally broken. Nothing is ever matched
    approximately: an existing name counts as "already there" on an exact
    match, and separately as a near-duplicate warning when it differs
    only by spacing or case, which is a human's call to make, not this
    function's.
    """
    sections = read_workbook_sections(xlsx_path)
    placements, errors = [], []

    def norm(s):
        return re.sub(r"\s+", " ", s.strip()).lower()

    # Slots claimed earlier in this same batch, so two people added at
    # once can't both be assigned the same row.
    taken = {role: set() for role in sections}

    for add in additions:
        role = add.get("role")
        name = add.get("tip_sheet_name", "")
        eid = str(add.get("eid", "")).strip()
        label = name.strip() or "(blank name)"

        if role not in sections:
            errors.append(f"{label}: {role!r} isn't one of this workbook's roles "
                          f"({', '.join(ROLE_ORDER)}).")
            continue
        if not name.strip():
            errors.append(f"{role}: no Tip Sheet name to write.")
            continue
        if not eid:
            errors.append(f"{label}: an EID is required -- it goes in column A next to the name.")
            continue

        section = sections[role]
        clash = next((o for o in section["occupied"] if o["name"] == name), None)
        if clash:
            errors.append(f"{label} is already in the workbook's {role} section at row "
                          f"{clash['row']} (EID {clash['eid']}). Nothing to add.")
            continue
        near = next((o for o in section["occupied"] if norm(o["name"]) == norm(name)), None)
        if near:
            errors.append(
                f"{label}: the {role} section already has {near['name']!r} at row {near['row']}, "
                "which differs only in spacing or capitalization. Adding both would give this "
                "person two dropdown entries that look identical but only one of which the "
                "formulas add up. Use the existing spelling, or fix the workbook first."
            )
            continue
        eid_clash = next((o for o in section["occupied"]
                          if o["eid"] and o["eid"].lstrip("0") == eid.lstrip("0")), None)
        if eid_clash:
            errors.append(f"{label}: EID {eid} is already in the {role} section at row "
                          f"{eid_clash['row']} as {eid_clash['name']!r}.")
            continue

        free = [r for r in section["free_rows"] if r not in taken[role]]
        if not free:
            first, last = section["range"]
            detail = ""
            if section["broken_rows"]:
                detail = (f" ({len(section['broken_rows'])} empty rows in that range were skipped "
                          "because their formulas point at the wrong row -- see the workbook "
                          "structure warning above)")
            errors.append(
                f"{label}: the {role} section (rows {first}-{last}) has no free rows left{detail}. "
                "Someone needs to extend that section in Excel -- and the dropdown range with it -- "
                "before anyone else can be added to it."
            )
            continue

        row = free[0]
        taken[role].add(row)
        placements.append({"role": role, "row": row, "eid": eid, "tip_sheet_name": name})

    return placements, errors


def add_employees_to_workbook(xlsx_path, additions):
    """Writes each addition into the workbook and returns the new .xlsx as
    bytes, alongside the placements actually made.

    Raises TipSheetWriteError if anything at all is wrong -- this is
    all-or-nothing on purpose. A partially written workbook handed back
    to a manager is worse than no workbook: they'd have no way to tell
    which half of it landed.
    """
    placements, errors = plan_additions(xlsx_path, additions)
    if errors:
        raise TipSheetWriteError(" ".join(errors))
    if not placements:
        raise TipSheetWriteError("Nothing to add.")

    with zipfile.ZipFile(xlsx_path) as zf:
        part = _sheet_part_name(zf, WEEKLY_TOTALS_SHEET)
        sheet_xml = zf.read(part).decode("utf-8")
        try:
            shared_xml = zf.read("xl/sharedStrings.xml").decode("utf-8")
        except KeyError:
            raise TipSheetWriteError(
                "This workbook has no shared string table, which the real Tip Sheet always has -- "
                "confirm this is the right file."
            )
        entries = [(i, zf.read(i.filename)) for i in zf.infolist()]

    for p in placements:
        shared_xml, idx = _append_shared_string(shared_xml, p["tip_sheet_name"])
        sheet_xml = _write_cell(sheet_xml, f"B{p['row']}", str(idx), as_shared_string=True)

        # Most EIDs in the sheet are numeric cells; a couple are text
        # (leading zeros would otherwise be lost). Match whichever the
        # value itself needs rather than forcing one type.
        eid = p["eid"]
        if eid.isdigit() and not eid.startswith("0"):
            sheet_xml = _write_cell(sheet_xml, f"A{p['row']}", eid, as_shared_string=False)
        else:
            shared_xml, eid_idx = _append_shared_string(shared_xml, eid)
            sheet_xml = _write_cell(sheet_xml, f"A{p['row']}", str(eid_idx), as_shared_string=True)

    buf = BytesIO()
    replacements = {part: sheet_xml.encode("utf-8"), "xl/sharedStrings.xml": shared_xml.encode("utf-8")}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            # Every other part -- styles, the 14 day tabs with their
            # dropdown definitions, printer settings, calcChain -- is
            # copied straight through. Nothing this app doesn't
            # understand gets rewritten, which is the whole reason for
            # not going through openpyxl.
            out.writestr(info, replacements.get(info.filename, data))
    return buf.getvalue(), placements


def suggested_tip_sheet_name(full_name, existing_names=()):
    """A first draft of the "LAST, FIRST" text for a typed-in name,
    following the convention the workbook already uses.

    Only ever a suggestion shown back for confirmation -- the manager can
    edit it, and whatever they confirm is what gets written into BOTH the
    workbook and the mapping, so the two match regardless of whether this
    guess was right. That's deliberate: the app never has to be correct
    about name formatting, it only has to be consistent.

    `existing_names` (the section's current names) is used only to match
    the local convention when there's a clear one -- e.g. a section that
    writes "LAST, FIRST" without a space after the comma.
    """
    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0].upper()
    first, last = parts[0], " ".join(parts[1:])
    sep = ", "
    with_space = sum(1 for n in existing_names if re.search(r",\s", n))
    without = sum(1 for n in existing_names if re.search(r",\S", n))
    if without > with_space:
        sep = ","
    return f"{last.upper()}{sep}{first.upper()}"


def repair_formula_rows(xlsx_path):
    """Re-points every wrong-row SUMIF found by audit_formula_rows() at
    its own row, and returns the corrected .xlsx as bytes plus a record
    of what changed.

    Deliberately narrow: it only rewrites the `'WEEKLY TOTALS'!B<n>`
    criteria reference inside rows that are (a) inside a role's own
    discovered dropdown range and (b) already flagged as pointing
    somewhere other than themselves. The lookup ranges and sum ranges in
    those same formulas are left exactly as they are -- they were never
    the thing that was wrong, and a broader rewrite would be this app
    inventing what the workbook's arithmetic ought to be.

    The equivalent by hand is copying a correct row over the broken ones
    and letting Excel's relative references re-point themselves; this
    just does it without the chance of pasting over a row that had a
    name in it.

    Also sets fullCalcOnLoad so Excel recomputes the whole book when it
    opens. The cached results sitting in the file were computed from the
    WRONG formulas, and without this Excel is entitled to keep showing
    them until something forces a recalculation -- which would mean a
    repaired workbook still displaying the bad numbers.
    """
    findings = audit_formula_rows(xlsx_path)
    if not findings:
        raise TipSheetWriteError("Nothing to repair -- every row already totals its own name.")

    with zipfile.ZipFile(xlsx_path) as zf:
        part = _sheet_part_name(zf, WEEKLY_TOTALS_SHEET)
        sheet_xml = zf.read(part).decode("utf-8")
        workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")
        entries = [(i, zf.read(i.filename)) for i in zf.infolist()]

    rows = _rows_by_number(sheet_xml)
    repaired = []
    for f in findings:
        r = f["row"]
        old_row_xml = rows.get(r)
        if old_row_xml is None:
            continue
        new_row_xml = re.sub(
            r"('%s'!B)(\d+)" % re.escape(WEEKLY_TOTALS_SHEET),
            lambda m: f"{m.group(1)}{r}",
            old_row_xml,
        )
        if new_row_xml == old_row_xml:
            continue
        # Replacing the row's exact XML rather than editing by offset:
        # each row string is unique (it carries its own r="N"), and this
        # can't drift as earlier replacements change the document length.
        sheet_xml = sheet_xml.replace(old_row_xml, new_row_xml, 1)
        rows[r] = new_row_xml
        repaired.append({"role": f["role"], "row": r, "name": f["name"],
                         "was": f["refs"], "now": [r]})

    if not repaired:
        raise TipSheetWriteError("Found rows to repair but couldn't rewrite them -- this workbook's "
                                 "WEEKLY TOTALS sheet isn't laid out the way this app expects.")

    if "fullCalcOnLoad" not in workbook_xml:
        if "<calcPr" in workbook_xml:
            workbook_xml = re.sub(r"<calcPr\b([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>',
                                  workbook_xml, count=1)
        else:
            workbook_xml = workbook_xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>', 1)

    buf = BytesIO()
    replacements = {part: sheet_xml.encode("utf-8"), "xl/workbook.xml": workbook_xml.encode("utf-8")}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            out.writestr(info, replacements.get(info.filename, data))
    return buf.getvalue(), repaired
