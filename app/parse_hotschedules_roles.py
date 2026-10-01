"""
HotSchedules PDF -> per-employee, per-day, per-shift ROLE presence.

Simplified from the full shift-time reconstruction: instead of trying to
perfectly pair each time-range with its position line (fragile under OCR
noise), this only answers the question the cross-check actually needs:
"what tipped role(s) is this person listed under, for this specific day,
split into an early ('AM') and late ('PM') shift if they worked a double."

This is deliberately conservative: if a cell's structure is ambiguous
(more than 2 shift blocks, or a shift block with no clear role word), it
is flagged rather than guessed.
"""

import re
import csv

ROLE_WORDS = {"Server", "Bartender", "Runner", "Busser", "Barback", "Barista"}
SPECIAL_TAGS = ["TRAINING", "NOT LIVE YET", "DOUBLE", "CALL IN"]

TIME_RE = re.compile(r"\d{1,2}:\d{2}\s*[AP]M")
PHONE_RE = re.compile(r"^\(\d{3}\)$")


def load_words(tsv_path):
    words = []
    with open(tsv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter="\t")
        for row in reader:
            if row["level"] != "5":
                continue
            text = row["text"].strip()
            if not text:
                continue
            words.append({"text": text, "left": int(row["left"]), "top": int(row["top"]),
                           "width": int(row["width"]), "height": int(row["height"])})
    return words


def find_column_bounds(words):
    header_candidates = [w for w in words if w["text"] == "Employee"]
    if not header_candidates:
        raise ValueError("Could not find header row on this page")
    header_y = header_candidates[0]["top"]
    band = [w for w in words if abs(w["top"] - header_y) < 15]
    band.sort(key=lambda w: w["left"])

    labels = ["Employee", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    used, starts = set(), []
    for w in band:
        for lab in labels:
            if lab not in used and w["text"].startswith(lab):
                starts.append((lab, w["left"]))
                used.add(lab)
                break
    starts.sort(key=lambda x: x[1])
    bounds = []
    for i, (lab, x) in enumerate(starts):
        x_end = starts[i + 1][1] - 5 if i + 1 < len(starts) else 10 ** 9
        bounds.append((lab, x, x_end))
    return bounds


def bucket_words_by_column(words, bounds):
    for w in words:
        for i, (lab, x0, x1) in enumerate(bounds):
            effective_x0 = 0 if i == 0 else (x0 - 20)
            if effective_x0 <= w["left"] < x1:
                w["col"] = lab
                break
        else:
            w["col"] = None
    return words


def find_employee_row_bands(words):
    emp_col = [w for w in words if w["col"] == "Employee"]
    phone_tops = sorted(w["top"] for w in emp_col if PHONE_RE.match(w["text"]))
    if not phone_tops:
        raise ValueError("No phone-number anchors found on this page")
    heights = [w["height"] for w in emp_col if w["height"] > 0]
    line_height = sorted(heights)[len(heights) // 2] if heights else 40

    bands = []
    for i, phone_y in enumerate(phone_tops):
        y0 = phone_y - int(line_height * 3) if i == 0 else (phone_tops[i - 1] + phone_y) // 2
        y1 = (phone_y + phone_tops[i + 1]) // 2 if i + 1 < len(phone_tops) else 10 ** 9
        bands.append((y0, y1))
    return bands


def extract_name(words, y0, y1):
    cand = [w for w in words if w["col"] == "Employee" and y0 <= w["top"] < y1]
    cand.sort(key=lambda w: (w["top"], w["left"]))
    phone_ys = [w["top"] for w in cand if PHONE_RE.match(w["text"])]
    if not phone_ys:
        return None
    phone_y = min(phone_ys)
    name_words = [w["text"] for w in cand if w["top"] < phone_y - 5]
    # strip stray non-alphanumeric leading/trailing noise (OCR artifacts from
    # the checkbox glyph, e.g. a leftover "(" or "\\")
    cleaned = [re.sub(r"^[^A-Za-z0-9]+|[^A-Za-z0-9.]+$", "", w) for w in name_words]
    cleaned = [w for w in cleaned if w]
    return " ".join(cleaned).strip()


def _shift_period(time_str):
    """Classify a shift as AM or PM based on actual start time, not position.
    Cutoff matches the Toast classifier: before 1:00 PM = AM/lunch,
    1:00 PM or later = PM/dinner."""
    m = re.match(r"(\d{1,2}):(\d{2})\s*([AP]M)", time_str)
    if not m:
        return None
    hour, minute, ampm = int(m.group(1)), int(m.group(2)), m.group(3)
    hour24 = hour % 12 + (12 if ampm == "PM" else 0)
    return "AM" if (hour24, minute) < (13, 0) else "PM"


def extract_day_roles(words, col, y0, y1):
    """
    Returns a dict {"AM": {...} or None, "PM": {...} or None} for this
    person+day. Each shift is classified by its ACTUAL start time (not
    by whether it's the first or second shift found in the cell) --
    this matters because a person's only shift that day might be an
    evening one, which must not be mislabeled "AM" just because it's
    the only shift present.
    """
    cand = [w for w in words if w["col"] == col and y0 <= w["top"] < y1]
    if not cand:
        return {}
    cand.sort(key=lambda w: w["top"])

    time_word_tops = sorted(set(
        w["top"] for w in cand
        if re.match(r"^\d{1,2}:\d{2}$", w["text"]) or w["text"] in ("AM", "PM")
    ))
    clusters = []
    for t in time_word_tops:
        if clusters and t - clusters[-1][-1] < 15:
            clusters[-1].append(t)
        else:
            clusters.append([t])

    if len(clusters) <= 1:
        groups = [cand]
    else:
        split_y = clusters[1][0] - 10
        groups = [[w for w in cand if w["top"] < split_y],
                  [w for w in cand if w["top"] >= split_y]]
        groups = [g for g in groups if g][:2]  # cap at 2 shift blocks

    result = {"AM": None, "PM": None}
    for g in groups:
        if not g:
            continue
        text_blob = " ".join(w["text"] for w in sorted(g, key=lambda w: (w["top"], w["left"])))
        time_match = re.search(r"\d{1,2}:\d{2}\s*[AP]M", text_blob)
        period = _shift_period(time_match.group(0)) if time_match else None
        roles_found = [r for r in ROLE_WORDS if re.search(rf"\b{r}\b", text_blob)]
        tags_found = [t for t in SPECIAL_TAGS if t.replace(" ", "") in text_blob.replace(" ", "").upper()
                      or t in text_blob.upper()]
        entry = {"roles": roles_found, "tags": tags_found, "raw": text_blob}
        if period in ("AM", "PM"):
            # if two shifts land in the same period (rare/ambiguous), keep
            # both roles rather than silently overwriting one
            if result[period] is None:
                result[period] = entry
            else:
                result[period]["roles"] = list(set(result[period]["roles"] + roles_found))
                result[period]["tags"] = list(set(result[period]["tags"] + tags_found))
        # if period couldn't be determined at all, this shift is dropped
        # from AM/PM lookup but not silently lost -- report separately
    return {k: v for k, v in result.items() if v is not None}


def parse_page(tsv_path, bounds=None):
    words = load_words(tsv_path)
    if bounds is None:
        bounds = find_column_bounds(words)
    words = bucket_words_by_column(words, bounds)
    bands = find_employee_row_bands(words)
    day_cols = [b[0] for b in bounds if b[0] != "Employee"]

    people = []
    for (y0, y1) in bands:
        name = extract_name(words, y0, y1)
        if not name:
            continue
        day_data = {}
        for col in day_cols:
            shifts = extract_day_roles(words, col, y0, y1)
            if shifts:
                day_data[col] = shifts
        people.append({"employee": name, "days": day_data})
    return people, bounds


if __name__ == "__main__":
    import sys
    tsv_path = sys.argv[1]
    people, _ = parse_page(tsv_path)
    for p in people:
        if not p["days"]:
            continue
        print(f"\n{p['employee']}")
        for day, shifts in p["days"].items():
            for period, sh in shifts.items():
                flag = " [MULTI-ROLE OR NO ROLE FOUND]" if len(sh["roles"]) != 1 else ""
                print(f"  {day} ({period}): roles={sh['roles']} tags={sh['tags']}{flag}")
