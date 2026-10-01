"""
AVRA Nightly Check -- web app.

Three separate pages, each with only the uploads it needs:
- Daily Close ("/"): Toast + HotSchedules + ADP -> the ready-to-paste Tip
  Sheet, plus the manager-reported cut confirmation.
- Daily Summary ("/daily-summary"): two Toast Sales Summary pulls (lunch
  cutoff + final/closing) -> the lunch/dinner sales report layout. Upload
  and raw-data extraction are real; most computed cells are still
  placeholders pending cell-by-cell mapping against the real workbook.
- Punch Report Check ("/punch-report-check"): ADP + the previous day's
  Tip Sheet workbook, independent of the Daily Close -- makes sure
  everyone clocked in correctly, is on the Tip Sheet if they worked, and
  is off it (or in the right department) if they didn't, any day.

No files are kept after the request finishes (uploads go to a per-request
temp directory that's cleaned up immediately after).
"""

import base64
import binascii
import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import mapping_store
from .add_employee_flow import analyze, build_entry, preview_new_workbook_name
from .logic import (
    run_full_check, add_alias_to_employee, add_employee_to_mapping, set_employee_statuses,
    set_tip_sheet_name, update_employee, EmployeeMapping,
)
from .parse_toast_sales_summary import REVENUE_CENTERS_IN_ORDER, parse_sales_summary_zip
from .tip_sheet_writer import add_employees_to_workbook, repair_formula_rows, TipSheetWriteError

DAILY_SUMMARY_OPEN_ITEMS = [
    "Comps: which Toast discount-reason lines count (Manager Comp - Check, Owners Comp, "
    "Open $ Check, Manager Comp - Item, Compliments, Birthday, Did Not Like) and which don't.",
    "Manager Meals: confirm it's just the \"Manager Meal\" line in Menu Item Discounts.csv.",
    "Voids: confirm Void summary.csv is the right source, and whether it needs the same "
    "lunch/final-pull subtraction as revenue center sales.",
    "Delivery: no confirmed Toast source found yet -- Dining options summary.csv is a "
    "candidate but hasn't been checked against a day that actually had delivery orders.",
    "Covers per revenue center: not present anywhere in the Sales Summary export -- River "
    "confirmed this comes from Resy instead, upload TBD. Shell only until that's wired in.",
    "Check avg. by section: blocked until Resy/Covers is wired in; also confirm the "
    "\"Dining Room\" bucket should keep its odd formula (Total minus Events minus Bar minus "
    "Delivery) rather than using the Dining Room revenue-center row directly.",
    "Daily Net Sales Lunch/Dinner totals: confirm these should be summed from the same "
    "per-revenue-center pull-subtraction rather than pulled from a different report field.",
    "Week to Date: deferred -- needs either multi-day upload support or real persistence.",
    "\"No Revenue Center\" stray-items line: decide whether/how to fold it into totals or "
    "just flag it when non-zero.",
    "Whether to cross-check the two-pull subtraction against Toast's own built-in "
    "Service Daypart summary.csv Lunch/Dinner split as a sanity check, or ignore that report.",
]

BASE_DIR = Path(__file__).resolve().parent.parent
# Overridable so a deployment can point the working copy at a mounted
# disk; the durable copy lives in GitHub either way (see mapping_store).
MAPPING_PATH = Path(os.environ.get("AVRA_MAPPING_PATH") or (BASE_DIR / "employee_mapping.json"))

app = FastAPI(title="Nightly Close (demo)")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
# Demo-only: the fictional sample exports the "Load sample files" button fetches.
app.mount("/sample", StaticFiles(directory=str(BASE_DIR / "sample_data")), name="sample")


def _static_asset_version():
    """Short hash of every static/*.css and *.js file's contents, used as a
    cache-busting query string (?v=...) on every <link>/<script> tag. Without
    this, browsers that already cached style.css from before a redesign just
    keep rendering the old CSS after a deploy -- no visible error, the page
    just silently looks stale (confirmed happening for real: River saw the
    pre-redesign monospace/orange theme days after the Aegean redesign
    shipped, purely from browser caching, not a deploy problem)."""
    h = hashlib.sha256()
    for path in sorted((BASE_DIR / "static").glob("*.css")) + sorted((BASE_DIR / "static").glob("*.js")):
        h.update(path.read_bytes())
    return h.hexdigest()[:10]


templates.env.globals["asset_version"] = _static_asset_version()


async def _save_upload(upload, tmp_dir, filename):
    if upload and upload.filename:
        path = tmp_dir / filename
        path.write_bytes(await upload.read())
        return path
    return None


def _json_for_script(data):
    # Safe to inline inside a <script> tag -- guards against a name
    # containing "</script>" from prematurely closing the tag.
    return json.dumps(data).replace("</", "<\\/")


@app.get("/", response_class=HTMLResponse)
async def daily_close_form(request: Request):
    # Keyword-argument form (request=, name=, context=) works across both
    # older and newer Starlette/FastAPI versions -- the older positional
    # "(name, {\"request\": request})" convention was deprecated in newer
    # Starlette releases, so this avoids a version-pinning trap.
    mapping = EmployeeMapping(str(MAPPING_PATH))
    return templates.TemplateResponse(
        request=request,
        name="daily_close_form.html",
        context={"active_page": "daily_close", "staff_list_json": _json_for_script(mapping.active_employees())},
    )


@app.post("/run-check", response_class=HTMLResponse)
async def run_check(
    request: Request,
    toast_csv: Optional[UploadFile] = File(None),
    hotschedules_daily: Optional[UploadFile] = File(None),
    roster_date: Optional[str] = Form(None),
    shift_filter: Optional[str] = Form(None),
    cut_employees_json: Optional[str] = Form(None),
    role_swap_json: Optional[str] = Form(None),
):
    tmp_dir = Path(tempfile.mkdtemp(prefix="avra_check_"))
    try:
        toast_path = await _save_upload(toast_csv, tmp_dir, "toast.csv")
        hotsched_daily_path = await _save_upload(hotschedules_daily, tmp_dir, "hotschedules_daily.xls")
        ocr_work_dir = tmp_dir / "ocr"

        # Both come from the staff-picker UI as JSON arrays, not free text
        # -- a malformed/tampered value (never expected from the real
        # form) falls back to "nothing selected" rather than a 500.
        try:
            cut_picks = json.loads(cut_employees_json) if cut_employees_json else []
            cut_canonicals = [p.get("canonical") for p in cut_picks if isinstance(p, dict) and p.get("canonical")]
        except (json.JSONDecodeError, TypeError):
            cut_canonicals = []
        try:
            role_swaps = json.loads(role_swap_json) if role_swap_json else []
        except (json.JSONDecodeError, TypeError):
            role_swaps = []

        error = None
        result = {}
        try:
            result = run_full_check(
                toast_csv_path=str(toast_path) if toast_path else None,
                tip_sheet_path=None,
                hotsched_pdf_path=None,
                mapping_path=str(MAPPING_PATH),
                work_dir=str(ocr_work_dir),
                hotsched_daily_path=str(hotsched_daily_path) if hotsched_daily_path else None,
                roster_date=roster_date or None,
                shift_filter=shift_filter,
                adp_punches_path=None,
                cut_canonicals=cut_canonicals,
                role_swaps=role_swaps,
            )
        except Exception as e:
            error = f"{type(e).__name__}: {e}"

        if result.get("server_review") is not None:
            result["server_rows_json"] = _json_for_script([
                {
                    "id": i,
                    "shift": r["shift"],
                    "name": r["tip_sheet_name"],
                    "cash": r["cash"],
                    "credit": r["credit"],
                    "largeGratuity": r["large_gratuity"],
                    "excluded": r["excluded_from_pool"],
                    "exclusionReason": r.get("exclusion_reason"),
                    "cashSales": r.get("cash_sales", 0),
                    "cashSalesNumber": r.get("cash_sales_number"),
                }
                for i, r in enumerate(result["server_review"])
            ])
        if result.get("bartender_pools"):
            result["bartender_pools_json"] = _json_for_script(result["bartender_pools"])

        if result.get("support_staff_review") is not None:
            for i, r in enumerate(result["support_staff_review"]):
                r["id"] = i
            result["support_staff_rows_json"] = _json_for_script([
                {
                    "id": r["id"],
                    "shift": r["shift"],
                    "role": r["role"],
                    "name": r["employee"],
                    "hasAdpPunch": r["has_adp_punch"],
                    "isOnCall": r["is_on_call"],
                    "isHouseShift": r.get("is_house_shift", False),
                    "isTrainee": r.get("is_trainee", False),
                    "isTerminated": r.get("is_terminated", False),
                    "cutByManager": r.get("cut_by_manager", False),
                }
                for r in result["support_staff_review"]
            ])

        return templates.TemplateResponse(
            request=request,
            name="daily_close_results.html",
            context={"result": result, "error": error, "active_page": "daily_close"},
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.get("/daily-summary", response_class=HTMLResponse)
async def daily_summary_form(request: Request):
    return templates.TemplateResponse(request=request, name="daily_summary_form.html", context={"active_page": "daily_summary"})


@app.post("/daily-summary", response_class=HTMLResponse)
async def daily_summary_run(
    request: Request,
    toast_sales_lunch: Optional[UploadFile] = File(None),
    toast_sales_final: Optional[UploadFile] = File(None),
    roster_date: Optional[str] = Form(None),
):
    tmp_dir = Path(tempfile.mkdtemp(prefix="avra_dailysummary_"))
    try:
        lunch_zip_path = await _save_upload(toast_sales_lunch, tmp_dir, "sales_lunch.zip")
        final_zip_path = await _save_upload(toast_sales_final, tmp_dir, "sales_final.zip")

        error = None
        result = {}
        try:
            if lunch_zip_path:
                result["raw_lunch"] = parse_sales_summary_zip(str(lunch_zip_path))
                result["lunch_layout"] = True
            if final_zip_path:
                result["raw_final"] = parse_sales_summary_zip(str(final_zip_path))
                result["dinner_layout"] = True
            if lunch_zip_path or final_zip_path:
                result["revenue_centers_in_order"] = REVENUE_CENTERS_IN_ORDER
                result["open_items"] = DAILY_SUMMARY_OPEN_ITEMS
                result["roster_date"] = roster_date or None
        except Exception as e:
            error = f"{type(e).__name__}: {e}"

        return templates.TemplateResponse(
            request=request,
            name="daily_summary_results.html",
            context={"result": result, "error": error, "active_page": "daily_summary"},
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.get("/punch-report-check", response_class=HTMLResponse)
async def punch_check_form(request: Request):
    return templates.TemplateResponse(request=request, name="punch_check_form.html", context={"active_page": "punch_check"})


@app.post("/punch-report-check", response_class=HTMLResponse)
async def punch_check_run(
    request: Request,
    adp_punches: Optional[UploadFile] = File(None),
    tip_sheet_daily: Optional[UploadFile] = File(None),
    hotschedules_daily: Optional[UploadFile] = File(None),
    roster_date: Optional[str] = Form(None),
    shift_filter: Optional[str] = Form(None),
):
    tmp_dir = Path(tempfile.mkdtemp(prefix="avra_punchcheck_"))
    try:
        adp_punches_path = await _save_upload(adp_punches, tmp_dir, "adp_punches.csv")
        tip_sheet_daily_path = await _save_upload(tip_sheet_daily, tmp_dir, "tip_sheet_daily.xlsx")
        hotsched_daily_path = await _save_upload(hotschedules_daily, tmp_dir, "hotschedules_daily.xls")
        ocr_work_dir = tmp_dir / "ocr"

        error = None
        result = {}
        try:
            result = run_full_check(
                toast_csv_path=None,
                tip_sheet_path=None,
                hotsched_pdf_path=None,
                mapping_path=str(MAPPING_PATH),
                work_dir=str(ocr_work_dir),
                hotsched_daily_path=str(hotsched_daily_path) if hotsched_daily_path else None,
                roster_date=roster_date or None,
                shift_filter=shift_filter,
                adp_punches_path=str(adp_punches_path) if adp_punches_path else None,
                tip_sheet_daily_path=str(tip_sheet_daily_path) if tip_sheet_daily_path else None,
            )
        except Exception as e:
            error = f"{type(e).__name__}: {e}"

        return templates.TemplateResponse(
            request=request,
            name="punch_check_results.html",
            context={"result": result, "error": error, "active_page": "punch_check"},
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.get("/api/employees")
async def api_employees():
    # Fetched client-side by the Add Employee page on every load AND on
    # bfcache restore (browser back/forward showing a cached snapshot
    # instead of re-requesting the page) -- a picker built only from
    # server-rendered JSON at initial render can go stale if someone adds
    # an employee, navigates away, then hits "back" instead of reloading;
    # a real case of this is what "someone I just added doesn't show up
    # in the terminate picker" traced back to.
    mapping = EmployeeMapping(str(MAPPING_PATH))
    return JSONResponse(mapping.entries)

# ---------------------------------------------------------------------------
# Add Employee -- the guided, file-driven flow
# ---------------------------------------------------------------------------
#
# The page works from the two real files rather than from memory: the
# manager uploads the master Tip Sheet workbook and a HotSchedules
# roster, and every value that has to match one of those files exactly
# is then either picked from it or written into it by this app. Nothing
# that has to be byte-exact is ever typed.
#
# Both uploads are carried between steps as base64 in a hidden form
# field rather than parked in a temp directory or a session. That keeps
# the app stateless per request the way the rest of it is, and -- more
# usefully here -- means the workbook being edited is unambiguously the
# one the manager is looking at, even across a redeploy mid-edit. Each
# add hands back the UPDATED workbook bytes, so adding three people in a
# row produces one file containing all three rather than three files
# with one apiece.


def _b64_or_none(raw):
    if not raw:
        return None
    try:
        return base64.b64decode(raw, validate=True)
    except (binascii.Error, ValueError):
        return None


async def _upload_bytes(upload):
    if upload and upload.filename:
        data = await upload.read()
        if data:
            return data
    return None


def _write_temp(tmp_dir, name, data):
    if data is None:
        return None
    path = tmp_dir / name
    path.write_bytes(data)
    return str(path)


def _add_employee_context(workbook_bytes, roster_bytes, workbook_name, roster_name,
                          outcome=None, download=None):
    """Renders the Add Employee page against whichever files are in hand.

    The reconciliation is recomputed from the CURRENT bytes on every
    render, including straight after a change, so what the page shows is
    always the state of the files as they now are -- a person just added
    drops off the "missing" list in the same response that added them,
    instead of lingering until someone reloads.
    """
    tmp_dir = Path(tempfile.mkdtemp(prefix="avra_addemp_"))
    try:
        review = analyze(
            str(MAPPING_PATH),
            _write_temp(tmp_dir, "tip_sheet.xlsx", workbook_bytes),
            _write_temp(tmp_dir, "roster", roster_bytes),
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    mapping = EmployeeMapping(str(MAPPING_PATH))
    return {
        "active_page": "add_employee",
        "review": review,
        "outcome": outcome,
        "download": download,
        "workbook_b64": base64.b64encode(workbook_bytes).decode("ascii") if workbook_bytes else "",
        "roster_b64": base64.b64encode(roster_bytes).decode("ascii") if roster_bytes else "",
        "workbook_name": workbook_name or "",
        "roster_name": roster_name or "",
        "sync_status": mapping_store.status(),
        "all_employees": mapping.active_employees(),
        "all_employees_json": _json_for_script(mapping.entries),
        # Kept for the two long-standing forms further down the page.
        "result": None,
        "status_result": None,
        "edit_result": None,
    }


@app.get("/add-employee", response_class=HTMLResponse)
async def add_employee_form(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="add_employee_form.html",
        context=_add_employee_context(None, None, None, None),
    )


@app.post("/add-employee/review", response_class=HTMLResponse)
async def add_employee_review(
    request: Request,
    tip_sheet_workbook: Optional[UploadFile] = File(None),
    hotschedules_roster: Optional[UploadFile] = File(None),
    workbook_b64: str = Form(""),
    roster_b64: str = Form(""),
    workbook_name: str = Form(""),
    roster_name: str = Form(""),
):
    # A fresh upload wins; otherwise whatever was already loaded is kept,
    # so re-uploading only the roster doesn't silently drop the workbook.
    new_workbook = await _upload_bytes(tip_sheet_workbook)
    new_roster = await _upload_bytes(hotschedules_roster)
    workbook = new_workbook or _b64_or_none(workbook_b64)
    roster = new_roster or _b64_or_none(roster_b64)
    return templates.TemplateResponse(
        request=request,
        name="add_employee_form.html",
        context=_add_employee_context(
            workbook, roster,
            (tip_sheet_workbook.filename if new_workbook else workbook_name),
            (hotschedules_roster.filename if new_roster else roster_name),
        ),
    )


@app.post("/add-employee/apply", response_class=HTMLResponse)
async def add_employee_apply(
    request: Request,
    action: str = Form(""),
    workbook_b64: str = Form(""),
    roster_b64: str = Form(""),
    workbook_name: str = Form(""),
    roster_name: str = Form(""),
    canonical: str = Form(""),
    eid: str = Form(""),
    tip_sheet: str = Form(""),
    role: str = Form(""),
    alias: str = Form(""),
    alias_system: str = Form("hotschedules"),
    hotschedules_names: str = Form(""),
    toast_names: str = Form(""),
    cash_sales_number: str = Form(""),
):
    workbook = _b64_or_none(workbook_b64)
    roster = _b64_or_none(roster_b64)
    outcome = None
    download = None

    def lines(text):
        return [n.strip() for n in text.splitlines() if n.strip()]

    if action == "map_existing":
        # The person is already a row in WEEKLY TOTALS. `tip_sheet` here
        # is that row's cell text, posted back from the analysis this
        # app read out of the workbook -- not something retyped -- so
        # the mapping ends up holding the same bytes Excel does.
        entry = build_entry(canonical, eid, tip_sheet, lines(hotschedules_names),
                            lines(toast_names), cash_sales_number)
        result = add_employee_to_mapping(str(MAPPING_PATH), entry)
        outcome = {
            "ok": result["added"],
            "title": f"{entry['canonical']} added to the mapping" if result["added"]
                     else "Not added",
            "detail": (f"Mapped to the workbook's existing row for {entry['tip_sheet']!r}. "
                       "Nothing in the Tip Sheet changed, so there's no file to download.")
                      if result["added"] else "",
            "errors": result["errors"],
            "sync": result.get("sync"),
        }

    elif action == "add_alias":
        result = add_alias_to_employee(str(MAPPING_PATH), canonical, alias, alias_system)
        outcome = {
            "ok": result["added"],
            "title": f"{alias!r} now matches {canonical}" if result["added"] else "Not added",
            "detail": (f"Added as a {alias_system} name. Nothing in the Tip Sheet changed.")
                      if result["added"] else "",
            "errors": result["errors"],
            "sync": result.get("sync"),
        }

    elif action == "fix_drift":
        result = set_tip_sheet_name(str(MAPPING_PATH), canonical, tip_sheet)
        outcome = {
            "ok": result["updated"],
            "title": f"{canonical}'s Tip Sheet name now matches the workbook"
                     if result["updated"] else "Not changed",
            "detail": (f"{result.get('old')!r} -> {result.get('new')!r}")
                      if result["updated"] else "",
            "errors": result["errors"],
            "sync": result.get("sync"),
        }

    elif action == "create_in_workbook":
        if workbook is None:
            outcome = {"ok": False, "title": "No workbook loaded", "errors": [
                "Upload the master Tip Sheet workbook first -- this app writes the new person "
                "into it, which is what keeps the mapping and the spreadsheet identical."
            ], "sync": None, "detail": ""}
        else:
            tmp_dir = Path(tempfile.mkdtemp(prefix="avra_addemp_write_"))
            try:
                src = tmp_dir / "tip_sheet.xlsx"
                src.write_bytes(workbook)
                try:
                    new_bytes, placements = add_employees_to_workbook(
                        str(src), [{"role": role, "eid": eid.strip(), "tip_sheet_name": tip_sheet}]
                    )
                except TipSheetWriteError as e:
                    outcome = {"ok": False, "title": "Nothing was written to the workbook",
                               "errors": [str(e)], "sync": None, "detail": ""}
                else:
                    # The workbook write is what makes the two sides
                    # identical, so it happens first; the mapping entry
                    # then stores the very same string. If the mapping
                    # write fails afterwards, the manager is told not to
                    # keep the downloaded file -- an unmapped name in the
                    # sheet is recoverable, a mismatched pair is the bug
                    # this whole page exists to prevent.
                    entry = build_entry(canonical or tip_sheet, eid, tip_sheet,
                                        lines(hotschedules_names), lines(toast_names),
                                        cash_sales_number)
                    result = add_employee_to_mapping(str(MAPPING_PATH), entry)
                    p = placements[0]
                    if result["added"]:
                        workbook = new_bytes
                        download = {
                            "filename": workbook_name or "TIP_SHEET.xlsx",
                            "b64": base64.b64encode(new_bytes).decode("ascii"),
                        }
                        outcome = {
                            "ok": True,
                            "title": f"{entry['canonical']} written into the Tip Sheet and mapped",
                            "detail": (
                                f"Added as {p['tip_sheet_name']!r} (EID {p['eid']}) in the "
                                f"{p['role']} section, row {p['row']} of WEEKLY TOTALS. The "
                                "mapping stores that exact same text, so the dropdown and the "
                                "app agree by construction. Download the updated workbook below "
                                "and save it over the one you use."
                            ),
                            "errors": [], "sync": result.get("sync"),
                        }
                    else:
                        outcome = {
                            "ok": False,
                            "title": "Written to the workbook, but NOT mapped -- don't use that file",
                            "detail": (
                                f"The row was written into WEEKLY TOTALS, but the mapping refused "
                                "the entry, so the two would not agree. Discard this attempt, fix "
                                "the problem below, and add them again from the original workbook."
                            ),
                            "errors": result["errors"], "sync": result.get("sync"),
                        }
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)

    elif action == "repair_workbook":
        # Purely a workbook repair -- no mapping change, so nothing to
        # sync. The corrected file is carried forward as the loaded
        # workbook too, so a repair followed by an add produces one file
        # with both changes rather than two files to reconcile by hand.
        if workbook is None:
            outcome = {"ok": False, "title": "No workbook loaded",
                       "errors": ["Upload the Tip Sheet workbook above first."],
                       "sync": None, "detail": ""}
        else:
            tmp_dir = Path(tempfile.mkdtemp(prefix="avra_addemp_repair_"))
            try:
                src = tmp_dir / "tip_sheet.xlsx"
                src.write_bytes(workbook)
                try:
                    new_bytes, repaired = repair_formula_rows(str(src))
                except TipSheetWriteError as e:
                    outcome = {"ok": False, "title": "Nothing was repaired", "errors": [str(e)],
                               "sync": None, "detail": ""}
                else:
                    workbook = new_bytes
                    download = {
                        "filename": workbook_name or "TIP_SHEET.xlsx",
                        "b64": base64.b64encode(new_bytes).decode("ascii"),
                    }
                    named = [r for r in repaired if r["name"].strip()]
                    outcome = {
                        "ok": True,
                        "title": f"Repaired {len(repaired)} rows",
                        "detail": (
                            f"{len(named)} of them had a name in them and were totalling the wrong "
                            f"person's tips: {', '.join(r['name'].strip() for r in named)}. "
                            "Every row now totals its own name. No name, EID, dropdown or sum range "
                            "was changed. Download the corrected workbook below and save it over "
                            "the one you use."
                        ) if named else (
                            "All of them were empty slots, so no money was being mis-totalled yet "
                            "-- but the next person added to that section would have been. "
                            "Download the corrected workbook below."
                        ),
                        "errors": [], "sync": None,
                    }
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)

    else:
        outcome = {"ok": False, "title": "Unknown action", "errors": [f"{action!r}"],
                   "sync": None, "detail": ""}

    return templates.TemplateResponse(
        request=request,
        name="add_employee_form.html",
        context=_add_employee_context(workbook, roster, workbook_name, roster_name,
                                      outcome=outcome, download=download),
    )


@app.post("/add-employee/preview-name")
async def add_employee_preview_name(
    full_name: str = Form(""), role: str = Form(""), workbook_b64: str = Form("")
):
    """The "LAST, FIRST" text the app proposes writing into the workbook,
    fetched as the manager types so they confirm a real string rather
    than trusting a convention they can't see."""
    workbook = _b64_or_none(workbook_b64)
    tmp_dir = Path(tempfile.mkdtemp(prefix="avra_addemp_name_"))
    try:
        path = _write_temp(tmp_dir, "tip_sheet.xlsx", workbook)
        suggestion = preview_new_workbook_name(full_name, path, role) if path else ""
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return JSONResponse({"tip_sheet": suggestion})


@app.post("/add-employee/download")
async def add_employee_download(workbook_b64: str = Form(""), filename: str = Form("TIP_SHEET.xlsx")):
    data = _b64_or_none(workbook_b64)
    if data is None:
        return JSONResponse({"error": "Nothing to download."}, status_code=400)
    safe = Path(filename).name or "TIP_SHEET.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{safe}"'},
    )


@app.post("/employee-status", response_class=HTMLResponse)
async def employee_status_run(
    request: Request,
    status_updates_json: str = Form(""),
    workbook_b64: str = Form(""),
    roster_b64: str = Form(""),
    workbook_name: str = Form(""),
    roster_name: str = Form(""),
):
    try:
        updates = json.loads(status_updates_json) if status_updates_json else []
    except (json.JSONDecodeError, TypeError):
        updates = []

    result = set_employee_statuses(str(MAPPING_PATH), updates)
    context = _add_employee_context(
        _b64_or_none(workbook_b64), _b64_or_none(roster_b64), workbook_name, roster_name
    )
    context["status_result"] = result
    return templates.TemplateResponse(request=request, name="add_employee_form.html", context=context)


@app.post("/edit-employee", response_class=HTMLResponse)
async def edit_employee_run(
    request: Request,
    original_canonical: str = Form(""),
    canonical: str = Form(""),
    eid: str = Form(""),
    tip_sheet: str = Form(""),
    toast_names: str = Form(""),
    hotschedules_names: str = Form(""),
    cash_sales_number: str = Form(""),
    workbook_b64: str = Form(""),
    roster_b64: str = Form(""),
    workbook_name: str = Form(""),
    roster_name: str = Form(""),
):
    entry = build_entry(canonical, eid, tip_sheet,
                        [n for n in hotschedules_names.splitlines()],
                        [n for n in toast_names.splitlines()], cash_sales_number)
    entry.pop("status", None)  # status is the terminate/reactivate picker's job, not this form's

    result = update_employee(str(MAPPING_PATH), original_canonical, entry)
    context = _add_employee_context(
        _b64_or_none(workbook_b64), _b64_or_none(roster_b64), workbook_name, roster_name
    )
    context["edit_result"] = result
    return templates.TemplateResponse(request=request, name="add_employee_form.html", context=context)


@app.on_event("startup")
async def _startup_sync():
    """Pull the durable mapping from GitHub before serving anything.

    Render rebuilds this container's filesystem on every deploy and on
    every wake from the free tier's idle spin-down, so the
    employee_mapping.json baked into the image is only as new as the
    last git push. Without this, an employee added through the website
    yesterday is simply gone this evening -- which is exactly what kept
    happening. Failures here are recorded, not raised: a GitHub outage
    must not take the nightly close down with it.
    """
    mapping_store.pull_to_local(str(MAPPING_PATH))
@app.get("/health")
async def health():
    """Liveness check, plus enough to diagnose the mapping store from a
    browser without shell access to the container.

    Deliberately reports only whether a token is *present* -- never any
    part of its value. Repo and branch are already public in the README.
    The useful signal is what this endpoint does NOT say: a deploy still
    running the pre-sync code returns a bare {"status": "ok"} with no
    "mapping" block at all, which distinguishes "Render hasn't picked up
    the new build" from "the build is live but can't see the variable" --
    two problems with completely different fixes that look identical
    from the Add Employee page's banner.
    """
    sync = mapping_store.status()
    return {
        "status": "ok",
        "mapping": {
            "token_configured": sync["configured"],
            "env_var_read": "AVRA_GITHUB_TOKEN",
            "repo": sync["repo"],
            "branch": sync["branch"],
            "last_sync_ok": sync["ok"],
            "last_sync_detail": sync["detail"],
            "diagnosis": sync.get("diagnosis"),
            "mapping_path": str(MAPPING_PATH),
            "employees_loaded": len(EmployeeMapping(str(MAPPING_PATH)).entries),
        },
        "env": mapping_store.env_diagnostics(),
        "deployment": mapping_store.deployment_info(),
    }
