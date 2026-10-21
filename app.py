"""Marks pipeline server (Flask). Built to run on a small free instance (512 MB RAM, a fraction of a CPU).

Endpoints (all but /health need the header  X-API-Key: <one of API_KEYS>):
  GET  /health   -> {"ok": true}                      no data, for the host's health check
  POST /sheets   JSON {class, section, output?, school?, academic_year?, exam_type?, working_days?, roster:[{id,name}], subjects:[{key,label,slots:[{name,max}]}], has_attendance?}
                 -> Excel workbook, PDF (all sheets merged), or (by default) PDF (all sheets merged)
  POST /merge    multipart: workbook=<filled .xlsx>, subjects, has_attendance?, working_days?
                 -> {"ok":true,"data":{...},"incomplete":[...]}  or 422 {"ok":false,"errors":[...]}
  POST /submit   multipart: class, roster=<JSON [{id,name}]>, subjects, exam_type?, working_days?,
                 photos=<1-8 auto-routed images> or photos_<subject key>=<images>,
                 photos_Attendance=<optional images>
                 -> {"data":{...},"tabs":[...],"scans":{...},"review":[...]}  partial fillMarks() draft
                 Requires the QR-coded PDFs made by /sheets, including all four page-corner codes.

Nothing is stored: every request is handled in memory or in a temporary folder that is deleted straight away,
and request contents are never logged.

Settings (environment variables):
  API_KEYS         required: comma-separated secret keys, one per teacher. Make one with
                   python -c "import secrets; print(secrets.token_urlsafe(32))"
  ALLOW_NO_AUTH=1  local testing only: allows running without API_KEYS
  ALLOWED_ORIGINS  comma-separated origins allowed to call from a browser page (e.g. chrome-extension://<id>)
  MAX_UPLOAD_MB    largest request body, default 20
  RATE_PER_MIN     requests per minute per key, default 30
  MODEL_PATH       default model.json next to app.py
"""
import functools
import hashlib
import hmac
import io
import json
import math
import os
import re
import sys
import tempfile
import threading
import time
import zipfile
from collections import defaultdict, deque
from typing import Any

HERE = os.path.dirname(os.path.abspath(__file__))
for p in (HERE, os.path.join(HERE, "scan")):
    if p not in sys.path:
        sys.path.insert(0, p)

from flask import Flask, jsonify, request, send_file  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from werkzeug.exceptions import HTTPException  # noqa: E402

import make_sheets  # noqa: E402
import merge_marks  # noqa: E402
import reader  # noqa: E402
from config import MAX_STUDENTS, SECTION_CODES  # noqa: E402
from digits import load_model  # noqa: E402


def _env_list(name):
    return [x.strip() for x in os.environ.get(name, "").split(",") if x.strip()]


API_KEYS = _env_list("API_KEYS")
ALLOW_NO_AUTH = os.environ.get("ALLOW_NO_AUTH") == "1"
ALLOWED_ORIGINS = set(_env_list("ALLOWED_ORIGINS"))
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "20"))
RATE_PER_MIN = int(os.environ.get("RATE_PER_MIN", "30"))
IP_RATE_PER_MIN = int(os.environ.get("IP_RATE_PER_MIN", "120"))  # higher: a school can share one address
MODEL_PATH = os.environ.get("MODEL_PATH", os.path.join(HERE, "model.json"))

if not API_KEYS and not ALLOWED_ORIGINS and not ALLOW_NO_AUTH:
    raise RuntimeError("Set ALLOWED_ORIGINS (the extension ID, e.g. chrome-extension://<id>), or API_KEYS. For local testing only, set ALLOW_NO_AUTH=1.")

NET = load_model(MODEL_PATH)
MAX_PHOTOS = 8
REVIEW_CONFIDENCE = 0.90
ID_RE = re.compile(r"^[0-9A-Za-z]{1,20}$")
SECTION_RE = re.compile(r"^[0-9A-Za-z]{1,3}$")
CONTROL_RE = re.compile(r"[\x00-\x1f\x7f]")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


# ---------- helpers ----------
def _err(code, message, **extra):
    return jsonify({"error": message, **extra}), code


class BadRequest(Exception):
    pass


def _int(value: object, name: str, lo: int, hi: int, optional: bool = False) -> int | None:
    if value is None or value == "":
        if optional:
            return None
        raise BadRequest(f"{name} must be a whole number")
    if isinstance(value, bool) or not isinstance(value, (str, int)):
        raise BadRequest(f"{name} must be a whole number")
    try:
        n = int(value)
    except (TypeError, ValueError):
        raise BadRequest(f"{name} must be a whole number") from None
    if not lo <= n <= hi:
        raise BadRequest(f"{name} must be from {lo} to {hi}")
    return n


def _class(value: object) -> int:
    n = _int(value, "class", 1, 99)
    if n is None:
        raise BadRequest("class must be a whole number")
    return n


def _parse_roster(raw):
    if not isinstance(raw, list) or not 1 <= len(raw) <= MAX_STUDENTS:
        raise BadRequest(f"roster must be a list of 1 to {MAX_STUDENTS} students")
    roster, seen = [], set()
    for i, s in enumerate(raw, start=1):
        sid = str(s.get("id", "")).strip() if isinstance(s, dict) else ""
        name = CONTROL_RE.sub("", str(s.get("name", ""))).strip() if isinstance(s, dict) else ""
        if not ID_RE.match(sid):
            raise BadRequest(f"roster entry {i}: id must be 1 to 20 letters or digits")
        if not 1 <= len(name) <= 80:
            raise BadRequest(f"roster entry {i}: name must be 1 to 80 characters")
        if sid in seen:
            raise BadRequest(f"roster entry {i}: duplicate id")
        seen.add(sid)
        roster.append((sid, name))
    return roster


def _scan_photos(
    photos: list[Any], tab: str, cls: int, working_days: int | None = None, identify: bool = False, roster=None,
    subjects=None,
):
    if not 1 <= len(photos) <= MAX_PHOTOS:
        raise BadRequest(f"send 1 to {MAX_PHOTOS} photos for {tab}, in page order")
    if tab == "Attendance":
        max_values, allow_a = [working_days or 366], False
    else:
        # Use dynamic subjects configuration
        subject_config = next((s for s in (subjects or []) if s["key"] == tab), None)
        if subject_config and subject_config["slots"]:
            max_values = [slot["max"] for slot in subject_config["slots"]]
        else:
            raise BadRequest(f"no slot configuration found for tab {tab}; subjects must be provided")
        allow_a = True

    rows = []
    for n, photo in enumerate(photos, start=1):
        img = None
        try:
            img = reader.load_image(photo.read(), preserve_resolution=identify)
            page_rows = reader.read_page(img, max_values, NET, allow_a, identify=identify, roster=roster)
            if identify:
                rows.extend(
                    {**row, "photo": n, "page_row": i}
                    for i, row in enumerate(page_rows, start=1)
                )
            else:
                rows.extend(page_rows)
        except reader.ReadError as e:
            raise BadRequest(f"photo {n}: {e}") from None
        finally:
            img = None
    return rows


def _portal_cell(cell, maximum):
    raw = cell.get("value", "")
    blank = bool(cell.get("blank"))
    reasons = []
    if raw == "A" or (blank and raw == ""):
        value, valid = "A", True
    elif isinstance(raw, str) and re.fullmatch(r"\d{1,3}", raw) and int(raw) <= maximum:
        value, valid = int(raw), True
    else:
        value, valid = "", False
        reasons.append("invalid_value")

    confidence = cell.get("conf")
    if (isinstance(confidence, (int, float)) and not isinstance(confidence, bool)
            and math.isfinite(confidence) and 0 <= confidence <= 1):
        if confidence < REVIEW_CONFIDENCE:
            reasons.append("low_confidence")
    else:
        confidence = None
        reasons.append("confidence_unavailable")

    result = dict(cell)
    result["confidence"] = confidence
    result["needs_review"] = bool(reasons)
    result["review_reasons"] = reasons
    return value if valid else "", result


# ---------- auth and rate limit ----------
_hits, _lock = defaultdict(deque), threading.Lock()


def _who():
    """-> an id for the caller (a hash, never the secret itself), or None if the caller is not allowed."""
    origin = request.headers.get("Origin", "")
    if origin and origin in ALLOWED_ORIGINS:
        return "ext:" + hashlib.sha256(origin.encode()).hexdigest()[:8]
    if not API_KEYS:
        return "dev" if ALLOW_NO_AUTH else None
    sent = request.headers.get("X-API-Key", "").encode()
    found = None
    for k in API_KEYS:  # check every key so the time taken doesn't reveal which one matched
        if hmac.compare_digest(sent, k.encode()):
            found = hashlib.sha256(k.encode()).hexdigest()[:8]
    return found


def _client_ip():
    """The address Render recorded for this request (the last entry it appends to X-Forwarded-For)."""
    forwarded = request.headers.get("X-Forwarded-For", "")
    return forwarded.split(",")[-1].strip() or request.remote_addr or "unknown"


def _rate_ok(who, limit=None):
    now = time.monotonic()
    limit = RATE_PER_MIN if limit is None else limit
    with _lock:
        q = _hits[who]
        while q and now - q[0] > 60:
            q.popleft()
        if len(q) >= limit:
            return False
        q.append(now)
        return True


def protected(fn):
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        who = _who()
        if who is None:
            return _err(401, "missing or wrong API key")
        if not _rate_ok(who) or not _rate_ok("ip:" + _client_ip(), IP_RATE_PER_MIN):
            return _err(429, "too many requests; wait a minute and try again")
        try:
            return fn(*args, **kwargs)
        except BadRequest as e:
            return _err(400, str(e))
    return wrapper


@app.before_request
def preflight():
    if request.method == "OPTIONS":
        return "", 204


@app.after_request
def headers(resp):
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    origin = request.headers.get("Origin")
    if origin and origin in ALLOWED_ORIGINS:
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Access-Control-Allow-Headers"] = "X-API-Key, X-Install-Id, Content-Type"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        resp.headers["Vary"] = "Origin"
    return resp

@app.errorhandler(Exception)
def handle(e):
    if isinstance(e, HTTPException):          # e.g. unknown route: keep its status code
        return {"error": e.description}, e.code
    # Log the error type and route only. Messages and tracebacks can contain student names or marks.
    app.logger.error("unhandled %s on %s", type(e).__name__, request.path)
    return {"error": "server error"}, 500
# def unexpected(e):
#     if isinstance(e, HTTPException):
#         message = "upload is too large" if e.code == 413 else e.description
#         return _err(e.code, message)
#     app.logger.error("unexpected %s", type(e).__name__)  # the type only: messages can contain student data
#     return _err(500, "something went wrong on the server")


# ---------- endpoints ----------
@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.post("/sheets")
@protected
def sheets():
    body = request.get_json(silent=True)
    if not isinstance(body, dict):
        raise BadRequest("send a JSON body")
    cls = _int(body.get("class"), "class", 1, 99)
    if cls is None:
        raise BadRequest("class must be a whole number")
    output = body.get("output", "all")
    if output not in ("all", "excel", "pdf"):
        raise BadRequest("output must be 'excel', 'pdf', or 'all'")
    section = str(body.get("section", "")).strip().upper()
    if section not in SECTION_CODES:
        raise BadRequest("section must be one letter or digit (A-Z or 0-9)")
    school = CONTROL_RE.sub("", str(body.get("school", ""))).strip()[:120]
    academic_year = CONTROL_RE.sub("", str(body.get("academic_year", ""))).strip()[:40]
    exam_type = CONTROL_RE.sub("", str(body.get("exam_type", ""))).strip()[:80]
    working_days = _int(body.get("working_days"), "working_days", 1, 366, optional=True)
    roster = _parse_roster(body.get("roster"))

    # Dynamic subjects configuration
    subjects = body.get("subjects")
    if not subjects or not isinstance(subjects, list):
        raise BadRequest("subjects must be provided as an array")
    has_attendance = body.get("hasAttendance", True)

    with tempfile.TemporaryDirectory() as d:
        result = make_sheets.generate(
            roster, cls, section, d, school, working_days,
            academic_year=academic_year, exam_type=exam_type,
            output=output,
            subjects=subjects,
            has_attendance=has_attendance,
        )
        
        if output == "excel":
            if result is None:
                raise RuntimeError("Excel generation did not produce a workbook")
            with open(result, "rb") as workbook_file:
                workbook_data = io.BytesIO(workbook_file.read())
            return send_file(
                workbook_data,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=f"class{cls}{section}_entry.xlsx",
            )

        # output == "pdf" or "all" - return PDF only
        if result is None:
            raise RuntimeError("PDF generation did not produce a PDF")
        with open(result, "rb") as pdf_file:
            pdf_data = io.BytesIO(pdf_file.read())
        return send_file(
            pdf_data,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"class{cls}{section}_sheets.pdf",
        )

def _safe_xlsx(raw):
    """Refuse workbooks that would unpack to something huge (a zip bomb)."""
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            if sum(i.file_size for i in z.infolist()) > 40 * 1024 * 1024:
                raise BadRequest("the workbook is too large when unpacked")
    except zipfile.BadZipFile:
        raise BadRequest("this is not an .xlsx file") from None


@app.post("/merge")
@protected
def merge():
    f = request.files.get("workbook")
    if f is None:
        raise BadRequest("upload the filled workbook as the field 'workbook'")
    working_days = _int(request.form.get("working_days"), "working_days", 1, 366, optional=True)
    
    # Get subjects from request
    try:
        subjects = json.loads(request.form.get("subjects", "[]"))
    except (TypeError, ValueError):
        raise BadRequest("subjects must be a JSON array") from None
    if not subjects or not isinstance(subjects, list):
        raise BadRequest("subjects must be provided as an array")
    
    has_attendance = request.form.get("has_attendance", "true").lower() == "true"
    
    raw = f.read()
    _safe_xlsx(raw)
    try:
        wb = load_workbook(io.BytesIO(raw), data_only=True)
        data, errors, incomplete = merge_marks.merge_workbook(wb, subjects=subjects, has_attendance=has_attendance, working_days=working_days)
    except merge_marks.WorkbookError as e:
        raise BadRequest(str(e)) from None
    except BadRequest:
        raise
    except Exception:
        raise BadRequest("could not read this workbook; use the one made by /sheets") from None
    if errors:
        return jsonify({"ok": False, "errors": errors}), 422
    return jsonify({"ok": True, "data": data,
                    "incomplete": [{"id": i, "name": n, "missing": m} for i, n, m in incomplete]})


@app.post("/submit")
@protected
def submit():
    cls = _class(request.form.get("class"))
    try:
        raw_roster = json.loads(request.form.get("roster", ""))
    except (TypeError, ValueError):
        raise BadRequest("roster must be a JSON list of student objects") from None
    roster = _parse_roster(raw_roster)
    days = _int(request.form.get("working_days"), "working_days", 1, 366, optional=True)
    exam_type = request.form.get("exam_type", "").strip()

    # Get subjects from request
    try:
        subjects = json.loads(request.form.get("subjects", "[]"))
    except (TypeError, ValueError):
        raise BadRequest("subjects must be a JSON array") from None
    if not subjects or not isinstance(subjects, list):
        raise BadRequest("subjects must be provided as an array")

    tabs = [s["key"] for s in subjects]
    submitted_subjects = [tab for tab in tabs if request.files.getlist(f"photos_{tab}")]
    attendance_photos = request.files.getlist("photos_Attendance")
    generic_photos = request.files.getlist("photos")
    uploaded_fields = {
        key for key in request.files.keys()
        if (key == "photos" or key.startswith("photos_")) and request.files.getlist(key)
    }
    allowed_fields = {f"photos_{tab}" for tab in tabs} | {"photos_Attendance", "photos"}
    unexpected_fields = sorted(uploaded_fields - allowed_fields)
    if unexpected_fields:
        raise BadRequest(f"unknown photo field(s): {', '.join(unexpected_fields)}")
    if generic_photos and (submitted_subjects or attendance_photos):
        raise BadRequest("use either generic 'photos' uploads or named photos_<tab> fields, not both")
    if not submitted_subjects and not attendance_photos and not generic_photos:
        raise BadRequest("upload photos for at least one subject or Attendance tab")

    photos_by_tab = {}
    section = request.form.get("section") or None
    if generic_photos:
        if len(generic_photos) > MAX_PHOTOS * (len(tabs) + 1):
            raise BadRequest(f"send no more than {MAX_PHOTOS * (len(tabs) + 1)} generic photos in one request")
        sections = set()
        for photo_number, photo in enumerate(generic_photos, start=1):
            image = None
            try:
                image_bytes = photo.read()
                photo.stream.seek(0)
                image = reader.load_image(image_bytes, preserve_resolution=True)
                metadata = reader.read_sheet_metadata(image, tab_codes=tabs)
            except reader.ReadError as e:
                raise BadRequest(f"photo {photo_number}: {e}") from None
            finally:
                image = None
            if metadata["class"] != cls:
                return _err(422, f"photo {photo_number}: sheet class {metadata['class']} does not match request class {cls}",
                            photo=photo_number, issue="class_mismatch")
            if metadata["tab"] not in tabs and metadata["tab"] != "Attendance":
                return _err(422, f"photo {photo_number}: sheet tab {metadata['tab']} is not valid for class {cls}",
                            photo=photo_number, issue="tab_not_in_class")
            # Validate exam type if both are present
            if exam_type and metadata.get("exam_type") and metadata["exam_type"] != exam_type:
                return _err(422, f"photo {photo_number}: sheet exam type {metadata['exam_type']} does not match portal exam type {exam_type}",
                            photo=photo_number, issue="exam_type_mismatch")
            sections.add(metadata["section"])
            photos_by_tab.setdefault(metadata["tab"], []).append(photo)
        if len(sections) != 1:
            return _err(422, "photos contain sheets from multiple sections", issue="section_mismatch")
        detected_section = next(iter(sections))
        requested_section = request.form.get("section")
        if requested_section and requested_section != detected_section:
            return _err(422, f"sheet section {detected_section} does not match request section {requested_section}",
                        issue="section_mismatch")
        section = detected_section
        submitted_subjects = [tab for tab in tabs if tab in photos_by_tab]
        attendance_photos = photos_by_tab.get("Attendance", [])
        if not submitted_subjects and not attendance_photos:
            raise BadRequest("photos did not contain a subject tab supported for this class")
    else:
        photos_by_tab = {
            tab: request.files.getlist(f"photos_{tab}") for tab in submitted_subjects
        }
        if attendance_photos:
            photos_by_tab["Attendance"] = attendance_photos

    roster_by_id = dict(roster)
    identified = {}
    identity_issues = []
    for tab in submitted_subjects:
        photos = photos_by_tab[tab]
        rows = _scan_photos(photos, tab, cls, days, identify=True, roster=roster, subjects=subjects)
        identified[tab] = {}
        for row in rows:
            sid = row["student_id"]
            location = {"tab": tab, "photo": row["photo"], "row": row["page_row"]}
            # Don't treat "roster_fallback" or "interpolated" as errors - they're legitimate fallbacks
            if row["qr_error"] and row["qr_error"] not in ("roster_fallback", "interpolated"):
                identity_issues.append({**location, "issue": f"qr_{row['qr_error']}"})
            elif sid not in roster_by_id:
                identity_issues.append({**location, "issue": "student_id_not_in_roster", "student_id": sid})
            elif sid in identified[tab]:
                identity_issues.append({**location, "issue": "duplicate_student_id", "student_id": sid})
            else:
                identified[tab][sid] = row

    if "Attendance" in photos_by_tab:
        rows = _scan_photos(photos_by_tab["Attendance"], "Attendance", cls, days, identify=True, roster=roster, subjects=subjects)
        identified["Attendance"] = {}
        for row in rows:
            sid = row["student_id"]
            location = {"tab": "Attendance", "photo": row["photo"], "row": row["page_row"]}
            # Don't treat "roster_fallback" or "interpolated" as errors - they're legitimate fallbacks
            if row["qr_error"] and row["qr_error"] not in ("roster_fallback", "interpolated"):
                identity_issues.append({**location, "issue": f"qr_{row['qr_error']}"})
            elif sid not in roster_by_id:
                identity_issues.append({**location, "issue": "student_id_not_in_roster", "student_id": sid})
            elif sid in identified["Attendance"]:
                identity_issues.append({**location, "issue": "duplicate_student_id", "student_id": sid})
            else:
                identified["Attendance"][sid] = row

    if identity_issues:
        return _err(
            422,
            "student QR validation failed; use QR-coded sheets from /sheets, keep every code visible, "
            "and include every roster student exactly once per tab; no portal data was generated",
            issues=identity_issues,
        )
    missing = [
        {"tab": tab, "student_id": sid, "student_name": name, "issue": "student_not_on_any_photo"}
        for tab, students in identified.items()
        for sid, name in roster if sid not in students
    ]
    if missing and len(missing) > 4:
        return _err(
            422,
            f"{len(missing)} roster student(s) were not found on their tab's photos; send every page for each "
            "subject in one request, or check the roster matches the sheets",
            issues=missing,
        )

    data: dict[str, dict[str, Any]] = {
        sid: {"marks": {tab: [] for tab in submitted_subjects}}
        for sid, _ in roster
    }
    if attendance_photos:
        for sid in data:
            data[sid]["attendance"] = "A"
    review_scans, review_fields = {}, []
    for tab, students in identified.items():
        enriched = []
        for sid, name in roster:
            if sid not in students:
                # Student not found in photos - skip and will need manual entry
                continue
            row = students[sid]
            cells = []
            if tab == "Attendance":
                maximum = days or 366
                value, cell = _portal_cell(row["cells"][0], maximum)
                data[sid]["attendance"] = value
                cell["field"] = "Attendance"
                cells.append(cell)
                if cell["needs_review"]:
                    review_fields.append({
                        "tab": tab, "student_id": sid, "student_name": name,
                        "field": "Attendance", "value": cell["value"],
                        "confidence": cell["confidence"], "reasons": cell["review_reasons"],
                    })
            else:
                values = []
                # Get slot configuration for this subject
                subject_config = next((s for s in subjects if s["key"] == tab), None)
                slots = subject_config["slots"] if subject_config else []
                for cell, slot in zip(row["cells"], slots):
                    maximum = slot["max"]
                    value, result = _portal_cell(cell, maximum)
                    result["field"] = slot["name"]
                    values.append(value)
                    cells.append(result)
                    if result["needs_review"]:
                        review_fields.append({
                            "tab": tab, "student_id": sid, "student_name": name,
                            "field": slot["name"], "value": result["value"],
                            "confidence": result["confidence"], "reasons": result["review_reasons"],
                        })
                data[sid]["marks"][tab] = values
            enriched.append({
                "student_id": sid, "name": name, "identity_verified": True, "cells": cells,
            })
        review_scans[tab] = enriched

    incomplete = []
    missing_students = []  # Students not found in photos (need manual entry)
    for sid, name in roster:
        missing = []
        record = data[sid]
        if attendance_photos and record["attendance"] == "":
            missing.append("attendance")
        for tab, values in record["marks"].items():
            subject_config = next((s for s in subjects if s["key"] == tab), None)
            label = subject_config["label"] if subject_config else tab
            slots = subject_config["slots"] if subject_config else []
            missing.extend(
                f"{label} {slots[i]['name']}" for i, value in enumerate(values) if value == ""
            )
        if missing:
            incomplete.append({"id": sid, "name": name, "missing": missing})

        # Check if student was identified in any photo
        student_found = False
        has_interpolated = False
        for tab, students in identified.items():
            if sid in students:
                row = students[sid]
                # Check if this student had interpolated QR code
                if row.get("qr_error") == "interpolated":
                    has_interpolated = True
                student_found = True
                break

        # Only add to missing_students if truly missing or interpolated
        if has_interpolated:
            missing_students.append({"id": sid, "name": name, "reason": "interpolated"})
        elif not student_found:
            missing_students.append({"id": sid, "name": name, "reason": "not_found"})

    return jsonify({
        "ok": True, "class": cls, "section": section, "tabs": list(identified),
        "data": data, "scans": review_scans,
        "review": review_fields, "review_confidence_threshold": REVIEW_CONFIDENCE,
        "incomplete": incomplete,
        "missing_students": missing_students,
    })


if __name__ == "__main__":  # local testing only; on the host the app is started by gunicorn
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "8000")))
