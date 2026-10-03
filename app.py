"""Marks pipeline server (Flask). Built to run on a small free instance (512 MB RAM, a fraction of a CPU).

Endpoints (all but /health need the header  X-API-Key: <one of API_KEYS>):
  GET  /health   -> {"ok": true}                      no data, for the host's health check
  POST /sheets   JSON {class, section, school?, academic_year?, exam_type?, working_days?, roster:[{id,name}]}
                 -> a zip: printable PDF sheets + the Excel typing workbook
  POST /merge    multipart: workbook=<filled .xlsx>, class, working_days?
                 -> {"ok":true,"data":{...},"incomplete":[...]}  or 422 {"ok":false,"errors":[...]}
  POST /scan     multipart: photos=<1-4 images, in page order>, class, tab, working_days?, expected_rows?
                 -> {"rows":[[{value,conf,blank,png} x cells] x students]}   tab = "Attendance" or a subject key

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
import os
import re
import sys
import tempfile
import threading
import time
import zipfile
from collections import defaultdict, deque

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
from config import CLASS_SUBJECTS, SLOTS  # noqa: E402
from digits import load_model  # noqa: E402


def _env_list(name):
    return [x.strip() for x in os.environ.get(name, "").split(",") if x.strip()]


API_KEYS = _env_list("API_KEYS")
ALLOW_NO_AUTH = os.environ.get("ALLOW_NO_AUTH") == "1"
ALLOWED_ORIGINS = set(_env_list("ALLOWED_ORIGINS"))
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "20"))
RATE_PER_MIN = int(os.environ.get("RATE_PER_MIN", "30"))
MODEL_PATH = os.environ.get("MODEL_PATH", os.path.join(HERE, "model.json"))

if not API_KEYS and not ALLOW_NO_AUTH:
    raise RuntimeError("Set API_KEYS (comma-separated secret keys). For local testing only, set ALLOW_NO_AUTH=1.")

NET = load_model(MODEL_PATH)
MAX_STUDENTS, MAX_PHOTOS = 120, 4
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


def _int(value, name, lo, hi, optional=False):
    if value in (None, "") and optional:
        return None
    try:
        n = int(value)
    except (TypeError, ValueError):
        raise BadRequest(f"{name} must be a whole number") from None
    if not lo <= n <= hi:
        raise BadRequest(f"{name} must be from {lo} to {hi}")
    return n


def _class(value):
    n = _int(value, "class", 1, 99)
    if n not in CLASS_SUBJECTS:
        raise BadRequest(f"class must be one of {sorted(CLASS_SUBJECTS)}")
    return n


# ---------- auth and rate limit ----------
_hits, _lock = defaultdict(deque), threading.Lock()


def _who():
    """-> an id for the caller (a hash, never the key itself), or None if the key is missing or wrong."""
    if not API_KEYS:
        return "dev"
    sent = request.headers.get("X-API-Key", "").encode()
    found = None
    for k in API_KEYS:  # check every key so the time taken doesn't reveal which one matched
        if hmac.compare_digest(sent, k.encode()):
            found = hashlib.sha256(k.encode()).hexdigest()[:8]
    return found


def _rate_ok(who):
    now = time.monotonic()
    with _lock:
        q = _hits[who]
        while q and now - q[0] > 60:
            q.popleft()
        if len(q) >= RATE_PER_MIN:
            return False
        q.append(now)
        return True


def protected(fn):
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        who = _who()
        if who is None:
            return _err(401, "missing or wrong API key")
        if not _rate_ok(who):
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
        resp.headers["Access-Control-Allow-Headers"] = "X-API-Key, Content-Type"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        resp.headers["Vary"] = "Origin"
    return resp


@app.errorhandler(Exception)
def unexpected(e):
    if isinstance(e, HTTPException):
        message = "upload is too large" if e.code == 413 else e.description
        return _err(e.code, message)
    app.logger.error("unexpected %s", type(e).__name__)  # the type only: messages can contain student data
    return _err(500, "something went wrong on the server")


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
    cls = _class(body.get("class"))
    section = str(body.get("section", "")).strip()
    if not SECTION_RE.match(section):
        raise BadRequest("section must be 1 to 3 letters or digits")
    school = CONTROL_RE.sub("", str(body.get("school", ""))).strip()[:120]
    academic_year = CONTROL_RE.sub("", str(body.get("academic_year", ""))).strip()[:40]
    exam_type = CONTROL_RE.sub("", str(body.get("exam_type", ""))).strip()[:80]
    working_days = _int(body.get("working_days"), "working_days", 1, 366, optional=True)
    raw = body.get("roster")
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

    with tempfile.TemporaryDirectory() as d:  # deleted as soon as the zip is built
        make_sheets.generate(
            roster, cls, section, d, school, working_days,
            academic_year=academic_year, exam_type=exam_type,
        )
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for sub in ("sheets", "entry"):
                for f in sorted(os.listdir(os.path.join(d, sub))):
                    z.write(os.path.join(d, sub, f), arcname=f"{sub}/{f}")
    buf.seek(0)
    return send_file(buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"class{cls}{section}_sheets.zip")


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
    cls = _class(request.form.get("class"))
    working_days = _int(request.form.get("working_days"), "working_days", 1, 366, optional=True)
    raw = f.read()
    _safe_xlsx(raw)
    try:
        wb = load_workbook(io.BytesIO(raw), data_only=True)
        data, errors, incomplete = merge_marks.merge_workbook(wb, cls, working_days)
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


@app.post("/scan")
@protected
def scan():
    photos = request.files.getlist("photos")
    if not 1 <= len(photos) <= MAX_PHOTOS:
        raise BadRequest(f"send 1 to {MAX_PHOTOS} photos as the field 'photos', in page order")
    cls = _class(request.form.get("class"))
    tab = request.form.get("tab", "")
    if tab == "Attendance":
        days = _int(request.form.get("working_days"), "working_days", 1, 366, optional=True)
        max_values, allow_a = [days or 40], False
    elif tab in CLASS_SUBJECTS[cls]:
        max_values, allow_a = [s["max"] if isinstance(s, dict) else s[2] for s in SLOTS], True
    else:
        raise BadRequest(f"tab must be 'Attendance' or one of {CLASS_SUBJECTS[cls]}")
    expected = _int(request.form.get("expected_rows"), "expected_rows", 1, MAX_STUDENTS, optional=True)

    rows = []
    for n, photo in enumerate(photos, start=1):
        try:
            img = reader.load_image(photo.read())
            rows.extend(reader.read_page(img, max_values, NET, allow_a))
        except reader.ReadError as e:
            raise BadRequest(f"photo {n}: {e}") from None
        finally:
            img = None  # let go of the big image before the next one (the host has little memory)
    if expected is not None and len(rows) != expected:
        return _err(422, f"found {len(rows)} rows in the photos but expected {expected}; "
                         "check the photos are in page order and show the whole table", rows_found=len(rows))
    return jsonify({"tab": tab, "rows": rows})


if __name__ == "__main__":  # local testing only; on the host the app is started by gunicorn
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "8000")))
