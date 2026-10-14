#!/usr/bin/env python3
"""Make printable marks sheets (PDF) and a typing workbook (XLSX) for one class + section.

Usage:
  python make_sheets.py roster.csv --class 3 --section A [--working-days 37] [--school "NAME"] [--out out]

roster.csv comes from scrapeRoster() in portal_tools.js (columns: student_id,name).
Output, inside --out:
  sheets/class3A_<Subject>.pdf     one printable sheet per subject, for that subject's teacher
  sheets/class3A_Attendance.pdf    attendance sheet
  entry/class3A_entry.xlsx         ONE Excel workbook to type the marks into: a tab per subject plus
                                   Attendance. IDs and names are locked; Tab moves between the boxes.
The rows keep the roster's order, which is the portal's order.

Every page carries AprilTags: four corner tags (for straightening), a tag per student row (for identity),
and three tags at the bottom (class, section and subject tab), so a plain phone camera can read all of them.
"""
import argparse
import csv
import io
import json
import os
import sys
from typing import Any
from xml.sax.saxutils import escape

import cv2
import numpy as np
from moms_apriltag import TagGenerator3
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Flowable, Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from config import (
    EXAM_TYPES, MAX_STUDENTS, META_CLASS_BASE, META_EXAM_BASE, META_SECTION_BASE, META_TAB_BASE,
    SECTION_CODES, STUDENT_TAG_ID_START,
)

MARGIN = 36
ROW_H = 27
APRILTAG_SIZE = 22          # student tags, points (fits the 27 pt row with padding)
CORNER_APRILTAG_SIZE = 40   # corner tags, points
META_TAG_SIZE = 22          # class / section / tab tags, points
CORNER_CENTER_PT = 38       # corner tag centres, from the page edges; the tags span 18-58 pt in, clear of printer margins
HEADER_TOP = 66             # distance from the top of the page to the title block (below the top corner tags)
PRINT_DPI = 300             # tags are rendered at this resolution so they print sharply
PAGE_WIDTH, PAGE_HEIGHT = A4

TAG_FAMILY = "tagStandard41h12"  # must match reader.py
CORNER_TAG_IDS = [0, 1, 2, 3]    # TL, TR, BR, BL

_apriltag_generator = TagGenerator3(TAG_FAMILY)


def read_roster(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if not rows or "student_id" not in rows[0] or "name" not in rows[0]:
        sys.exit("roster.csv needs the columns: student_id,name")
    roster, seen = [], set()
    for n, r in enumerate(rows, start=2):
        sid, name = r["student_id"].strip(), r["name"].strip()
        if not sid or not name:
            sys.exit(f"roster.csv line {n}: empty student_id or name")
        if sid in seen:
            sys.exit(f"roster.csv line {n}: duplicate student_id {sid}")
        seen.add(sid)
        roster.append((sid, name))
    return roster


# ---------- AprilTags ----------
def _apriltag_png(tag_id, size_pt):
    """AprilTag with a white quiet zone, rendered at print resolution, as PNG bytes in memory."""
    tag = np.clip(np.asarray(_apriltag_generator.generate(tag_id, scale=3), dtype=np.float32), 0, 255)
    gray = tag.astype(np.uint8)  # generate() returns 0..255 values
    if gray.min() == gray.max():
        raise RuntimeError(f"AprilTag {tag_id} generated as a single colour; check moms_apriltag")
    margin = max(1, gray.shape[0] // 8)  # white border of about one cell
    gray = cv2.copyMakeBorder(gray, margin, margin, margin, margin, cv2.BORDER_CONSTANT, value=255)
    px = max(gray.shape[0], round(size_pt / 72 * PRINT_DPI))
    rgb = cv2.cvtColor(cv2.resize(gray, (px, px), interpolation=cv2.INTER_NEAREST), cv2.COLOR_GRAY2RGB)
    _, buffer = cv2.imencode(".png", rgb)
    return io.BytesIO(buffer.tobytes())


def _apriltag_flowable(tag_id, size_pt):
    """AprilTag as a reportlab Image, for placing inside a table cell."""
    return Image(_apriltag_png(tag_id, size_pt), width=size_pt, height=size_pt)


def _corner_centres():
    c = CORNER_CENTER_PT
    return [
        (c, PAGE_HEIGHT - c),                 # TL
        (PAGE_WIDTH - c, PAGE_HEIGHT - c),    # TR
        (PAGE_WIDTH - c, c),                  # BR
        (c, c),                               # BL
    ]


def _draw_corner_codes(canvas):
    half = CORNER_APRILTAG_SIZE / 2
    for tag_id, (cx, cy) in zip(CORNER_TAG_IDS, _corner_centres()):
        canvas.drawImage(
            ImageReader(_apriltag_png(tag_id, CORNER_APRILTAG_SIZE)),
            cx - half, cy - half,
            width=CORNER_APRILTAG_SIZE, height=CORNER_APRILTAG_SIZE,
        )


def _draw_metadata_tags(canvas, cls, section, tab_key, tab_codes, exam_type):
    """Four tags centred at the bottom of the page: class, section, subject tab, and exam type."""
    ids = [
        META_CLASS_BASE + cls,
        META_SECTION_BASE + SECTION_CODES.index(section),
        META_TAB_BASE + tab_codes.index(tab_key),
    ]
    
    # Add exam type tag if provided
    if exam_type and exam_type in EXAM_TYPES:
        ids.append(META_EXAM_BASE + EXAM_TYPES.index(exam_type))
    
    gap = 6
    total = len(ids) * META_TAG_SIZE + (len(ids) - 1) * gap
    x0 = PAGE_WIDTH / 2 - total / 2
    y = 20
    for i, tag_id in enumerate(ids):
        canvas.drawImage(
            ImageReader(_apriltag_png(tag_id, META_TAG_SIZE)),
            x0 + i * (META_TAG_SIZE + gap), y,
            width=META_TAG_SIZE, height=META_TAG_SIZE,
        )


# ---------- PDF sheets ----------
def build_pdf(
    path, title, subtitle, footer_text, roster, extra_headers, extra_widths, digit_counts,
    meta, tab_codes, exam_type,
):
    """meta = (class number, section, subject tab key), printed as tags at the bottom of every page."""
    cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=8.5, leading=9.5)
    head = ParagraphStyle("head", fontName="Helvetica-Bold", fontSize=8, leading=9.5, alignment=1)
    h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=15, leading=19)
    small = ParagraphStyle("small", fontName="Helvetica", fontSize=9, leading=12, textColor=colors.HexColor("#333333"))

    column_headers = [Paragraph(t, head) for t in ["No.", "Student ID", "Student name", "Student code"]]
    digit_headers = []
    for text, count in zip(extra_headers, digit_counts):
        digit_headers.append(Paragraph(text, head))
        digit_headers.extend([""] * (count - 1))
    data: list[list[Any]] = [column_headers + digit_headers]
    for n, (sid, name) in enumerate(roster, start=1):
        data.append([str(n), sid, Paragraph(escape(name), cell),
                     _apriltag_flowable(STUDENT_TAG_ID_START + n - 1, APRILTAG_SIZE)]
                    + [""] * sum(digit_counts))
    widths = [26, 68, 126, 48] + [
        width / count for width, count in zip(extra_widths, digit_counts) for _ in range(count)
    ]

    style = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e6e6e6")),
        ("BACKGROUND", (0, 1), (2, -1), colors.HexColor("#f6f6f6")),
        ("FONTNAME", (0, 1), (2, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (2, -1), 8.5),
        ("ALIGN", (0, 0), (2, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 1), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 2),
    ]
    col = 4
    for count in digit_counts:
        if count > 1:
            style.append(("SPAN", (col, 0), (col + count - 1, 0)))
        col += count
    t = Table(data, colWidths=widths, rowHeights=[48] + [ROW_H] * len(roster),
              repeatRows=1, style=TableStyle(style))

    header = Table(
        [[[Paragraph(title, h1), Spacer(1, 3), Paragraph(subtitle, small)]]],
        colWidths=[A4[0] - 2 * MARGIN],
        style=TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]),
    )
    _, header_height = header.wrap(A4[0] - 2 * MARGIN, 10000)

    def footer(canvas, doc):
        canvas.saveState()
        header.drawOn(canvas, MARGIN, A4[1] - HEADER_TOP - header_height)
        canvas.restoreState()
        _draw_corner_codes(canvas)
        _draw_metadata_tags(canvas, *meta, tab_codes, exam_type)
        canvas.setFont("Helvetica", 8)
        canvas.drawString(MARGIN + 30, 22, footer_text)
        canvas.drawRightString(A4[0] - MARGIN - 30, 22, f"Page {doc.page}")

    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN,
                        topMargin=HEADER_TOP + header_height + 10, bottomMargin=66)
    story: list[Flowable] = [t]
    doc.build(story, onFirstPage=footer, onLaterPages=footer)


# ---------- Excel workbook ----------
def write_workbook(path, tag, where, roster, subjects, working_days):
    """One workbook: Read me, Attendance, then a tab per subject. Only the yellow input boxes are unlocked,
    so Tab jumps between them. Protection has no password (Review > Unprotect Sheet removes it)."""
    arial = lambda **kw: Font(name="Arial", size=kw.pop("size", 10), **kw)
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    grey = PatternFill("solid", fgColor="E6E6E6")
    yellow = PatternFill("solid", fgColor="FFF9C4")
    n = len(roster)

    wb = Workbook()
    rm = wb.active
    rm.title = "Read me"
    rm.column_dimensions["A"].width = 110
    lines = [
        ("How to fill in this workbook", arial(size=14, bold=True)),
        (where, arial(color="555555")),
        ("", arial()),
        ("1. Type from the printed sheet the teacher filled in. Each subject has its own tab, plus an Attendance tab.", arial()),
        ("2. Only the yellow boxes can be typed in. Student IDs and names are locked.", arial()),
        ("3. Press Tab to move right across the boxes; at the end of a row it jumps to the next student.", arial()),
        ("4. Type a whole number (not above the maximum in the column heading), or A if the student was absent.", arial()),
        ("5. A blank mark box is treated as A (absent).", arial()),
        ("6. If Tab does not jump as described, or you need to change something locked: Review > Unprotect Sheet (no password).", arial()),
        ("", arial()),
        ("Example: 8, 9, A, 15 (Tool 4 is out of 20); blank mark boxes also mean A. Attendance is days attended, or A if absent.", arial(italic=True)),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = rm.cell(row=i, column=1, value=text)
        c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top")

    def make_sheet(title, input_heads, widths, validations):
        ws = wb.create_sheet(title)
        heads = ["Student ID", "Student name"] + input_heads
        for j, h in enumerate(heads, start=1):
            c = ws.cell(row=1, column=j, value=h)
            c.font = arial(bold=True)
            c.fill = grey
            c.border = box
            c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 50
        for j, w in enumerate([16, 34] + widths, start=1):
            ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = w
        for i, (sid, name) in enumerate(roster, start=2):
            a = ws.cell(row=i, column=1, value=sid)  # stored as text so long IDs never turn into 1.5E+09
            a.number_format = "@"
            b = ws.cell(row=i, column=2, value=name)
            for c in (a, b):
                c.font = arial()
                c.border = box
                c.fill = PatternFill("solid", fgColor="F6F6F6")
            for j in range(3, 3 + len(input_heads)):
                c = ws.cell(row=i, column=j)
                c.font = arial()
                c.border = box
                c.fill = yellow
                c.alignment = Alignment(horizontal="center")
                c.protection = Protection(locked=False)
        for dv in validations:
            ws.add_data_validation(dv)
        ws.freeze_panes = "C2"
        ws.protection.sheet = True
        ws.protection.selectLockedCells = True  # locked cells cannot be selected, so Tab skips them
        return ws

    att_head = "Days attended" + (f"\n(out of {working_days})" if working_days else "")
    att_dv = DataValidation(
        type="custom",
        formula1=f'IF(ISNUMBER(C2),AND(C2>=0,C2<={working_days or 366},C2=INT(C2)),UPPER(C2)="A")',
        allow_blank=True, showErrorMessage=True, errorStyle="stop", errorTitle="Attendance",
        error=f"Type a whole number from 0 to {working_days or 366}, or A if absent.",
    )
    att_dv.add(f"C2:C{n + 1}")
    make_sheet("Attendance", [att_head], [16], [att_dv])

    for subject in subjects:
        key = subject["key"]
        label = subject["label"]
        slots = subject["slots"]
        heads, dvs = [], []
        for k, slot in enumerate(slots):
            col = "CDEFGHIJKLMNOPQRSTUVWXYZ"[k]
            name = slot["name"]
            mx = slot["max"]
            heads.append(f"{name}\n(max {mx})")
            dv = DataValidation(
                type="custom",
                formula1=f'IF(ISNUMBER({col}2),AND({col}2>=0,{col}2<={mx},{col}2=INT({col}2)),UPPER({col}2)="A")',
                allow_blank=True, showErrorMessage=True, errorStyle="stop", errorTitle=f"{name} (max {mx})",
                error=f"Type a whole number from 0 to {mx}, or A if absent.")
            dv.add(f"{col}2:{col}{n + 1}")
            dvs.append(dv)
        make_sheet(label, heads, [15] * len(slots), dvs)

    wb.save(path)


# ---------- entry point used by the server and the command line ----------
def generate(
    roster, cls, section, out_dir, school="", working_days=None,
    academic_year="", exam_type="", output="all",
    subjects=None, has_attendance=True,
):
    """Make selected output for one class + section. roster = [(student_id, name), ...].
    Returns the workbook path when output includes Excel."""
    if output not in {"all", "excel", "pdf"}:
        raise ValueError("output must be 'all', 'excel', or 'pdf'")
    if section not in SECTION_CODES:
        raise ValueError(f"section must be one letter or digit: {SECTION_CODES}")
    if not 1 <= len(roster) <= MAX_STUDENTS:
        raise ValueError(f"roster must have 1 to {MAX_STUDENTS} students")
    if subjects is None or not isinstance(subjects, list):
        raise ValueError("subjects must be provided as a list")

    # Build dynamic TAB_CODES
    tab_codes = [s["key"] for s in subjects]
    if has_attendance:
        tab_codes.append("Attendance")
    
    # Normalize exam type for encoding
    normalized_exam_type = exam_type if exam_type in EXAM_TYPES else ""

    tag = f"class{cls}{section}"
    sheets, entry = os.path.join(out_dir, "sheets"), os.path.join(out_dir, "entry")
    if output in {"all", "pdf"}:
        os.makedirs(sheets, exist_ok=True)
    if output in {"all", "excel"}:
        os.makedirs(entry, exist_ok=True)
    details = [f"Class {cls} - Section {section}"]
    if school:
        details.append(school)
    if academic_year:
        details.append(f"Academic year: {academic_year}")
    if exam_type:
        details.append(f"Exam type: {exam_type}")
    where = "  |  ".join(details)
    where_pdf = escape(where)  # the PDF text is markup, so & and < must be escaped

    if output in {"all", "pdf"}:
        for subject in subjects:
            key = subject["key"]
            label = subject["label"]
            slots = subject["slots"]
            fname = label.replace(" ", "_")
            headers = [f"{slot['name']}<br/>({slot['max']} marks)" for slot in slots]
            build_pdf(
                os.path.join(sheets, f"{tag}_{fname}.pdf"),
                f"{label} - marks sheet",
                f"{where_pdf}<br/>Keep the student code visible. Write one digit per box, right-aligned. "
                f"Leave both boxes empty to record <b>A</b> (absent). "
                f"Teacher: ____________________   Date: ____________",
                f"{tag}  |  {label}", roster, headers, [60] * len(slots),
                [len(str(slot["max"])) for slot in slots],
                (cls, section, key),
                tab_codes,
                normalized_exam_type,
            )

        if has_attendance:
            att_head = "Days attended" + (f"<br/>(out of {working_days})" if working_days else "")
            build_pdf(
                os.path.join(sheets, f"{tag}_Attendance.pdf"),
                "Attendance sheet",
                f"{where_pdf}<br/>Keep the student code visible. Write the number of days each student attended."
                + (f" Working days: {working_days}." if working_days else "")
                + " Right-align the digits. Leave every box empty to record A (absent).",
                f"{tag}  |  Attendance",
                roster, [att_head], [100], [len(str(working_days or 366))],
                (cls, section, "Attendance"),
                tab_codes,
                normalized_exam_type,
            )

    if output in {"all", "excel"}:
        workbook = os.path.join(entry, f"{tag}_entry.xlsx")
        write_workbook(workbook, tag, where, roster, subjects, working_days)
        return workbook
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("roster")
    ap.add_argument("--class", dest="cls", type=int, required=True)
    ap.add_argument("--section", required=True)
    ap.add_argument("--working-days", type=int)
    ap.add_argument("--school", default="")
    ap.add_argument("--academic-year", default="")
    ap.add_argument("--exam-type", default="")
    ap.add_argument("--out", default="out")
    ap.add_argument("--subjects", required=True, help="JSON file with subjects configuration")
    a = ap.parse_args()

    if a.section not in SECTION_CODES:
        sys.exit(f"section must be one letter or digit: {SECTION_CODES}")
    roster = read_roster(a.roster)
    if len(roster) > MAX_STUDENTS:
        sys.exit(f"roster has {len(roster)} students; the maximum is {MAX_STUDENTS}")

    # Load subjects from JSON file
    with open(a.subjects, 'r') as f:
        subjects_data = json.load(f)

    generate(
        roster, a.cls, a.section, a.out, a.school, a.working_days,
        academic_year=a.academic_year, exam_type=a.exam_type,
        subjects=subjects_data, has_attendance=True,
    )
    tag = f"class{a.cls}{a.section}"
    print(f"{len(roster)} students, {len(subjects_data)} subjects + attendance -> {a.out}/ (typing workbook: entry/{tag}_entry.xlsx)")


if __name__ == "__main__":
    main()