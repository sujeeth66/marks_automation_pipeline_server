#!/usr/bin/env python3
"""Make printable marks sheets (PDF) and blank entry files (CSV) for one class + section.

Usage:
  python make_sheets.py roster.csv --class 3 --section A [--working-days 37] [--school "NAME"] [--out out]

roster.csv comes from scrapeRoster() in portal_tools.js (columns: student_id,name).
Output, inside --out:
  sheets/class3A_<Subject>.pdf     one printable sheet per subject, for that subject's teacher
  sheets/class3A_Attendance.pdf    attendance sheet
  entry/class3A_entry.xlsx         ONE Excel workbook to type the marks into: a tab per subject plus
                                   Attendance. IDs and names are locked; Tab moves between the boxes.
The rows keep the roster's order, which is the portal's order.
"""
import argparse
import csv
import os
import sys
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from config import CLASS_SUBJECTS, SLOTS, SUBJECT_LABELS

MARGIN = 36
ROW_H = 30


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


def build_pdf(path, title, subtitle, footer_text, roster, extra_headers, extra_widths):
    cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=8.5, leading=9.5)
    head = ParagraphStyle("head", fontName="Helvetica-Bold", fontSize=8, leading=9.5, alignment=1)
    h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=15, leading=19)
    small = ParagraphStyle("small", fontName="Helvetica", fontSize=9, leading=12, textColor=colors.HexColor("#333333"))

    header = [Paragraph(t, head) for t in ["No.", "Student ID", "Student name"] + extra_headers]
    data = [header]
    for n, (sid, name) in enumerate(roster, start=1):
        data.append([str(n), sid, Paragraph(escape(name), cell)] + [""] * len(extra_headers))
    widths = [28, 74, 146] + extra_widths

    t = Table(data, colWidths=widths, rowHeights=[40] + [ROW_H] * len(roster), repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e6e6e6")),
        ("BACKGROUND", (0, 1), (1, -1), colors.HexColor("#f6f6f6")),
        ("FONTNAME", (0, 1), (1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (1, -1), 8.5),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    def footer(canvas, doc):
        canvas.setFont("Helvetica", 8)
        canvas.drawString(MARGIN, 22, footer_text)
        canvas.drawRightString(A4[0] - MARGIN, 22, f"Page {doc.page}")

    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=MARGIN, bottomMargin=40)
    story = [Paragraph(title, h1), Spacer(1, 3), Paragraph(subtitle, small), Spacer(1, 8), t]
    doc.build(story, onFirstPage=footer, onLaterPages=footer)


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
        ("5. Leave a box empty if that mark is not available yet. Do not type anything else (no decimals, no text).", arial()),
        ("6. If Tab does not jump as described, or you need to change something locked: Review > Unprotect Sheet (no password).", arial()),
        ("", arial()),
        ("Example of one student's boxes on a subject tab: 8, 9, A, 15   (Tool 4 is out of 20)    On the Attendance tab: 34", arial(italic=True)),
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
    att_dv = DataValidation(type="whole", operator="between", formula1="0", formula2=str(working_days or 366),
                            allow_blank=True, showErrorMessage=True, errorStyle="stop", errorTitle="Attendance",
                            error=f"Type a whole number from 0 to {working_days or 366}.")
    att_dv.add(f"C2:C{n + 1}")
    make_sheet("Attendance", [att_head], [16], [att_dv])

    for key in subjects:
        label = SUBJECT_LABELS[key]
        heads, dvs = [], []
        for k, (name, meaning, mx) in enumerate(SLOTS):
            col = "CDEFGH"[k]
            heads.append(f"{name}\n{meaning}\n(max {mx})")
            dv = DataValidation(
                type="custom",
                formula1=f'IF(ISNUMBER({col}2),AND({col}2>=0,{col}2<={mx},{col}2=INT({col}2)),UPPER({col}2)="A")',
                allow_blank=True, showErrorMessage=True, errorStyle="stop", errorTitle=f"{name} (max {mx})",
                error=f"Type a whole number from 0 to {mx}, or A if absent.")
            dv.add(f"{col}2:{col}{n + 1}")
            dvs.append(dv)
        make_sheet(label, heads, [15] * len(SLOTS), dvs)

    wb.save(path)


def generate(
    roster, cls, section, out_dir, school="", working_days=None,
    academic_year="", exam_type="", output="all",
):
    """Make selected output for one class + section. roster = [(student_id, name), ...].
    Returns the workbook path when output includes Excel."""
    if output not in {"all", "excel", "pdf"}:
        raise ValueError("output must be 'all', 'excel', or 'pdf'")
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
        for key in CLASS_SUBJECTS[cls]:
            label = SUBJECT_LABELS[key]
            fname = label.replace(" ", "_")
            headers = [f"{name}<br/>{meaning}<br/>({mx} marks)" for name, meaning, mx in SLOTS]
            build_pdf(
                os.path.join(sheets, f"{tag}_{fname}.pdf"),
                f"{label} - marks sheet",
                f"{where_pdf}<br/>Write marks as numbers. Write <b>A</b> if the student was absent. "
                f"Teacher: ____________________   Date: ____________",
                f"{tag}  |  {label}",
                roster, headers, [62] * len(SLOTS),
            )

        att_head = "Days attended" + (f"<br/>(out of {working_days})" if working_days else "")
        build_pdf(
            os.path.join(sheets, f"{tag}_Attendance.pdf"),
            "Attendance sheet",
            f"{where_pdf}<br/>Write the number of days each student attended."
            + (f" Working days: {working_days}." if working_days else ""),
            f"{tag}  |  Attendance",
            roster, [att_head], [100],
        )

    if output in {"all", "excel"}:
        workbook = os.path.join(entry, f"{tag}_entry.xlsx")
        write_workbook(workbook, tag, where, roster, CLASS_SUBJECTS[cls], working_days)
        return workbook
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("roster")
    ap.add_argument("--class", dest="cls", type=int, required=True, choices=sorted(CLASS_SUBJECTS))
    ap.add_argument("--section", required=True)
    ap.add_argument("--working-days", type=int)
    ap.add_argument("--school", default="")
    ap.add_argument("--academic-year", default="")
    ap.add_argument("--exam-type", default="")
    ap.add_argument("--out", default="out")
    a = ap.parse_args()

    roster = read_roster(a.roster)
    generate(
        roster, a.cls, a.section, a.out, a.school, a.working_days,
        academic_year=a.academic_year, exam_type=a.exam_type,
    )
    tag = f"class{a.cls}{a.section}"
    print(f"{len(roster)} students, {len(CLASS_SUBJECTS[a.cls])} subjects + attendance -> {a.out}/ (typing workbook: entry/{tag}_entry.xlsx)")


if __name__ == "__main__":
    main()
