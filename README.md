# Marks pipeline

Per class and section. Needs Python 3 and `pip install reportlab`.

1. **Roster.** Log in to the portal, pick the class and section, open the browser console, paste
   `portal_tools.js`, run `scrapeRoster()`, and paste the clipboard into `roster.csv`.
2. **Sheets.** `python make_sheets.py roster.csv --class 3 --section A --working-days 37`
   Print `out/sheets/*.pdf` (one per subject, plus attendance) and give each teacher their sheet.
   For each score, write one digit per small box, starting at the right. Leave all boxes empty for `A` (absent);
   blank attendance entries are also `A`.
3. **Typing.** Open `out/entry/class3A_entry.xlsx` in Excel and type from the filled-in sheets, one tab per
   subject plus Attendance. Only the yellow boxes are unlocked, so Tab jumps between them. Use `A` for absent.
4. **Merge.** `python merge_marks.py --class 3 --section A --working-days 37`
   Fix anything it lists; it writes `data_class3A.json` only when everything is valid.
5. **Fill the portal.** On the same portal page, in the console: `var data = <contents of the json>`,
   then `fillMarks(data)` (dry run), then `fillMarks(data, {dryRun:false})`.
   Review the page, use **Save&Edit**, and only then Final Submit (it cannot be edited afterwards).

Settings (subjects per class, the four tools and their max marks) are in `config.py`.
Never commit real student data: the `.gitignore` here excludes the roster, entry workbooks and output.

## Server (for the browser extension)
`app.py` is a small Flask server that wraps the same code: `POST /sheets` (roster in, zip of PDF sheets and the typing
workbook out), `POST /merge` (filled workbook in, data for the portal out), `POST /scan` (one tab's photos in, cell
values and pictures out) and `POST /submit` (one or more tabs' photos in, a compiled partial `fillMarks()` draft plus per-cell
review evidence out). See `DEPLOY.md` to run it on Render, and `python -m unittest discover -s tests -v` to test it.

The generated PDFs include a metadata QR in the header on every page, a QR code for each student row, and a distinct
QR fiducial in each page corner. The compact row layout fits roughly 40% more students per page while retaining
separate digit boxes. `/submit` can accept photos in a generic repeated `photos` field: it reads the metadata QR to
route each page to its subject, checks the requested class/section, then uses the corner codes to straighten the
page and student QRs to associate rows with roster IDs. QR detection uses `zxing-cpp` against the original uploaded
photo resolution; the image is reduced only after the page has been rectified for mark reading. Include the full
page and all QR codes. Or use named `photos_<subject key>` fields to explicitly choose tabs. Do not mix generic
and named uploads in one request.

Send multipart fields `class`, `roster` (JSON array of `{ "id", "name" }` objects), optional `section`,
`working_days`, and either generic `photos` or named subject fields; `photos_Attendance` is optional for named
uploads. Each tab accepts up to 8 page photos. Photos need not be in roster order, but each roster student must
appear exactly once per submitted tab. An unreadable, duplicate, unknown, or missing QR makes the request fail
with a list of identity issues; the server never falls back to row-order matching.

Mark fields and attendance use one box per digit, right-aligned. If every digit box for an entry is empty, the
server records `A` (absent); partial or malformed digits are flagged for review. On success, `data` uses the portal
`fillMarks()` format, `scans` maps rows to verified student IDs, and `review` lists invalid or below-0.90-confidence
cells. The `tabs` list and each student's `marks` keys contain only the submitted tabs, and `attendance` is included
only when attendance photos were submitted. This makes single-subject testing safe: unsubmitted portal fields are
not present in the payload and must remain untouched by a caller. Empty fields within a scanned sheet still become
`A`. This is a draft for the extension to fill into the portal; the teacher reviews/highlights there before final
submission.
Done