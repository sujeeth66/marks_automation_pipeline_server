# Marks pipeline

Per class and section. Needs Python 3 and `pip install reportlab`.

1. **Roster.** Log in to the portal, pick the class and section, open the browser console, paste
   `portal_tools.js`, run `scrapeRoster()`, and paste the clipboard into `roster.csv`.
2. **Sheets.** `python make_sheets.py roster.csv --class 3 --section A --working-days 37`
   Print `out/sheets/*.pdf` (one per subject, plus attendance) and give each teacher their sheet.
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
workbook out), `POST /merge` (filled workbook in, data for the portal out) and `POST /scan` (photos in, cell values
and pictures out, for a review screen). See `DEPLOY.md` to run it on Render, and `python -m unittest discover -s tests -v`
to test it.
Done