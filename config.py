"""Shared settings for make_sheets.py and merge_marks.py.

Subject keys (left) are the portal's own names, the ones listSubjects() prints in
portal_tools.js. Labels (right) are only what gets printed on the sheets.
"""

SUBJECT_LABELS = {
    "1stLang": "First Language",
    "Eng": "English",
    "Math": "Maths",
    "ES": "EVS",
}

CLASS_SUBJECTS = {
    1: ["1stLang", "Eng", "Math"],
    2: ["1stLang", "Eng", "Math"],
    3: ["1stLang", "Eng", "Math", "ES"],
    4: ["1stLang", "Eng", "Math", "ES"],
    5: ["1stLang", "Eng", "Math", "ES"],
}

# The four slots, in the portal's order (Tool 1 .. 4): (portal name, teacher-facing name, max marks).
# ASSUMPTION: the meanings below match Tool 1..4. Check this, because teachers write marks in these columns.
# The maximum marks match the portal headings on the saved page.
SLOTS = [
    ("Tool 1", "Classroom responses", 10),
    ("Tool 2", "Homework", 10),
    ("Tool 3", "Project", 10),
    ("Tool 4", "Slip test / exam", 20),
]
