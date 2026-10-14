"""Shared settings for make_sheets.py and merge_marks.py.

Dynamic configuration (subjects, slots, attendance) is now scraped from the portal
and passed via API. The constants below are kept only for backward compatibility
with legacy CLI tools.
"""

# DEPRECATED: Used only for backward compatibility with legacy CLI tools.
# The API now uses dynamic subjects scraped from the portal.
SUBJECT_LABELS = {
    "1stLang": "First Language",
    "Eng": "English",
    "Math": "Maths",
    "ES": "EVS",
}

STUDENT_TAG_ID_START = 100
MAX_STUDENTS = 150
SECTION_CODES = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

# DEPRECATED: Used only for backward compatibility with legacy CLI tools.
# The API now builds dynamic tab codes from portal subjects.
TAB_CODES = ["1stLang", "Eng", "Math", "ES", "Attendance"]

# AprilTag metadata ID bases for QR code encoding
META_CLASS_BASE = 300
META_SECTION_BASE = 310
META_TAB_BASE = 350
META_EXAM_BASE = 400
EXAM_TYPES = ["FA1", "FA2", "FA3", "FA4", "SA1", "SA2"]

# DEPRECATED: Used only for backward compatibility with legacy CLI tools.
# The API now uses dynamic subjects scraped from the portal.
CLASS_SUBJECTS = {
    1: ["1stLang", "Eng", "Math"],
    2: ["1stLang", "Eng", "Math"],
    3: ["1stLang", "Eng", "Math", "ES"],
    4: ["1stLang", "Eng", "Math", "ES"],
    5: ["1stLang", "Eng", "Math", "ES"],
}

# DEPRECATED: Used only for backward compatibility with legacy CLI tools.
# The API now uses dynamic slot configurations scraped from the portal.
# The four slots, in the portal's order (Tool 1 .. 4): (portal name, teacher-facing name, max marks).
# ASSUMPTION: the meanings below match Tool 1..4. Check this, because teachers write marks in these columns.
# The maximum marks match the portal headings on the saved page.
SLOTS = [
    ("Tool 1", "Classroom responses", 10),
    ("Tool 2", "Homework", 10),
    ("Tool 3", "Project", 10),
    ("Tool 4", "Slip test / exam", 20),
]
