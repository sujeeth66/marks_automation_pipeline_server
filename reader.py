"""Read photographed marks sheets: image bytes in; for every cell a value, a confidence and a small picture out.

Used by the server's /submit and /scan endpoints. Every sheet carries AprilTags: four corners (page
straightening), one tag per student row (identity), and class / section / tab tags (metadata).
"""
import base64

import cv2
import numpy as np
try:
    import pyapriltags
except ImportError:
    pyapriltags = None

from config import (
    EXAM_TYPES, META_CLASS_BASE, META_EXAM_BASE, META_SECTION_BASE, META_TAB_BASE, SECTION_CODES, STUDENT_TAG_ID_START,
)
from digits import decode, normalize, predict_proba, to_ink
from grid import cells, cells_with_map, mark_lines

MAX_SIDE = 1600  # phone photos are much bigger; the model was tested on photos of about this size
TAG_FAMILY = "tagStandard41h12"  # must match make_sheets.py
CORNER_TAG_IDS = [0, 1, 2, 3]    # TL, TR, BR, BL
A4_WIDTH_PT, A4_HEIGHT_PT = 595.2756, 841.8898
CORNER_CENTER_PT = 38            # must match make_sheets.py
RECTIFIED_WIDTH = 2200

_detector = None


class ReadError(Exception):
    """The photo could not be read (not an image, or the table could not be found)."""


def load_image(data, preserve_resolution=False):
    """Decode image bytes to BGR. Keep full resolution when the markers must be read from the original."""
    img = cv2.imdecode(np.frombuffer(data, np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        raise ReadError("this file is not a readable image")
    if not preserve_resolution:
        scale = MAX_SIDE / max(img.shape[:2])
        if scale < 1:
            img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    return img


def _get_detector():
    global _detector
    if pyapriltags is None:
        raise ReadError("AprilTag detection requires pyapriltags; install the packages in requirements.txt")
    if _detector is None:
        _detector = pyapriltags.Detector(TAG_FAMILY)
    return _detector


def _detect_apriltags(img):
    """-> [(tag_id, corners)] where corners is a 4x2 array in detector order."""
    gray = img if img.ndim == 2 else cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    results = _get_detector().detect(gray)
    return [(r.tag_id, np.asarray(r.corners, dtype="float32").reshape(-1, 2)) for r in results]


def _corner_positions(apriltags):
    """-> the four corner centres as an array in TL, TR, BR, BL order."""
    positions = {}
    for tag_id, corners in apriltags:
        if tag_id in CORNER_TAG_IDS:
            if tag_id in positions:
                raise ReadError(f"found duplicate corner AprilTag ID {tag_id}")
            positions[tag_id] = corners.mean(axis=0)
    missing = [tag_id for tag_id in CORNER_TAG_IDS if tag_id not in positions]
    if missing:
        raise ReadError(
            f"missing corner AprilTags with IDs {missing}; keep all four printed corner codes visible in the photo"
        )
    return [positions[i] for i in CORNER_TAG_IDS]


def straighten_page(img):
    """Use the four corner AprilTags to rectify the page to an A4 portrait canvas."""
    positions = _corner_positions(_detect_apriltags(img))

    width = RECTIFIED_WIDTH
    height = int(round(width * A4_HEIGHT_PT / A4_WIDTH_PT))
    x = CORNER_CENTER_PT * width / A4_WIDTH_PT
    y = CORNER_CENTER_PT * height / A4_HEIGHT_PT
    source = np.asarray(positions, dtype="float32")
    target = np.asarray([[x, y], [width - x, y], [width - x, height - y], [x, height - y]], dtype="float32")
    matrix = cv2.getPerspectiveTransform(source, target)
    return cv2.warpPerspective(img, matrix, (width, height), borderValue=(255, 255, 255))


def _single(tags, low, high, what, allow_missing=False):
    """The one tag in [low, high) -> its offset from low. Anything other than exactly one is an error."""
    found = [t - low for t in tags if low <= t < high]
    if len(found) != 1:
        if allow_missing and len(found) == 0:
            return None
        raise ReadError(f"expected exactly one {what} tag on each sheet, found {len(found)}")
    return found[0]


def read_sheet_metadata(img, tab_codes=None):
    """Read class, section, subject tab and exam type from the metadata tags at the bottom of the page."""
    tags = [tid for tid, _ in _detect_apriltags(img)]
    if not set(CORNER_TAG_IDS).issubset(tags):
        raise ReadError("expected all four corner AprilTags; keep the whole page in view")
    cls = _single(tags, META_CLASS_BASE, META_SECTION_BASE, "class")
    section = SECTION_CODES[_single(tags, META_SECTION_BASE, META_TAB_BASE, "section")]
    
    # Use provided tab_codes or fallback to hardcoded
    codes = tab_codes if tab_codes else []
    tab = codes[_single(tags, META_TAB_BASE, META_TAB_BASE + len(codes), "subject")] if codes else None
    
    # Decode exam type if present
    exam_type = None
    exam_index = _single(tags, META_EXAM_BASE, META_EXAM_BASE + len(EXAM_TYPES), "exam_type", allow_missing=True)
    if exam_index is not None and 0 <= exam_index < len(EXAM_TYPES):
        exam_type = EXAM_TYPES[exam_index]
    
    return {"version": 1, "class": cls, "section": section, "tab": tab, "exam_type": exam_type} 



def _read_digit_field(gray, bounds, maximum, layers, margin):
    """Recognize a right-aligned, one-character-per-box number or an empty entry (A)."""
    left, right, top, bottom = bounds
    count = len(str(maximum))
    edges = np.linspace(left, right, count + 1).astype(int)
    digits, confidences = [], []
    for x0, x1 in zip(edges, edges[1:]):
        box = gray[top + margin:bottom - margin, x0 + margin:x1 - margin]
        vec, ink = normalize(to_ink(box))
        if not ink:
            digits.append("")
            continue
        digit, confidence = decode(predict_proba(layers, vec), 9, allow_a=False)
        digits.append(digit)
        confidences.append(confidence)

    if not confidences:
        return "A", 1.0, True

    first_digit = next(i for i, digit in enumerate(digits) if digit)
    if any(not digit for digit in digits[first_digit:]):
        value = ""
    else:
        value = "".join(digits[first_digit:])
    confidence = min(confidences)
    if value and int(value) > maximum:
        value = ""
    return value, confidence, False

def _rows_from_student_tags(straight, M, roster, table_height):
    """Find each row from its student tag. The tag ID is the student's roster position, so each row is
    identified directly. -> (row_bounds, roster_positions, missing_indices), in printed order."""
    found = {}
    for tid, corners in _detect_apriltags(straight):
        if not STUDENT_TAG_ID_START <= tid < META_CLASS_BASE:
            continue
        idx = tid - STUDENT_TAG_ID_START
        if idx >= len(roster):
            raise ReadError(
                f"the sheet has a student code beyond the {len(roster)} students in the roster; "
                "check the roster matches the sheet"
            )
        if idx in found:
            raise ReadError(f"the student code for {roster[idx][1]} appears twice on this photo")
        pts = cv2.perspectiveTransform(corners.reshape(-1, 1, 2).astype("float32"), M).reshape(-1, 2)
        found[idx] = float(pts[:, 1].mean())

    if len(found) < 2:
        raise ReadError("found fewer than two student codes; take the photo with the whole table in view")
    lo, hi = min(found), max(found)
    gaps = [i for i in range(lo, hi + 1) if i not in found]
    if gaps:
        if len(gaps) > 4:
            i = gaps[0]
            raise ReadError(
                f"could not read the student code for {len(gaps)} row(s) on this photo, for example "
                f"{roster[i][0]} {roster[i][1]}. Take clearer photos."
            )
        # If <= 4 gaps, interpolate missing row positions
        ids = list(range(lo, hi + 1))
        ys = []
        for i in ids:
            if i in found:
                ys.append(found[i])
            else:
                # Interpolate based on neighbors
                prev_found = max([x for x in found.keys() if x < i], default=None)
                next_found = min([x for x in found.keys() if x > i], default=None)
                if prev_found is not None and next_found is not None:
                    # Linear interpolation
                    prev_y = found[prev_found]
                    next_y = found[next_found]
                    ratio = (i - prev_found) / (next_found - prev_found)
                    ys.append(prev_y + ratio * (next_y - prev_y))
                elif prev_found is not None:
                    # Extrapolate using the last known pitch
                    ys.append(found[prev_found] + (i - prev_found))
                elif next_found is not None:
                    # Extrapolate backwards
                    ys.append(found[next_found] - (next_found - i))
                else:
                    # Should not happen
                    ys.append((found[lo] + found[hi]) / 2)
    else:
        ids = list(range(lo, hi + 1))
        ys = [found[i] for i in ids]

    pitch = float(np.median(np.diff(ys)))
    bounds = [(int(round(y - pitch / 2)), int(round(y + pitch / 2))) for y in ys]
    if bounds[0][0] < 0 or bounds[-1][1] > table_height:
        raise ReadError("part of the table is cut off in this photo; take the photo with the whole table in view")
    return bounds, ids, gaps

def read_page(img, max_values, layers, allow_a=True, margin=6, identify=False, roster=None, tab_codes=None):
    """One photo of one printed page. max_values has one entry per mark column (e.g. [10, 10, 10, 20]).
    identify=True needs the roster in printed order; each row is then matched to its student code.
    -> rows of cells with value, confidence, blank state and a PNG. With identify, each row also has
    student_id and roster_index."""
    n = len(max_values)
    if identify:
        if roster is None:
            raise ReadError("identify needs the roster")
        img = straighten_page(img)
        try:
            warped, hl, vl, M = cells_with_map(img)
        except Exception as e:
            raise ReadError("could not find the table in this photo; is the whole page in view?") from e
        row_bounds, row_ids, missing_indices = _rows_from_student_tags(img, M, roster, warped.shape[0])
    else:
        try:
            warped, hl, vl = cells(img)
        except Exception as e:
            raise ReadError("could not find the table in this photo; is the whole page in view?") from e
        if len(hl) < 3:
            raise ReadError(f"found only {len(hl)} horizontal lines of the table; "
                            "take the photo flat, in good light, with the whole table in view")
        gaps = [b - a for a, b in zip(hl, hl[1:])]
        if gaps[0] > 1.4 * float(np.median(gaps)):  # the first gap is the taller header row
            hl = hl[1:]
        row_bounds = list(zip(hl, hl[1:]))
        row_ids = [None] * len(row_bounds)
        missing_indices = []

    box_counts = [len(str(maximum)) for maximum in max_values]
    try:
        digit_cols = mark_lines(vl, warped.shape[1], sum(box_counts))
    except ValueError as e:
        raise ReadError(f"{e}; take the photo flat, in good light, with the whole table in view") from e
    slot_cols = [digit_cols[sum(box_counts[:j])] for j in range(n + 1)]

    gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
    rows = []
    for row_index, (r0, r1) in enumerate(row_bounds):
        row = []
        for j, (c0, c1) in enumerate(zip(slot_cols, slot_cols[1:])):
            label, conf, blank = _read_digit_field(gray, (c0, c1, r0, r1), max_values[j], layers, margin)
            cell = gray[r0 + margin:r1 - margin, c0 + margin:c1 - margin]
            _, png = cv2.imencode(".png", cell)
            row.append({"value": label, "conf": round(conf, 3), "blank": blank,
                        "png": base64.b64encode(png.tobytes()).decode("ascii")})
        if identify:
            pos = row_ids[row_index]
            qr_error = "interpolated" if pos in missing_indices else ""
            rows.append({"student_id": roster[pos][0], "roster_index": pos, "qr_error": qr_error, "cells": row})
        else:
            rows.append(row)
    return rows
