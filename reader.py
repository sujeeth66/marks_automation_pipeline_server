"""Read photographed marks sheets: image bytes in; for every cell a value, a confidence and a small picture out.

Used by the server's /scan endpoint. The same steps as measure.py: find the table, straighten it, cut out the
marks cells (grid.py), then recognize each one with the model (digits.py).
"""
import base64

import cv2
import numpy as np

from digits import decode, normalize, predict_proba, to_ink
from grid import cells

MAX_SIDE = 1600  # phone photos are much bigger; the model was tested on photos of about this size


class ReadError(Exception):
    """The photo could not be read (not an image, or the table could not be found)."""


def load_image(data):
    """Bytes of a JPEG/PNG -> BGR image, shrunk so its longest side is at most MAX_SIDE."""
    img = cv2.imdecode(np.frombuffer(data, np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        raise ReadError("this file is not a readable image")
    scale = MAX_SIDE / max(img.shape[:2])
    if scale < 1:
        img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    return img


def read_page(img, max_values, layers, allow_a=True, margin=6):
    """One photo of one printed page. max_values has one entry per mark column (e.g. [10, 10, 10, 20]).
    -> a list of rows, each a list of {"value", "conf", "blank", "png"} (png = base64 picture of the cell)."""
    n = len(max_values)
    try:
        warped, hl, vl = cells(img)
    except Exception as e:  # cells() fails with assorted errors when it cannot find a table
        raise ReadError("could not find the table in this photo; is the whole page in view?") from e
    if len(hl) < 3 or len(vl) < n + 1:
        raise ReadError(f"found only {len(hl)} horizontal and {len(vl)} vertical lines of the table; "
                        "take the photo flat, in good light, with the whole table in view")
    gaps = [b - a for a, b in zip(hl, hl[1:])]
    if gaps[0] > 1.4 * float(np.median(gaps)):  # the first gap is the taller header row
        hl = hl[1:]
    cols = vl[-(n + 1):]
    gray = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
    rows = []
    for r0, r1 in zip(hl, hl[1:]):
        row = []
        for j, (c0, c1) in enumerate(zip(cols, cols[1:])):
            cell = gray[r0 + margin:r1 - margin, c0 + margin:c1 - margin]
            vec, ink = normalize(to_ink(cell))
            if ink:
                label, conf = decode(predict_proba(layers, vec), max_values[j], allow_a)
            else:
                label, conf = "", 1.0
            ok, png = cv2.imencode(".png", cell)
            row.append({"value": label, "conf": round(conf, 3), "blank": not ink,
                        "png": base64.b64encode(png.tobytes()).decode("ascii")})
        rows.append(row)
    return rows
