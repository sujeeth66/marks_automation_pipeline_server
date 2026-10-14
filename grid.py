import cv2, numpy as np

def order_pts(p):
    p = p.reshape(4, 2).astype("float32")
    s, d = p.sum(1), np.diff(p, axis=1).ravel()
    return np.array([p[np.argmin(s)], p[np.argmin(d)], p[np.argmax(s)], p[np.argmax(d)]], dtype="float32")  # tl,tr,br,bl

def binarize(gray):
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    return cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 31, 15)

def mark_lines(vl_raw, W, n_boxes):
    """Digit-box boundaries for the mark columns, left to right (n_boxes + 1 values).
    Only lines in the right half of the table are used, so the student-code and name borders are ignored.
    The right edge and the box width come from the detected lines, and every boundary is placed from the
    right edge, so faint dividers that the detector misses (such as Tool 1) are still correct."""
    found = sorted(x for x in vl_raw if x > 0.5 * W)
    if len(found) < 3:
        raise ValueError(f"found only {len(found)} mark-column lines")
    right = found[-1]
    box_w = float(np.median(np.diff(found)))
    return [int(round(right - k * box_w)) for k in range(n_boxes, -1, -1)]

def warp_table(img, W=1000, pad=12):
    """-> (table straightened to W pixels wide, M) where M maps pixels of img into the straightened table."""
    th = binarize(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
    h = cv2.morphologyEx(th, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (60, 1)))
    v = cv2.morphologyEx(th, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40)))
    grid = cv2.dilate(cv2.add(h, v), np.ones((5, 5), np.uint8))
    cnts, _ = cv2.findContours(grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    c = max(cnts, key=cv2.contourArea)
    approx = cv2.approxPolyDP(c, 0.02 * cv2.arcLength(c, True), True)
    pts = order_pts(approx if len(approx) == 4 else cv2.boxPoints(cv2.minAreaRect(c)))
    wd = (np.linalg.norm(pts[1] - pts[0]) + np.linalg.norm(pts[2] - pts[3])) / 2
    ht = (np.linalg.norm(pts[3] - pts[0]) + np.linalg.norm(pts[2] - pts[1])) / 2
    H = int(W * ht / wd)
    dst = np.array([[pad, pad], [W - 1 - pad, pad], [W - 1 - pad, H - 1 - pad], [pad, H - 1 - pad]], dtype="float32")
    M = cv2.getPerspectiveTransform(pts, dst)
    warped = cv2.warpPerspective(img, M, (W, H), borderValue=(255, 255, 255))
    return warped, M

def peaks(profile, thresh, gap=6):
    idx = np.where(profile > thresh)[0]
    groups, cur = [], [idx[0]]
    for i in idx[1:]:
        if i - cur[-1] <= gap: cur.append(i)
        else: groups.append(cur); cur = [i]
    groups.append(cur)
    return [int(np.mean(g)) for g in groups]

def fill_missing_lines(lines):
    """Insert ruled lines the detector missed, splitting any gap of about two or more rows into equal rows."""
    if len(lines) < 4:
        return lines
    step = float(np.median(np.diff(lines[1:])))
    out = [lines[0], lines[1]]
    for a, b in zip(lines[1:], lines[2:]):
        n = max(1, int(round((b - a) / step)))
        for k in range(1, n):
            out.append(int(round(a + (b - a) * k / n)))
        out.append(b)
    return out

def cells_with_map(img, W=1000):
    """-> (warped table image, hlines, vlines, M). M maps pixels of img into the warped table."""
    w, M = warp_table(img, W)
    th = binarize(cv2.cvtColor(w, cv2.COLOR_BGR2GRAY))
    h = cv2.morphologyEx(th, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (W // 5, 1)))
    v = cv2.morphologyEx(th, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30)))
    hl = fill_missing_lines(peaks(h.sum(axis=1) / 255, W * 0.25))
    vl = peaks(v.sum(axis=0) / 255, 0.35 * (hl[-1] - hl[0]))
    return w, hl, vl, M

def cells(img, W=1000):
    """-> (warped image, hlines, vlines)."""
    w, hl, vl, _ = cells_with_map(img, W)
    return w, hl, vl