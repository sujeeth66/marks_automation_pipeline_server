import cv2, numpy as np

def order_pts(p):
    p = p.reshape(4, 2).astype("float32")
    s, d = p.sum(1), np.diff(p, axis=1).ravel()
    return np.array([p[np.argmin(s)], p[np.argmin(d)], p[np.argmax(s)], p[np.argmax(d)]], dtype="float32")  # tl,tr,br,bl

def binarize(gray):
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    return cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 31, 15)

def warp_table(img, W=1000, pad=12):
    """Find the table as the biggest connected block of ruled lines, then straighten it."""
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
    # map the table to the inside of a padded canvas so the outer border lines are not cut off
    dst = np.array([[pad, pad], [W - 1 - pad, pad], [W - 1 - pad, H - 1 - pad], [pad, H - 1 - pad]], dtype="float32")
    M = cv2.getPerspectiveTransform(pts, dst)
    return cv2.warpPerspective(img, M, (W, H), borderValue=(255, 255, 255))

def peaks(profile, thresh, gap=6):
    idx = np.where(profile > thresh)[0]
    groups, cur = [], [idx[0]]
    for i in idx[1:]:
        if i - cur[-1] <= gap: cur.append(i)
        else: groups.append(cur); cur = [i]
    groups.append(cur)
    return [int(np.mean(g)) for g in groups]

def cells(img, W=1000):
    """-> (warped image, hlines, vlines). Lines are found by long morphological opening + projection."""
    w = warp_table(img, W)
    th = binarize(cv2.cvtColor(w, cv2.COLOR_BGR2GRAY))
    H = w.shape[0]
    h = cv2.morphologyEx(th, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (W // 5, 1)))
    v = cv2.morphologyEx(th, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30)))
    hl = peaks(h.sum(axis=1) / 255, W * 0.5)
    vl = peaks(v.sum(axis=0) / 255, 0.35 * (hl[-1] - hl[0]))
    return w, hl, vl
