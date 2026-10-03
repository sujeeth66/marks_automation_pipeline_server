"""Cell recognizer: preprocessing, synthetic training data, and a small MLP.

Classes: the numbers 0..40 and "A" (absent), 42 in all. A cell image goes through normalize() (the same
function is used for real crops and for synthetic ones), then a small MLP predicts one of the classes.
Which classes are allowed for a given column (e.g. 0..10 or 0..20, plus A) is applied afterwards, in decode().
The web app's js/vision.js re-implements to_ink() and normalize() and must stay in step with them.
"""
import glob, os, random
import cv2, numpy as np
from PIL import Image, ImageDraw, ImageFont

CLASSES = [str(i) for i in range(41)] + ["A"]          # index -> label
A_INDEX = len(CLASSES) - 1
CW, CH = 40, 20                                         # model input: 20 rows x 40 cols
FONTS = glob.glob("/usr/share/fonts/truetype/dejavu/DejaVu*.ttf")  # used to draw the letter A (no real handwritten A in MNIST)


# ---------- preprocessing (shared by real and synthetic cells) ----------
def to_ink(gray):
    """Real crop (dark ink on light paper, uint8) -> ink-bright float image 0..1."""
    g = gray.astype("float32")
    bg, dark = np.percentile(g, 85), np.percentile(g, 1)
    if bg - dark < 25:                                   # almost no contrast: treat as empty
        return np.zeros_like(g)
    return np.clip((bg - g) / (bg - dark), 0, 1)


def normalize(ink):
    """ink-bright float image of any size -> (CH x CW float32 vector 0..1, ink_amount).
    Removes ruled-line fragments, crops to the writing, scales to fit, centres."""
    h, w = ink.shape
    binary = (ink > 0.45).astype("uint8")
    n, lab, st, _ = cv2.connectedComponentsWithStats(binary, connectivity=8)
    keep = np.zeros_like(binary)
    for k in range(1, n):
        x, y, cw, chh, area = st[k]
        if area < 10:                                    # specks
            continue
        if chh <= 6 and cw >= 0.7 * w:                   # horizontal ruled line
            continue
        if cw <= 6 and chh >= 0.8 * h:                   # vertical ruled line
            continue
        keep[lab == k] = 1
    ys, xs = np.nonzero(keep)
    if len(xs) < 15:
        return np.zeros(CH * CW, "float32"), 0
    x0, x1, y0, y1 = xs.min(), xs.max() + 1, ys.min(), ys.max() + 1
    mask = cv2.dilate(keep, np.ones((3, 3), "uint8"))
    crop = (ink * mask)[y0:y1, x0:x1]
    bh, bw = crop.shape
    s = min((CH - 4) / bh, (CW - 4) / bw)
    nh, nw = max(1, int(round(bh * s))), max(1, int(round(bw * s)))
    small = cv2.resize(crop, (nw, nh), interpolation=cv2.INTER_AREA)
    out = np.zeros((CH, CW), "float32")
    oy, ox = (CH - nh) // 2, (CW - nw) // 2
    out[oy:oy + nh, ox:ox + nw] = small
    m = out.max()
    return (out / m if m > 0 else out).ravel(), int(keep.sum())


# ---------- synthetic data ----------
def tight(d28):
    ys, xs = np.nonzero(d28 > 0.25)
    return d28[ys.min():ys.max() + 1, xs.min():xs.max() + 1]


def paste_digits(digs, rng, H=44, W=96):
    """Place digit images (tight crops) side by side on a canvas, with jitter."""
    canvas = np.zeros((H, W), "float32")
    hh = rng.randint(22, 34)
    parts, tw = [], 0
    for d in digs:
        ar = d.shape[1] / d.shape[0]
        wd = max(4, int(hh * ar * rng.uniform(0.85, 1.25)))
        parts.append(cv2.resize(d, (wd, hh), interpolation=cv2.INTER_AREA)); tw += wd
    gap = rng.randint(1, 7)
    tw += gap * (len(parts) - 1)
    x = max(2, (W - tw) // 2 + rng.randint(-4, 4))
    for p in parts:
        y = max(1, (H - hh) // 2 + rng.randint(-3, 3))
        w = min(p.shape[1], W - x)
        if w > 0:
            canvas[y:y + hh, x:x + w] = np.maximum(canvas[y:y + hh, x:x + w], p[:, :w])
        x += p.shape[1] + gap
    return canvas


def render_A(rng, H=44, W=96):
    f = ImageFont.truetype(rng.choice(FONTS), rng.randint(28, 40))
    im = Image.new("L", (W, H), 0)
    ImageDraw.Draw(im).text((W // 2, H // 2), "A", fill=255, font=f, anchor="mm")
    return np.asarray(im, "float32") / 255


def augment(img, rng, nprng):
    H, W = img.shape
    k = rng.choice([0, 0, 1, 1, 2])                                # stroke thickness
    if k:
        img = cv2.dilate(img, np.ones((2, 2), "uint8"), iterations=k)
    ang, sh = rng.uniform(-12, 12), rng.uniform(-0.25, 0.25)       # rotation, shear, scale
    sx, sy = rng.uniform(0.8, 1.2), rng.uniform(0.85, 1.15)
    img = cv2.warpAffine(img, cv2.getRotationMatrix2D((W / 2, H / 2), ang, 1.0), (W, H))
    img = cv2.warpAffine(img, np.array([[sx, sh, (1 - sx) * W / 2 - sh * H / 2], [0, sy, (1 - sy) * H / 2]], "float32"), (W, H))
    if rng.random() < 0.6:                                         # mild elastic distortion
        dx = cv2.GaussianBlur(nprng.uniform(-1, 1, (H, W)).astype("float32"), (0, 0), 6) * rng.uniform(8, 22)
        dy = cv2.GaussianBlur(nprng.uniform(-1, 1, (H, W)).astype("float32"), (0, 0), 6) * rng.uniform(8, 22)
        gx, gy = np.meshgrid(np.arange(W, dtype="float32"), np.arange(H, dtype="float32"))
        img = cv2.remap(img, gx + dx, gy + dy, cv2.INTER_LINEAR)
    for _ in range(rng.choice([0, 1, 1, 2, 3])):                   # fragments of the cell's ruled lines
        v, t = rng.uniform(0.5, 1.0), rng.choice([1, 2, 2, 3])
        if rng.random() < 0.6:
            y = rng.choice([rng.randint(0, 3), H - 1 - rng.randint(0, 3)])
            x0 = rng.randint(0, 8) if rng.random() < 0.3 else 0
            x1 = W - (rng.randint(0, 8) if rng.random() < 0.3 else 0)
            img[y:y + t, x0:x1] = v
        else:
            x = rng.choice([rng.randint(0, 3), W - 1 - rng.randint(0, 3)])
            img[:, x:x + t] = v
    sig = rng.uniform(0, 1.0)
    if sig > 0.3:
        img = cv2.GaussianBlur(img, (0, 0), sig)
    img = img * rng.uniform(0.7, 1.0) + nprng.normal(0, rng.uniform(0, 0.06), img.shape).astype("float32")
    return np.clip(img, 0, 1)


def make_dataset(mnist_x, mnist_y, seed=0, n_small=2500, n_big=1500, n_a=3000, with_aug=True):
    """n_small samples for each of 0..20, n_big for 21..40, n_a for A."""
    rng, nprng = random.Random(seed), np.random.RandomState(seed)
    by_digit = {d: [tight(mnist_x[i].reshape(28, 28)) for i in np.nonzero(mnist_y == d)[0][:3000]] for d in range(10)}
    X, y = [], []
    for ci, label in enumerate(CLASSES):
        n = n_a if label == "A" else (n_small if int(label) <= 20 else n_big)
        for _ in range(n):
            if label == "A":
                img = render_A(rng)
            else:
                img = paste_digits([rng.choice(by_digit[int(ch)]) for ch in label], rng)
            if with_aug:
                img = augment(img, rng, nprng)
            vec, ink = normalize(img)
            if ink:
                X.append(vec); y.append(ci)
    return np.array(X, "float32"), np.array(y)


# ---------- model: export, load, predict ----------
def export_model(clf, path):
    """Write the MLP weights as JSON, for both measure.py and the web app (js/recognizer.js)."""
    import json
    layers = [{"in": int(w.shape[0]), "out": int(w.shape[1]),
               "w": np.round(w, 4).ravel().tolist(), "b": np.round(b, 4).tolist()}
              for w, b in zip(clf.coefs_, clf.intercepts_)]
    json.dump({"classes": CLASSES, "input": [CH, CW], "layers": layers}, open(path, "w"), separators=(",", ":"))


def load_model(path):
    import json
    m = json.load(open(path))
    return [(np.array(l["w"], "float32").reshape(l["in"], l["out"]), np.array(l["b"], "float32")) for l in m["layers"]]


def predict_proba(layers, vec):
    x = vec
    for i, (w, b) in enumerate(layers):
        x = x @ w + b
        if i < len(layers) - 1:
            x = np.maximum(x, 0)
    e = np.exp(x - x.max())
    return e / e.sum()


def decode(proba, max_value, allow_a=True):
    """Keep only the classes 0..max_value (and A when allow_a), renormalise, -> (label, confidence)."""
    mask = np.zeros(len(CLASSES), "float32")
    mask[:max_value + 1] = 1
    if allow_a:
        mask[A_INDEX] = 1
    p = proba * mask
    total = p.sum()
    p = p / total if total > 1e-9 else mask / mask.sum()
    k = int(p.argmax())
    return CLASSES[k], float(p[k])
