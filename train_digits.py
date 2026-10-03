#!/usr/bin/env python3
"""Train the cell recognizer on public handwriting (MNIST) plus synthetic cells, and export model.json.

Usage:
  python train_digits.py --out ../webapp/model/model.json [--mnist mnist.npz]

No photo of any teacher's writing is used for training. To measure it on real sheets, use
  python measure.py --recognizer mlp --model model.json --photos ... --xlsx ... --tab ...
MNIST: pass an .npz with xtr,ytr,xva,yva (float images 0..1, 784 values), or leave --mnist out to
download it with scikit-learn (fetch_openml, about 70 MB, needs internet).
Needs: numpy, opencv-python, scikit-learn, pillow. Takes a few minutes on a laptop CPU.
"""
import argparse
import time

import numpy as np
from sklearn.neural_network import MLPClassifier

from digits import export_model, make_dataset


def load_mnist(path):
    if path:
        d = np.load(path)
        return d["xtr"], d["ytr"], d["xva"], d["yva"]
    from sklearn.datasets import fetch_openml
    x, y = fetch_openml("mnist_784", version=1, return_X_y=True, as_frame=False)
    x, y = x.astype("float32") / 255, y.astype(int)
    return x[:60000], y[:60000], x[60000:], y[60000:]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mnist")
    ap.add_argument("--out", default="model.json")
    ap.add_argument("--epochs", type=int, default=40)
    a = ap.parse_args()

    xtr, ytr, xva, yva = load_mnist(a.mnist)
    t = time.time()
    X, y = make_dataset(xtr, ytr, seed=1)
    Xv, yv = make_dataset(xva, yva, seed=2, n_small=100, n_big=60, n_a=150)
    print(f"made {len(X)} training and {len(Xv)} validation cells in {time.time() - t:.0f}s")

    clf = MLPClassifier(hidden_layer_sizes=(256, 128), batch_size=256, learning_rate_init=1e-3, alpha=1e-4,
                        max_iter=a.epochs, random_state=0)
    t = time.time()
    clf.fit(X, y)
    print(f"trained in {time.time() - t:.0f}s, accuracy on synthetic validation cells: {clf.score(Xv, yv):.3f}")
    export_model(clf, a.out)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
