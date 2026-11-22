"""Shared helpers used by every method notebook: download EuroSAT, load the fixed split, compute metrics.

Source and AI disclosure: this file was drafted with Claude (Anthropic, model Claude Opus 5.5) and
reviewed by Qiwen Xue. The split logic reproduces the one in 01_dataset_check.ipynb exactly.
"""
import glob
import os
import urllib.request
import zipfile

import numpy as np
import pandas as pd
from sklearn.metrics import accuracy_score, precision_recall_fscore_support

URL = "https://zenodo.org/records/7711810/files/EuroSAT_RGB.zip"  # official EuroSAT RGB release (MIT)
DATA_DIR = "data"
ROOT = os.path.join(DATA_DIR, "EuroSAT_RGB")
SPLITS_CSV = "splits.csv"
CLASSES = ["AnnualCrop", "Forest", "HerbaceousVegetation", "Highway", "Industrial",
           "Pasture", "PermanentCrop", "Residential", "River", "SeaLake"]


def download_eurosat():
    """Download and unzip EuroSAT RGB (about 90 MB) into data/ unless it is already there."""
    if os.path.isdir(ROOT):
        return ROOT
    os.makedirs(DATA_DIR, exist_ok=True)
    zip_path = os.path.join(DATA_DIR, "EuroSAT_RGB.zip")
    print("Downloading", URL)
    urllib.request.urlretrieve(URL, zip_path)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(DATA_DIR)
    return ROOT


def make_splits(seed=42):
    """Stratified 70/15/15 split, identical to the one built in 01_dataset_check.ipynb."""
    records = []
    for label_id, cls in enumerate(CLASSES):
        for path in sorted(glob.glob(os.path.join(ROOT, cls, "*.jpg"))):
            records.append({"path": path, "label": cls, "label_id": label_id})
    df = pd.DataFrame(records)
    rng = np.random.default_rng(seed)
    split = np.empty(len(df), dtype=object)
    for cls in CLASSES:
        idx = rng.permutation(df.index[df["label"] == cls].to_numpy())
        n_tr, n_va = int(0.70 * len(idx)), int(0.15 * len(idx))
        split[idx[:n_tr]] = "train"
        split[idx[n_tr:n_tr + n_va]] = "val"
        split[idx[n_tr + n_va:]] = "test"
    df["split"] = split
    return df


def load_splits():
    """Return the fixed split as a DataFrame (path, label, label_id, split)."""
    download_eurosat()
    if os.path.exists(SPLITS_CSV):
        return pd.read_csv(SPLITS_CSV)
    df = make_splits()
    df.to_csv(SPLITS_CSV, index=False)
    return df


def compute_metrics(y_true, y_pred):
    """The project's evaluation metrics: accuracy and macro-averaged precision, recall and F1."""
    p, r, f1, _ = precision_recall_fscore_support(y_true, y_pred, average="macro", zero_division=0)
    return {"accuracy": accuracy_score(y_true, y_pred),
            "macro_precision": p, "macro_recall": r, "macro_f1": f1}
