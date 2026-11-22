# Land-Use Classification of Satellite Images (EuroSAT)

CPS 4841 Computer Vision — course project · Qiwen Xue · Fall 2026

## Problem

Classify 64 × 64 RGB Sentinel-2 satellite patches into one of 10 land-use / land-cover classes
(AnnualCrop, Forest, HerbaceousVegetation, Highway, Industrial, Pasture, PermanentCrop, Residential,
River, SeaLake). Three genuinely different methods are compared on the same fixed test split with
accuracy and macro-averaged precision, recall and F1-score.

## Dataset

- **EuroSAT (RGB)** — Helber et al., IEEE JSTARS 2019. 27,000 images, 10 classes, MIT license.
- Download: <https://zenodo.org/records/7711810> (`EuroSAT_RGB.zip`, ~90 MB). Project page: <https://github.com/phelber/EuroSAT>
- The data is **not** stored in this repository. `python download_data.py` downloads it into `data/`.
- Split: stratified 70 / 15 / 15 (train 18,900 / val 4,050 / test 4,050), seed 42, saved in [`splits.csv`](splits.csv).

## Repository layout

| Path | What it is |
|---|---|
| `01_dataset_check.ipynb` | Downloads the data, class counts and plots, sample images, annotation checks, builds the split |
| `02_method1_hog_color_svm.ipynb` | **Method 1** — HOG + LBP + colour features with an RBF SVM (traditional baseline) |
| `common.py` | Shared helpers: download, fixed split, metric function |
| `download_data.py` | One-line data download |
| `results/` | Saved outputs of every method (metrics JSON/CSV, classification reports, confusion matrices) |
| `figures/` | Dataset figures used in the reports |
| `reports/` | Scripts that build the Word/PDF reports (the PDFs themselves are submitted on Canvas) |

## How to run

Locally (Python 3.10+):

```bash
pip install numpy pandas pillow matplotlib scikit-learn scikit-image jupyter
python download_data.py
jupyter nbconvert --to notebook --execute --inplace 02_method1_hog_color_svm.ipynb
```

In Google Colab: open a notebook, run `!git clone https://github.com/xueqw/eurosat-landuse-classification.git` and `%cd eurosat-landuse-classification`, then *Run all*.
Every notebook downloads the data itself if `data/` is missing. All notebooks are committed **with their
executed outputs**.

## Results (test split, 4,050 images)

| Method | Accuracy | Macro Precision | Macro Recall | Macro F1 |
|---|---|---|---|---|
| Method 1: HOG + LBP + colour features, RBF SVM (C = 3) | 85.53 % | 85.10 % | 84.84 % | 84.90 % |
| Method 2: small CNN from scratch | — | — | — | — |
| Method 3: ResNet-18 transfer learning | — | — | — | — |

Per-class results and the confusion matrix are in [`results/`](results/).

Method 2 (small CNN from scratch) and Method 3 (ResNet-18 transfer learning) will be added for the final report.

## Source and AI Disclosure

| Source / tool | Used for | What I changed or added |
|---|---|---|
| EuroSAT dataset and paper (Helber et al. 2019) | Data, class definitions | Data used as-is; added my own stratified split |
| HOG — Dalal & Triggs, CVPR 2005; LBP — Ojala et al., TPAMI 2002 (via scikit-image) | Method 1 features | Chose the feature combination and parameters |
| scikit-learn `SVC`, `StandardScaler`, metrics | Method 1 classifier and evaluation | Chose RBF kernel; tuned `C` on the validation split |
| Claude (Anthropic), model Claude Opus 5.5, via Claude Code | Suggested the dataset; drafted `common.py`, the notebooks and the report text | Reviewed the code and results; decisions and final edits are mine |

Places that use AI-drafted code are marked with comments in the code. Details, including a case where
the AI's output was wrong and how it was fixed, are in the midterm report.
