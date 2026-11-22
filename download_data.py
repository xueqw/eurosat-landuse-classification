"""Download EuroSAT RGB into data/ and write splits.csv if it is missing.

Usage:  python download_data.py
"""
from common import load_splits

if __name__ == "__main__":
    df = load_splits()
    print(df["split"].value_counts().to_string())
