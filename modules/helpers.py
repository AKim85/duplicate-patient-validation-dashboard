# modules/helpers.py
import os, re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.utils.dataframe import dataframe_to_rows

def normalize_colname(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip())

def find_col(cols, candidates):
    norm_map = {normalize_colname(c).lower(): c for c in cols}
    for cand in candidates:
        key = normalize_colname(cand).lower()
        if key in norm_map:
            return norm_map[key]
    stripped_map = {re.sub(r"\s+", "", normalize_colname(c)).lower(): c for c in cols}
    for cand in candidates:
        key = re.sub(r"\s+", "", normalize_colname(cand)).lower()
        if key in stripped_map:
            return stripped_map[key]
    return None

def is_int_like(x):
    if pd.isna(x): return False
    if isinstance(x, (int, np.integer)): return True
    if isinstance(x, (float, np.floating)): return float(x).is_integer()
    if isinstance(x, str): return x.strip().isdigit()
    return False

def to_int(x):
    if isinstance(x, (int, np.integer)): return int(x)
    if isinstance(x, (float, np.floating)): return int(round(float(x)))
    if isinstance(x, str) and x.strip().isdigit(): return int(x.strip())
    return None

def interpret_match_value(v):
    if pd.isna(v): return np.nan
    if isinstance(v, (bool, np.bool_)): return bool(v)
    if isinstance(v, (int, np.integer)):
        return bool(v) if v in (0, 1) else np.nan
    if isinstance(v, (float, np.floating)):
        if float(v).is_integer() and int(v) in (0, 1):
            return bool(int(v))
        return np.nan
    s = str(v).strip().lower()
    if s in ("match", "matched", "m", "true", "t", "yes", "y", "1"):
        return True
    if s in ("mismatch", "not match", "notmatch", "mm", "false", "f", "no", "n", "0"):
        return False
    return np.nan

def valid_indicator_series(ser: pd.Series) -> bool:
    vals = ser.dropna().map(interpret_match_value)
    return vals.notna().all()

def write_df_to_sheet(ws, df: pd.DataFrame, index=False, header=True):
    ws.delete_rows(1, ws.max_row)
    for r_idx, row in enumerate(dataframe_to_rows(df, index=index, header=header), start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

def save_fig_to_png(fig, path):
    fig.savefig(path, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    return path

def contains_todelete(row) -> bool:
    for v in row.values:
        if pd.isna(v): 
            continue
        if "todelete" in str(v).lower():
            return True
    return False
``
