"""
file_reader.py
--------------
Reads arbitrary Excel files (.xlsx, .xls, .xlsb).
Automatically detects the best data sheet and header row.
Returns a clean DataFrame with original Excel row numbers preserved.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Public dataclass returned by read_excel_file()
# ---------------------------------------------------------------------------

@dataclass
class ExcelReadResult:
    df: pd.DataFrame                # cleaned DataFrame with original row numbers
    sheet_name: str                 # sheet that was chosen
    header_row_index: int           # 0-based index of the header row in the sheet
    original_row_numbers: pd.Series # Excel 1-based row numbers for each data row
    all_sheets: list[str]           # all sheet names in the workbook


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _score_sheet(df_raw: pd.DataFrame) -> float:
    """Score a raw sheet DataFrame for likelihood of being a data sheet."""
    if df_raw.empty or df_raw.shape[1] < 2:
        return 0.0
    total_cells = df_raw.size
    non_null = df_raw.notna().sum().sum()
    if total_cells == 0:
        return 0.0
    density = non_null / total_cells
    # Prefer sheets with at least 3 rows and 2 columns
    row_bonus = min(df_raw.shape[0] / 100, 0.2)
    col_bonus = min(df_raw.shape[1] / 20, 0.1)
    return density + row_bonus + col_bonus


def _detect_best_sheet(sheets: dict[str, pd.DataFrame]) -> str:
    """Return the sheet name with the highest data score."""
    scores = {name: _score_sheet(df) for name, df in sheets.items()}
    return max(scores, key=scores.__getitem__)


def _score_header_row(row: pd.Series) -> float:
    """
    Score a candidate header row.
    Good headers: mostly strings, non-null, short values, high uniqueness.
    """
    vals = row.dropna().astype(str)
    if len(vals) < 2:
        return 0.0

    # How many are string-like (not purely numeric)?
    str_count = sum(1 for v in vals if not _is_numeric(v))
    str_ratio = str_count / len(vals)

    # Uniqueness
    unique_ratio = len(set(v.strip().lower() for v in vals)) / len(vals)

    # Prefer shorter average length (headers are usually short labels)
    avg_len = np.mean([len(v) for v in vals])
    length_score = max(0.0, 1.0 - avg_len / 60)

    return 0.5 * str_ratio + 0.3 * unique_ratio + 0.2 * length_score


def _is_numeric(s: str) -> bool:
    try:
        float(s)
        return True
    except ValueError:
        return False


def _detect_header_row(df_raw: pd.DataFrame, scan_rows: int = 20) -> int:
    """
    Scan the first `scan_rows` rows and return the 0-based index of the
    most likely header row.
    """
    n = min(scan_rows, len(df_raw))
    best_idx, best_score = 0, -1.0
    for i in range(n):
        score = _score_header_row(df_raw.iloc[i])
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx


def _read_raw(file_obj, engine: Optional[str] = None) -> dict[str, pd.DataFrame]:
    """Read all sheets as raw DataFrames (no header promotion)."""
    kwargs: dict = dict(header=None, dtype=str)
    if engine:
        kwargs["engine"] = engine
    return pd.read_excel(file_obj, sheet_name=None, **kwargs)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def read_excel_raw(file_obj) -> pd.DataFrame:
    """
    Lightweight raw reader for pre-check use only.

    Reads all sheets without header promotion, picks the best-scoring sheet,
    and returns its raw rows as a DataFrame of strings.  No sheet/header
    detection beyond choosing which sheet to use, no column-name inference,
    no _excel_row_ injection.

    This intentionally avoids every step that runs during read_excel_file()
    so the pre-check is independent of header mapping.
    """
    name = getattr(file_obj, "name", str(file_obj))
    ext  = name.rsplit(".", 1)[-1].lower() if "." in name else "xlsx"
    engine_map = {"xlsb": "pyxlsb", "xls": "xlrd"}
    engine = engine_map.get(ext)

    try:
        raw_sheets = _read_raw(file_obj, engine=engine)
    except Exception as exc:
        raise ValueError(f"Cannot read Excel file: {exc}") from exc

    sheet_name = _detect_best_sheet(raw_sheets)
    df_raw = raw_sheets[sheet_name].copy()

    # Normalise all cell values to stripped strings (NaN → "")
    df_raw = df_raw.fillna("").astype(str).apply(
        lambda col: col.str.strip()
    )
    # Drop completely blank rows (same as read_excel_file's dropna(how='all'))
    # Without this, large files with trailing empty rows massively inflate
    # the row comparison time.
    df_raw = df_raw[~(df_raw == "").all(axis=1)].reset_index(drop=True)
    return df_raw


def read_excel_file(
    file_obj,
    force_sheet: Optional[str] = None,
    force_header_row: Optional[int] = None,
) -> ExcelReadResult:
    """
    Read an Excel file from a file-like object or path.

    Parameters
    ----------
    file_obj          : file path or file-like object
    force_sheet       : if provided, use this sheet name instead of auto-detect
    force_header_row  : if provided, use this 0-based row index as the header

    Returns
    -------
    ExcelReadResult
    """
    # Determine file extension for engine selection
    name = getattr(file_obj, "name", str(file_obj))
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else "xlsx"

    engine_map = {"xlsb": "pyxlsb", "xls": "xlrd"}
    engine = engine_map.get(ext)  # None → openpyxl (default)

    # --- Step 1: read all sheets without header promotion ---
    try:
        raw_sheets = _read_raw(file_obj, engine=engine)
    except Exception as exc:
        raise ValueError(f"Cannot read Excel file: {exc}") from exc

    all_sheet_names = list(raw_sheets.keys())

    # --- Step 2: pick the best sheet ---
    sheet_name = force_sheet if force_sheet else _detect_best_sheet(raw_sheets)
    df_raw = raw_sheets[sheet_name]

    # --- Step 3: detect header row ---
    header_idx = force_header_row if force_header_row is not None else _detect_header_row(df_raw)

    # --- Step 4: promote header ---
    header = df_raw.iloc[header_idx].fillna("").astype(str).tolist()
    # De-duplicate header names
    seen: dict[str, int] = {}
    clean_header = []
    for h in header:
        h = h.strip()
        if h == "" or h == "nan":
            h = f"_col_{len(clean_header)}"
        if h in seen:
            seen[h] += 1
            h = f"{h}.{seen[h]}"
        else:
            seen[h] = 0
        clean_header.append(h)

    # Data rows start immediately after the header row
    data_start = header_idx + 1
    df_data = df_raw.iloc[data_start:].reset_index(drop=True).copy()
    df_data.columns = clean_header

    # --- Step 5: drop completely empty rows ---
    df_data = df_data.dropna(how="all").reset_index(drop=True)

    # --- Step 6: compute original Excel row numbers (1-based, row 1 = first row) ---
    # Excel row 1 = header label row (the sheet's row 1)
    # header_idx is 0-based → Excel row = header_idx + 1
    # data rows start at Excel row = header_idx + 2
    excel_row_start = header_idx + 2  # 1-based Excel row of first data row
    original_row_numbers = pd.Series(
        range(excel_row_start, excel_row_start + len(df_data)),
        name="_excel_row_",
    )

    # Attach to df for easy access (will be stripped before comparison)
    df_data = df_data.reset_index(drop=True)
    df_data["_excel_row_"] = original_row_numbers.values

    return ExcelReadResult(
        df=df_data,
        sheet_name=sheet_name,
        header_row_index=header_idx,
        original_row_numbers=original_row_numbers,
        all_sheets=all_sheet_names,
    )
