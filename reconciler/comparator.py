"""
comparator.py
-------------
Cell-by-cell comparison of two aligned DataFrames.

Key features:
  - Vectorized numpy/pandas operations — no row-by-row Python loops
  - Pluggable numeric tolerance: diff <= max(abs_tol, rel_tol * max(|a|, |b|))
  - Normalizes text (strip/lower), dates (to common format), and numeric strings
  - Reports exact Excel cell addresses (A1 notation)
  - Preserves original Excel row numbers in all output
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional
import datetime as dt_mod

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Public dataclasses
# ---------------------------------------------------------------------------

@dataclass
class CellDifference:
    row_key: str
    client_excel_row: Optional[int]
    output_excel_row: Optional[int]
    column_name: str             # client column name
    client_value: object
    output_value: object
    numeric_diff: Optional[float]
    diff_pct: Optional[float]
    issue: str                   # "MISMATCH" | "MISSING_CLIENT" | "MISSING_OUTPUT"


@dataclass
class RowResult:
    row_key: str
    client_excel_row: Optional[int]
    output_excel_row: Optional[int]
    status: str                  # "MATCH" | "MISMATCH" | "MISSING" | "EXTRA"
    mismatch_columns: list[str] = field(default_factory=list)


@dataclass
class ComparisonSummary:
    total_client_rows: int
    total_output_rows: int
    matched_rows: int
    mismatched_rows: int
    missing_rows: int            # in client, not in output
    extra_rows: int              # in output, not in client
    total_cells_compared: int
    matched_cells: int
    mismatched_cells: int
    missing_cell_values: int     # one side is null, other is not
    match_pct: float
    abs_tolerance: float
    rel_tolerance: float


@dataclass
class ComparisonResult:
    summary: ComparisonSummary
    row_results: list[RowResult]
    cell_diffs: list[CellDifference]
    missing_rows_df: pd.DataFrame    # client rows missing in output
    extra_rows_df: pd.DataFrame      # output rows missing in client
    column_mappings: list            # passed through for the report


# ---------------------------------------------------------------------------
# Value normalisation helpers (all vectorized)
# ---------------------------------------------------------------------------

def _normalize_series_text(s: pd.Series) -> pd.Series:
    """Strip whitespace and lowercase strings; leave non-strings as-is."""
    mask = s.apply(lambda v: isinstance(v, str))
    result = s.copy()
    result[mask] = s[mask].str.strip().str.lower()
    return result


def _try_numeric(s: pd.Series) -> pd.Series:
    """Attempt to coerce a series to float. Returns NaN where not possible."""
    return pd.to_numeric(s, errors="coerce")


def _is_datetime_value(v) -> bool:
    return isinstance(v, (dt_mod.datetime, dt_mod.date, pd.Timestamp)) and not isinstance(v, bool)


def _is_excel_serial_date(s: pd.Series) -> bool:
    """True if values look like Excel date serials (e.g. 46113 = 2026-04-01)."""
    sample = s.dropna().head(300)
    if sample.empty:
        return False
    if pd.api.types.is_datetime64_any_dtype(s):
        return False
    if sample.map(_is_datetime_value).mean() > 0.5:
        return False
    num = pd.to_numeric(sample, errors="coerce")
    if num.notna().mean() < 0.7:
        return False
    vals = num.dropna().astype(float)
    if vals.empty:
        return False
    integerish = (np.abs(vals - np.round(vals)) < 1e-6).mean() >= 0.9
    in_range = ((vals >= 20000) & (vals <= 80000)).mean() >= 0.8
    unique_ratio = vals.nunique() / len(vals)
    return bool(integerish and in_range and unique_ratio < 0.8)


def _try_datetime(s: pd.Series) -> pd.Series:
    """Parse strings/datetimes. Do not treat raw integers as ns timestamps."""
    try:
        if pd.api.types.is_numeric_dtype(s):
            return pd.Series(pd.NaT, index=s.index)
        return pd.to_datetime(s, errors="coerce")
    except Exception:
        return pd.Series(pd.NaT, index=s.index)


def _to_datetime_series(s: pd.Series) -> pd.Series:
    """Normalize a column to calendar dates, including Excel serials."""
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    valid = s.notna()
    if not valid.any():
        return out

    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce").dt.normalize()

    is_dt = s.map(_is_datetime_value)
    if is_dt.any():
        out.loc[is_dt] = pd.to_datetime(s[is_dt], errors="coerce")

    is_str = s.map(lambda v: isinstance(v, str))
    if is_str.any():
        # Do not use dayfirst=True: it turns ISO "2026-04-01" into 4 January.
        out.loc[is_str] = pd.to_datetime(s[is_str], errors="coerce")

    still = valid & out.isna()
    if still.any():
        num = pd.to_numeric(s[still], errors="coerce")
        ok = num.notna()
        if ok.any():
            idx = num.index[ok]
            out.loc[idx] = pd.to_datetime(num[ok].astype(float), unit="D", origin="1899-12-30")

    return pd.to_datetime(out, errors="coerce").dt.normalize()


def _detect_col_type(s: pd.Series) -> str:
    """Detect dominant type: 'numeric', 'datetime', 'text'."""
    sample = s.dropna().head(300)
    if sample.empty:
        return "text"
    if pd.api.types.is_datetime64_any_dtype(s) or sample.map(_is_datetime_value).mean() > 0.7:
        return "datetime"
    num = _try_numeric(sample)
    if num.notna().mean() > 0.7:
        return "numeric"
    dt = _try_datetime(sample)
    if dt.notna().mean() > 0.7:
        return "datetime"
    return "text"


def pair_is_date_column(col_c: pd.Series, col_o: pd.Series) -> bool:
    """True when this mapped pair should be compared/displayed as dates."""
    typ_c = _detect_col_type(col_c)
    typ_o = _detect_col_type(col_o)
    return (
        typ_c == "datetime"
        or typ_o == "datetime"
        or (_is_excel_serial_date(col_c) and _is_excel_serial_date(col_o))
    )


def format_date_cell(v):
    """Display a date or Excel serial as YYYY-MM-DD; otherwise return v."""
    if v is None or v == "":
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    ts = _to_datetime_series(pd.Series([v])).iloc[0]
    if pd.isna(ts):
        return v
    return pd.Timestamp(ts).strftime("%Y-%m-%d")


def _normalize_column_pair(
    col_c: pd.Series,
    col_o: pd.Series,
) -> tuple[pd.Series, pd.Series, str]:
    """
    Normalize a pair of columns to the same representation.
    Returns (norm_c, norm_o, detected_type).
    """
    typ_c = _detect_col_type(col_c)
    typ_o = _detect_col_type(col_o)

    # Datetime wins over numeric so Excel serials (46113) match real dates.
    if typ_c == "datetime" or typ_o == "datetime" or (
        _is_excel_serial_date(col_c) and _is_excel_serial_date(col_o)
    ):
        return _to_datetime_series(col_c), _to_datetime_series(col_o), "datetime"

    priority = {"numeric": 2, "datetime": 1, "text": 0}
    col_type = typ_c if priority[typ_c] >= priority[typ_o] else typ_o

    if col_type == "numeric":
        return _try_numeric(col_c), _try_numeric(col_o), "numeric"
    if col_type == "datetime":
        return _to_datetime_series(col_c), _to_datetime_series(col_o), "datetime"
    # text — strip and lowercase
    return _normalize_series_text(col_c.astype(str).where(col_c.notna(), other=np.nan)), \
           _normalize_series_text(col_o.astype(str).where(col_o.notna(), other=np.nan)), \
           "text"


# ---------------------------------------------------------------------------
# Tolerance comparison (vectorized)
# ---------------------------------------------------------------------------

def _numeric_match_mask(
    arr_c: np.ndarray,
    arr_o: np.ndarray,
    abs_tol: float,
    rel_tol: float,
) -> np.ndarray:
    """Return boolean mask where numeric values are within tolerance."""
    diff = np.abs(arr_c - arr_o)
    tol = np.maximum(abs_tol, rel_tol * np.maximum(np.abs(arr_c), np.abs(arr_o)))
    return diff <= tol


# ---------------------------------------------------------------------------
# Excel column letter helper (A1 notation)
# ---------------------------------------------------------------------------

def _col_letter(n: int) -> str:
    """Convert 1-based column index to Excel letter (1→A, 27→AA …)."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def excel_address(row: int, col_idx: int) -> str:
    """Return A1-notation cell address. row is 1-based Excel row number."""
    return f"{_col_letter(col_idx)}{row}"


# ---------------------------------------------------------------------------
# Main comparison engine
# ---------------------------------------------------------------------------

def compare_dataframes(
    aligned: "AlignedDataFrames",   # from row_matcher
    column_mappings: list,          # list of ColumnMapping (accepted + confirmed)
    df_client_orig: pd.DataFrame,
    df_output_orig: pd.DataFrame,
    abs_tol: float = 1e-6,
    rel_tol: float = 1e-4,
) -> ComparisonResult:
    """
    Compare aligned rows cell-by-cell using the provided column mappings.

    Parameters
    ----------
    aligned          : AlignedDataFrames from row_matcher.align_rows()
    column_mappings  : list of accepted/confirmed ColumnMapping objects
    df_client_orig   : original client df (for missing rows content)
    df_output_orig   : original output df (for extra rows content)
    abs_tol          : absolute numeric tolerance
    rel_tol          : relative numeric tolerance

    Returns
    -------
    ComparisonResult
    """
    merged = aligned.merged
    client_only = aligned.client_only
    output_only = aligned.output_only

    row_results: list[RowResult] = []
    cell_diffs: list[CellDifference] = []

    total_cells_compared = 0    
    matched_cells = 0
    mismatched_cells = 0
    missing_cell_values = 0

    # --- Build a lookup: client_col_name → output_col_name (in the merged df) ---
    # Merged df has columns like "SomeCol_client" and "SomeCol_output"
    # But column mappings use original names. We need to map them.
    col_map_pairs = [
        (m.client_col, m.output_col)
        for m in column_mappings
        if m.output_col is not None
    ]

    # #region agent log
    import json as _dbg_json, time as _dbg_time
    _DBG_PATH = r"c:\Users\anush\OneDrive\Documents\project_folders\excel_comparison\debug-54883f.log"
    def _agent_dbg(hypothesis_id, location, message, data=None):
        try:
            with open(_DBG_PATH, "a", encoding="utf-8") as _f:
                _f.write(_dbg_json.dumps({
                    "sessionId": "54883f",
                    "runId": "post-fix",
                    "hypothesisId": hypothesis_id,
                    "location": location,
                    "message": message,
                    "data": data or {},
                    "timestamp": int(_dbg_time.time() * 1000),
                    "id": f"log_{int(_dbg_time.time() * 1000)}",
                }, default=str) + "\n")
        except Exception:
            pass
    _t_cmp = _dbg_time.time()
    _agent_dbg("E", "comparator.py:compare_dataframes", "start", {
        "n_merged": len(merged),
        "n_missing": len(client_only),
        "n_extra": len(output_only),
        "n_cols": len(col_map_pairs),
    })
    _type_logs = 0
    # #endregion

    # For Excel column index (A1 notation) we need to know column position in client
    client_col_positions = {col: idx + 1 for idx, col in enumerate(df_client_orig.columns)}
    output_col_positions = {col: idx + 1 for idx, col in enumerate(df_output_orig.columns)}

    # --- Vectorized comparison for each column pair ---
    # Collect per-row mismatch flags across all columns
    n_rows = len(merged)
    mismatch_flags = np.zeros(n_rows, dtype=bool)

    # Store per-column comparison results for cell diff extraction
    col_comparison_results: dict[str, dict] = {}

    for client_col, output_col in col_map_pairs:
        c_key = f"{client_col}_client"
        o_key = f"{output_col}_output"

        if c_key not in merged.columns or o_key not in merged.columns:
            continue

        col_c = merged[c_key]
        col_o = merged[o_key]

        typ_c = _detect_col_type(col_c)
        typ_o = _detect_col_type(col_o)
        norm_c, norm_o, col_type = _normalize_column_pair(col_c, col_o)
        # #region agent log
        _serial = _is_excel_serial_date(col_c) or _is_excel_serial_date(col_o)
        _dt_mix = (typ_c == "datetime") != (typ_o == "datetime")
        if _serial or _dt_mix or (_type_logs < 5 and typ_c != typ_o):
            if not (_serial or _dt_mix):
                _type_logs += 1
            _sc = col_c.dropna()
            _so = col_o.dropna()
            _nc = norm_c.dropna()
            _no = norm_o.dropna()
            _agent_dbg("F", "comparator.py:compare_dataframes", "col type", {
                "client_col": str(client_col)[:60],
                "typ_c": typ_c,
                "typ_o": typ_o,
                "chosen": col_type,
                "serial_c": _is_excel_serial_date(col_c),
                "serial_o": _is_excel_serial_date(col_o),
                "sample_c": str(_sc.iloc[0])[:40] if len(_sc) else None,
                "sample_o": str(_so.iloc[0])[:40] if len(_so) else None,
                "norm_c": str(_nc.iloc[0])[:40] if len(_nc) else None,
                "norm_o": str(_no.iloc[0])[:40] if len(_no) else None,
            })
        # #endregion

        # Both null → match
        both_null = norm_c.isna() & norm_o.isna()
        # One null → missing value
        one_null = norm_c.isna() ^ norm_o.isna()
        # Both present
        both_present = ~norm_c.isna() & ~norm_o.isna()

        if col_type == "numeric":
            arr_c = norm_c.values.astype(float)
            arr_o = norm_o.values.astype(float)
            # Handle NaN in float arrays
            both_present_np = both_present.values
            match_arr = np.zeros(n_rows, dtype=bool)
            match_arr[both_null.values] = True
            if both_present_np.any():
                match_arr[both_present_np] = _numeric_match_mask(
                    arr_c[both_present_np],
                    arr_o[both_present_np],
                    abs_tol,
                    rel_tol,
                )
        elif col_type == "datetime":
            match_arr = np.zeros(n_rows, dtype=bool)
            match_arr[both_null.values] = True
            if both_present.any():
                match_arr[both_present.values] = (
                    norm_c[both_present].dt.strftime("%Y-%m-%d").values
                    == norm_o[both_present].dt.strftime("%Y-%m-%d").values
                )
        else:
            match_arr = np.zeros(n_rows, dtype=bool)
            match_arr[both_null.values] = True
            match_arr[both_present.values] = (norm_c[both_present] == norm_o[both_present]).values

        match_mask = pd.Series(match_arr, index=merged.index)
        mismatch_mask = ~match_mask & ~one_null
        one_null_mismatch = one_null

        total_cells_compared += n_rows
        matched_cells += int(match_mask.sum())
        mismatched_cells += int(mismatch_mask.sum())
        missing_cell_values += int(one_null_mismatch.sum())

        mismatch_flags |= (~match_mask).values

        col_comparison_results[client_col] = {
            "output_col": output_col,
            "col_type": col_type,
            "match_mask": match_mask,
            "mismatch_mask": mismatch_mask,
            "one_null": one_null_mismatch,
            "norm_c": norm_c,
            "norm_o": norm_o,
            "raw_c": col_c,
            "raw_o": col_o,
        }

    # --- Build row results ---
    row_key_col = merged.get("_row_key_", pd.Series([""] * n_rows))
    client_excel_rows = merged.get("_excel_row__client", pd.Series([None] * n_rows))
    output_excel_rows = merged.get("_excel_row__output", pd.Series([None] * n_rows))

    mismatch_series = pd.Series(mismatch_flags, index=merged.index)

    for i in range(n_rows):
        rk = str(row_key_col.iloc[i])
        cer = _safe_int(client_excel_rows.iloc[i])
        oer = _safe_int(output_excel_rows.iloc[i])
        is_mismatch = mismatch_flags[i]

        mismatch_cols = [
            cc for cc, res in col_comparison_results.items()
            if not res["match_mask"].iloc[i]
        ]
        status = "MISMATCH" if is_mismatch else "MATCH"
        row_results.append(RowResult(
            row_key=rk,
            client_excel_row=cer,
            output_excel_row=oer,
            status=status,
            mismatch_columns=mismatch_cols,
        ))

    # --- Build cell differences ---
    for client_col, res in col_comparison_results.items():
        output_col = res["output_col"]
        col_type = res["col_type"]
        mismatch_mask = res["mismatch_mask"]
        one_null = res["one_null"]
        raw_c = res["raw_c"]
        raw_o = res["raw_o"]
        norm_c = res["norm_c"]
        norm_o = res["norm_o"]

        diff_indices = mismatch_mask[mismatch_mask].index.tolist()
        null_indices = one_null[one_null].index.tolist()

        client_col_idx = client_col_positions.get(client_col, 1)
        output_col_idx = output_col_positions.get(output_col, 1)

        for idx in diff_indices + null_indices:
            row_i = merged.index.get_loc(idx)
            rk = str(row_key_col.iloc[row_i])
            cer = _safe_int(client_excel_rows.iloc[row_i])
            oer = _safe_int(output_excel_rows.iloc[row_i])
            cv = raw_c.iloc[row_i]
            ov = raw_o.iloc[row_i]

            # Numeric diff
            ndiff: Optional[float] = None
            dpct: Optional[float] = None
            if col_type == "numeric":
                cn = pd.to_numeric(cv, errors="coerce")
                on = pd.to_numeric(ov, errors="coerce")
                if pd.notna(cn) and pd.notna(on):
                    ndiff = float(abs(cn - on))
                    denom = max(abs(cn), abs(on))
                    dpct = (ndiff / denom * 100) if denom != 0 else 0.0

            issue = "MISMATCH"
            if pd.isna(cv) or str(cv).strip() in ("", "nan", "None"):
                issue = "MISSING_CLIENT"
            elif pd.isna(ov) or str(ov).strip() in ("", "nan", "None"):
                issue = "MISSING_OUTPUT"

            cell_diffs.append(CellDifference(
                row_key=rk,
                client_excel_row=cer,
                output_excel_row=oer,
                column_name=client_col,
                client_value=cv,
                output_value=ov,
                numeric_diff=ndiff,
                diff_pct=dpct,
                issue=issue,
            ))

    # De-duplicate cell diffs (indices appear in both diff_indices and null_indices)
    seen_diffs: set[tuple] = set()
    unique_diffs: list[CellDifference] = []
    for cd in cell_diffs:
        key = (cd.row_key, cd.column_name)
        if key not in seen_diffs:
            seen_diffs.add(key)
            unique_diffs.append(cd)
    cell_diffs = unique_diffs

    # --- Missing / extra rows ---
    missing_rows_df = _extract_original_rows(client_only, df_client_orig, side="client")
    extra_rows_df = _extract_original_rows(output_only, df_output_orig, side="output")

    missing_row_results = [
        RowResult(
            row_key=str(client_only["_row_key_"].iloc[i]) if "_row_key_" in client_only.columns else f"_missing_{i}",
            client_excel_row=_safe_int(client_only["_excel_row__client"].iloc[i]) if "_excel_row__client" in client_only.columns else None,
            output_excel_row=None,
            status="MISSING",
        )
        for i in range(len(client_only))
    ]
    extra_row_results = [
        RowResult(
            row_key=str(output_only["_row_key_"].iloc[i]) if "_row_key_" in output_only.columns else f"_extra_{i}",
            client_excel_row=None,
            output_excel_row=_safe_int(output_only["_excel_row__output"].iloc[i]) if "_excel_row__output" in output_only.columns else None,
            status="EXTRA",
        )
        for i in range(len(output_only))
    ]

    all_row_results = row_results + missing_row_results + extra_row_results
    matched_rows = sum(1 for r in row_results if r.status == "MATCH")
    mismatched_rows = sum(1 for r in row_results if r.status == "MISMATCH")

    total_compared_cells = total_cells_compared
    total_matched = matched_cells

    summary = ComparisonSummary(
        total_client_rows=len(df_client_orig),
        total_output_rows=len(df_output_orig),
        matched_rows=matched_rows,
        mismatched_rows=mismatched_rows,
        missing_rows=len(client_only),
        extra_rows=len(output_only),
        total_cells_compared=total_compared_cells,
        matched_cells=total_matched,
        mismatched_cells=mismatched_cells,
        missing_cell_values=missing_cell_values,
        match_pct=(total_matched / total_compared_cells * 100) if total_compared_cells > 0 else 0.0,
        abs_tolerance=abs_tol,
        rel_tolerance=rel_tol,
    )

    # #region agent log
    _agent_dbg("E", "comparator.py:compare_dataframes", "done", {
        "elapsed_ms": int((_dbg_time.time() - _t_cmp) * 1000),
        "n_diffs": len(cell_diffs),
        "matched_rows": matched_rows,
        "mismatched_rows": mismatched_rows,
    })
    # #endregion
    return ComparisonResult(
        summary=summary,
        row_results=all_row_results,
        cell_diffs=cell_diffs,
        missing_rows_df=missing_rows_df,
        extra_rows_df=extra_rows_df,
        column_mappings=column_mappings,
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_int(v) -> Optional[int]:
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _extract_original_rows(
    partial_merged: pd.DataFrame,
    orig_df: pd.DataFrame,
    side: str,
) -> pd.DataFrame:
    """Extract original column values for missing/extra rows from the original df."""
    suffix = f"_{side}"
    excel_col = f"_excel_row_{suffix}"
    if excel_col not in partial_merged.columns or partial_merged.empty:
        return pd.DataFrame()
    excel_rows = partial_merged[excel_col].dropna().astype(int).tolist()
    # Find matching rows in orig_df by _excel_row_
    if "_excel_row_" in orig_df.columns:
        return orig_df[orig_df["_excel_row_"].isin(excel_rows)].copy()
    return pd.DataFrame()
