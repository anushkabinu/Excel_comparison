"""
header_mapper.py
----------------
Normalizes column headers and computes a multi-factor similarity score
between columns from two DataFrames.

Scoring factors:
  1. name_score    – fuzzy string similarity on normalized header names
  2. synonym_score – curated synonym groups (generic business terms only)
  3. type_score    – inferred dtype compatibility
  4. overlap_score – Jaccard similarity on sampled unique values
  5. unique_score  – similarity in cardinality ratio

No file-specific column names are hard-coded.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd
from rapidfuzz import fuzz

# ---------------------------------------------------------------------------
# Confidence thresholds
# ---------------------------------------------------------------------------
AUTO_ACCEPT_THRESHOLD   = 0.50  # score >= 50% → auto-map; below → unmapped
NEEDS_CONFIRM_THRESHOLD = 0.50  # kept for backward compat (same as auto-accept)

# ---------------------------------------------------------------------------
# Generic synonym groups (no business-specific names)
# ---------------------------------------------------------------------------
_SYNONYM_GROUPS: list[set[str]] = [
    {"id", "identifier", "code", "key", "ref", "reference", "no", "num", "number"},
    {"name", "label", "title", "description", "desc", "details"},
    {"qty", "quantity", "count", "cnt", "volume", "vol", "units", "unit"},
    {"amt", "amount", "total", "sum", "value", "val"},
    {"price", "rate", "cost", "charge", "fee", "tariff", "unit_price"},
    {"date", "dt", "time", "timestamp", "datetime", "period", "day", "month", "year"},
    {"type", "category", "cat", "class", "group", "grp", "segment", "seg", "kind"},
    {"status", "state", "flag", "indicator", "ind"},
    {"pct", "percent", "percentage", "ratio", "share", "proportion"},
    {"weight", "wt", "mass"},
    {"height", "ht", "length", "len", "width", "wid", "size", "dimension"},
    {"address", "addr", "location", "loc", "site"},
    {"country", "nation", "region", "area", "zone", "territory"},
    {"currency", "ccy", "fx", "exch"},
    {"balance", "bal", "remaining", "net"},
    {"tax", "vat", "gst", "duty", "levy"},
    {"profit", "margin", "gain", "loss", "pnl", "pl"},
    {"remark", "remark", "note", "comment", "memo", "annotation"},
    {"seq", "sequence", "order", "rank", "priority", "position", "pos", "row"},
    {"concentration", "conc", "level", "grade", "strength"},
    {"material", "product", "item", "article", "part", "component", "sku", "stock"},
    {"supplier", "vendor", "provider", "source", "origin"},
    {"customer", "client", "buyer", "account", "party"},
    {"batch", "lot", "run", "job", "transaction", "txn"},
]

# Pre-build lookup: token → group index
_TOKEN_TO_GROUP: dict[str, int] = {}
for _gidx, _grp in enumerate(_SYNONYM_GROUPS):
    for _tok in _grp:
        _TOKEN_TO_GROUP[_tok] = _gidx


# ---------------------------------------------------------------------------
# Public dataclass
# ---------------------------------------------------------------------------

@dataclass
class ColumnMapping:
    client_col: str
    output_col: Optional[str]        # None if unmapped
    confidence: float                # 0–1
    name_score: float = 0.0
    synonym_score: float = 0.0
    type_score: float = 0.0
    overlap_score: float = 0.0
    unique_score: float = 0.0
    status: str = "auto"             # "auto" | "unmapped" | "user_mapped"
    mapping_reason: str = ""         # human-readable explanation


@dataclass
class MappingResult:
    auto_accepted: list[ColumnMapping] = field(default_factory=list)
    needs_confirmation: list[ColumnMapping] = field(default_factory=list)
    unmapped_client: list[str] = field(default_factory=list)
    unmapped_output: list[str] = field(default_factory=list)
    empty_client_cols: list[str] = field(default_factory=list)   # all-NaN columns in client file
    empty_output_cols: list[str] = field(default_factory=list)   # all-NaN columns in output file


# ---------------------------------------------------------------------------
# Normalization
# ---------------------------------------------------------------------------

def normalize_header(s: str) -> str:
    """
    Normalize a header string:
    - unicode NFC
    - lowercase
    - strip leading/trailing whitespace
    - replace runs of spaces, underscores, hyphens, dots with a single space
    - remove all remaining non-alphanumeric characters
    """
    s = unicodedata.normalize("NFC", str(s))
    s = s.lower().strip()
    # Replace separators with space
    s = re.sub(r"[\s_\-\.]+", " ", s)
    # Remove non-alphanumeric (except space)
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return s.strip()


def _tokenize(norm: str) -> list[str]:
    return norm.split()


# ---------------------------------------------------------------------------
# Sub-score functions
# ---------------------------------------------------------------------------

def _name_score(norm_a: str, norm_b: str) -> float:
    """Fuzzy string similarity on normalized names (0–1)."""
    if not norm_a or not norm_b:
        return 0.0
    # Use token_sort_ratio to handle word-order differences
    score = fuzz.token_sort_ratio(norm_a, norm_b) / 100.0
    # Also check exact match for a boost
    if norm_a == norm_b:
        return 1.0
    return score


def _synonym_score(norm_a: str, norm_b: str) -> float:
    """
    Return 1.0 if any token pair shares a synonym group,
    0.5 if partial overlap, 0.0 otherwise.
    """
    tokens_a = set(_tokenize(norm_a))
    tokens_b = set(_tokenize(norm_b))
    groups_a = {_TOKEN_TO_GROUP[t] for t in tokens_a if t in _TOKEN_TO_GROUP}
    groups_b = {_TOKEN_TO_GROUP[t] for t in tokens_b if t in _TOKEN_TO_GROUP}
    if not groups_a or not groups_b:
        return 0.0
    overlap = groups_a & groups_b
    if overlap:
        return 1.0
    return 0.0


def _infer_dtype_category(series: pd.Series) -> str:
    """Classify a series as 'numeric', 'datetime', or 'text'."""
    sample = series.dropna().head(200)
    if sample.empty:
        return "text"
    # Try numeric
    numeric = pd.to_numeric(sample, errors="coerce")
    if numeric.notna().mean() > 0.7:
        return "numeric"
    # Try common date formats
    date_formats = [
        "%b-%y",      # Mar-26
        "%b-%Y",      # Mar-2026
        "%B-%y",      # March-26
        "%B-%Y",      # March-2026
        "%d-%b-%y",
        "%d-%b-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
    ]

    for fmt in date_formats:
        parsed = pd.to_datetime(sample, format=fmt, errors="coerce")
        if parsed.notna().mean() > 0.7:
            return "datetime"

    # Final generic date check
    try:
        parsed = pd.to_datetime(sample, errors="coerce")
        if parsed.notna().mean() > 0.7:
            return "datetime"
    except Exception:
        pass

    return "text"

def _type_score(dtype_a: str, dtype_b: str) -> float:
    if dtype_a == dtype_b:
        return 1.0
    # Numeric ↔ text can still match (IDs stored as strings)
    if {dtype_a, dtype_b} == {"numeric", "text"}:
        return 0.4
    return 0.1




def _unique_score(series_a: pd.Series, series_b: pd.Series) -> float:
    """Similarity in uniqueness ratio (cardinality / len)."""
    def _uratio(s: pd.Series) -> float:
        n = len(s.dropna())
        return s.dropna().nunique() / n if n > 0 else 0.0

    ua, ub = _uratio(series_a), _uratio(series_b)
    return 1.0 - abs(ua - ub)


def _composite_score(
    name: float, synonym: float, type_: float, unique: float
) -> float:
    """Weighted composite of sub-scores (overlap removed)."""
    return (
        0.45 * name
        + 0.25 * synonym
        + 0.15 * type_
        + 0.15 * unique
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def compute_column_similarity(
    col_a: str,
    col_b: str,
    df_client: pd.DataFrame,
    df_output: pd.DataFrame,
) -> ColumnMapping:
    """Compute a full ColumnMapping between col_a (client) and col_b (output)."""
    try:
        norm_a = normalize_header(col_a)
        norm_b = normalize_header(col_b)

        ns = _name_score(norm_a, norm_b)
        ss = _synonym_score(norm_a, norm_b)

        try:
            dtype_a = _infer_dtype_category(df_client[col_a])
        except Exception:
            dtype_a = "text"
        try:
            dtype_b = _infer_dtype_category(df_output[col_b])
        except Exception:
            dtype_b = "text"
        ts = _type_score(dtype_a, dtype_b)

        try:
            us = _unique_score(df_client[col_a], df_output[col_b])
        except Exception:
            us = 0.5

        conf = _composite_score(ns, ss, ts, us)

        return ColumnMapping(
            client_col=col_a,
            output_col=col_b,
            confidence=conf,
            name_score=ns,
            synonym_score=ss,
            type_score=ts,
            overlap_score=0.0,
            unique_score=us,
        )
    except Exception:
        # Fallback: return a zero-confidence mapping so we never crash
        return ColumnMapping(
            client_col=col_a,
            output_col=col_b,
            confidence=0.0,
        )


def _build_mapping_reason(m: "ColumnMapping", mapped: bool) -> str:
    """
    Generate a plain-English explanation of why a column was mapped or
    left unmapped — written for a tester, with no score numbers.
    """
    if not mapped:
        best_col = m.output_col or "any output column"
        if m.name_score > 0.4:
            return (
                f"Column name is partially similar to '{best_col}' "
                f"but the difference is too large to map automatically — "
                f"please review and map manually if needed"
            )
        if m.overlap_score > 0.3:
            return (
                f"Some data values overlap with '{best_col}' "
                f"but the column names are too different to map automatically — "
                f"please review and map manually if needed"
            )
        return "No matching column found in the output file — map manually if this column should be compared"

    # ── Mapped — build a plain description of what evidence triggered the match ──
    client = m.client_col
    output = m.output_col or ""

    # Primary reason: name relationship
    if m.name_score >= 0.95:
        name_reason = f"'{client}' and '{output}' have identical or near-identical names"
    elif m.name_score >= 0.80:
        name_reason = f"'{client}' and '{output}' have very similar names (spelling/casing difference)"
    elif m.name_score >= 0.60:
        name_reason = f"'{client}' and '{output}' have partially similar names"
    else:
        name_reason = None

    # Synonym reason
    if m.synonym_score >= 0.8:
        syn_reason = f"'{client}' and '{output}' are known synonyms for the same concept"
    else:
        syn_reason = None

    # Data type reason
    if m.type_score >= 0.9:
        type_reason = "both columns contain the same type of data"
    elif m.type_score >= 0.6:
        type_reason = "both columns contain compatible types of data"
    else:
        type_reason = None

    # Value overlap reason
    if m.overlap_score >= 0.5:
        val_reason = "both columns share a large portion of the same values"
    elif m.overlap_score >= 0.3:
        val_reason = "both columns share some of the same values"
    else:
        val_reason = None

    # Build a single coherent sentence
    parts = [r for r in [name_reason, syn_reason, type_reason, val_reason] if r]
    if not parts:
        return f"'{client}' matched to '{output}' — sufficient overall similarity across multiple signals"

    if len(parts) == 1:
        return parts[0].capitalize()

    # Join: first part + supporting evidence
    return parts[0].capitalize() + "; additionally, " + " and ".join(parts[1:])


def _is_empty_col(series: pd.Series) -> bool:
    """
    Return True only if the column has NO real data at all.
    Handles both NaN (truly null) and empty/whitespace strings
    that can come from reading Excel files with dtype=str.
    A column is considered empty only if EVERY value is either:
      - NaN  (null cell in Excel)
      - ""   (empty string)
      - whitespace only (e.g. "  ")
    If even ONE cell has real content, the column is NOT empty.
    """
    for val in series:
        if pd.notna(val) and str(val).strip() != "":
            return False   # found a real value → not empty
    return True            # no real value found → empty


def map_columns(
    df_client: pd.DataFrame,
    df_output: pd.DataFrame,
    skip_cols: Optional[list[str]] = None,
) -> MappingResult:
    """
    Build the best 1-to-1 mapping from client columns to output columns.

    Columns scoring >= AUTO_ACCEPT_THRESHOLD (50%) are auto-mapped.
    Columns below the threshold go to unmapped_client.
    needs_confirmation is always empty (kept for backward compatibility).

    Empty columns (all-NaN / all-blank) are detected AFTER mapping completes
    and stored in MappingResult for UI display only — they do NOT affect the
    mapping or scoring in any way.
    """
    skip = set(skip_cols or [])
    client_cols = [c for c in df_client.columns if c not in skip and not c.startswith("_")]
    output_cols = [c for c in df_output.columns if c not in skip and not c.startswith("_")]

    # Build full score matrix — all columns participate, including empty ones
    score_matrix: dict[tuple[str, str], ColumnMapping] = {}
    for cc in client_cols:
        for oc in output_cols:
            mapping = compute_column_similarity(cc, oc, df_client, df_output)
            score_matrix[(cc, oc)] = mapping

    # Greedy 1-to-1 assignment
    assigned_client: set[str] = set()
    assigned_output: set[str] = set()
    sorted_pairs = sorted(score_matrix.keys(), key=lambda p: score_matrix[p].confidence, reverse=True)

    result = MappingResult()

    for (cc, oc) in sorted_pairs:
        if cc in assigned_client or oc in assigned_output:
            continue
        mapping = score_matrix[(cc, oc)]
        if mapping.confidence >= AUTO_ACCEPT_THRESHOLD:
            mapping.status = "auto"
            mapping.mapping_reason = _build_mapping_reason(mapping, mapped=True)
            result.auto_accepted.append(mapping)
            assigned_client.add(cc)
            assigned_output.add(oc)
        # below threshold: don't assign, let it fall to unmapped

    # Remaining unmapped — find best candidate to explain why not mapped
    for cc in client_cols:
        if cc in assigned_client:
            continue
        best_m = max(
            (score_matrix[(cc, oc)] for oc in output_cols),
            key=lambda m: m.confidence,
            default=None,
        )
        if best_m:
            best_m.mapping_reason = _build_mapping_reason(best_m, mapped=False)
        result.unmapped_client.append(cc)

    result.unmapped_output = [c for c in output_cols if c not in assigned_output]

    # Detect empty columns for UI display only — does not affect mapping
    all_c = [c for c in df_client.columns if not c.startswith("_")]
    all_o = [c for c in df_output.columns if not c.startswith("_")]
    result.empty_client_cols = [c for c in all_c if _is_empty_col(df_client[c])]
    result.empty_output_cols = [c for c in all_o if _is_empty_col(df_output[c])]

    return result
