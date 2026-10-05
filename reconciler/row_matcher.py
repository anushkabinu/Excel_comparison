"""
row_matcher.py
--------------
Automatically detects the best grouping field and identifier field from a
DataFrame, builds a composite row key, and aligns rows between two DataFrames.

Supports up to 3 key fields:
  - key_col_1 : primary grouping   (~2000 rows → groups)
  - key_col_2 : secondary grouping (~11 rows within each group)
  - key_col_3 : row identifier     (makes each row unique within group)

No column names are hard-coded. Detection is purely statistical.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
import numpy as np

# #region agent log
import json as _dbg_json
import time as _dbg_time

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
# #endregion


# ---------------------------------------------------------------------------
# Public dataclasses
# ---------------------------------------------------------------------------

@dataclass
class KeyFieldResult:
    grouping_col: Optional[str]   # primary grouping column
    id_col: Optional[str]         # identifier column
    grouping_score: float = 0.0
    id_score: float = 0.0
    all_scores: dict = field(default_factory=dict)


@dataclass
class AlignedDataFrames:
    merged: pd.DataFrame          # outer merged rows where both files matched
    client_only: pd.DataFrame     # rows in client not matched in output
    output_only: pd.DataFrame     # rows in output not matched in client
    inferred_identity_cols: list[str] = field(default_factory=list)
    ambiguous_rows: int = 0       # client rows whose identity values repeat


# ---------------------------------------------------------------------------
# Key-field detection
# ---------------------------------------------------------------------------

def _cardinality_ratio(series: pd.Series) -> float:
    """unique count / non-null count (0–1)."""
    nn = series.dropna()
    if len(nn) == 0:
        return 0.0
    return nn.nunique() / len(nn)


def _is_string_like(series: pd.Series) -> bool:
    """Return True if the majority of non-null values are non-numeric strings."""
    sample = series.dropna().head(200).astype(str)
    numeric_count = pd.to_numeric(sample, errors="coerce").notna().sum()
    return numeric_count / max(len(sample), 1) < 0.6


def _score_grouping(series: pd.Series) -> float:
    """
    Score a column as a grouping/category field.
    Ideal: medium cardinality (2–30%), string-like, not all-unique.
    """
    cr = _cardinality_ratio(series)
    # Sweet spot: 2% – 30% unique ratio
    if cr < 0.02 or cr > 0.30:
        return 0.0
    # Prefer lower cardinality (more repetition)
    cardinality_score = 1.0 - (cr / 0.30)
    str_score = 1.0 if _is_string_like(series) else 0.3
    # Penalise if too few unique values (1 value = useless)
    n_unique = series.dropna().nunique()
    diversity_score = min(n_unique / 5, 1.0)
    return 0.4 * cardinality_score + 0.3 * str_score + 0.3 * diversity_score


def _score_identifier(series: pd.Series) -> float:
    """
    Score a column as an identifier field.
    Ideal: high cardinality (>40%), string or integer, not all nulls.
    """
    cr = _cardinality_ratio(series)
    if cr < 0.40:
        return 0.0
    cardinality_score = min(cr / 1.0, 1.0)
    # Slight preference for string-like IDs (codes) but numeric IDs are fine
    str_score = 0.7 if _is_string_like(series) else 0.5
    return 0.7 * cardinality_score + 0.3 * str_score


def detect_key_fields(df: pd.DataFrame, skip_cols: Optional[list[str]] = None) -> KeyFieldResult:
    """Automatically detect the best grouping column and identifier column."""
    skip = set(skip_cols or [])
    # Exclude internal / metadata columns
    candidates = [c for c in df.columns if c not in skip and not c.startswith("_")]

    group_scores: dict[str, float] = {}
    id_scores: dict[str, float] = {}

    for col in candidates:
        group_scores[col] = _score_grouping(df[col])
        id_scores[col] = _score_identifier(df[col])

    best_group = max(group_scores, key=group_scores.__getitem__) if group_scores else None
    best_id = max(id_scores, key=id_scores.__getitem__) if id_scores else None

    # Avoid picking the same column for both roles
    if best_group == best_id and best_group is not None:
        # Pick second-best for identifier
        remaining_id = {k: v for k, v in id_scores.items() if k != best_group}
        best_id = max(remaining_id, key=remaining_id.__getitem__) if remaining_id else None

    # Zero scores → treat as not found
    if best_group and group_scores.get(best_group, 0) == 0:
        best_group = None
    if best_id and id_scores.get(best_id, 0) == 0:
        best_id = None

    return KeyFieldResult(
        grouping_col=best_group,
        id_col=best_id,
        grouping_score=group_scores.get(best_group, 0) if best_group else 0.0,
        id_score=id_scores.get(best_id, 0) if best_id else 0.0,
        all_scores={"grouping": group_scores, "identifier": id_scores},
    )


# ---------------------------------------------------------------------------
# Composite key builder  (supports up to 3 fields)
# ---------------------------------------------------------------------------

def _normalize_key_val(v) -> str:
    """
    Normalize ANY cell value to a canonical string for composite key matching.

    Handles every type that may appear from Excel/pandas:

    Type            → Canonical form
    ─────────────────────────────────────────────────────────────────────────
    None / NaN      → "__NULL__"
    bool            → "true" / "false"
    int / np.int*   → plain integer string  (42 → "42")
    float / np.float*
      whole number  → integer string        (100.0 → "100")
      fractional    → g-format, no trailing zeros  (3.50 → "3.5")
      nan / inf     → "__NULL__" / "__INF__"
    datetime / date → ISO-8601 date  (2024-03-15)  — time part stripped
    datetime string → normalised to ISO date where parseable
    general string  → lowercase, unicode NFC, collapse whitespace
    ─────────────────────────────────────────────────────────────────────────
    """
    import unicodedata
    import datetime as dt_mod

    # ── 1. None / pandas NA ────────────────────────────────────────────────
    try:
        if pd.isna(v):
            return "__NULL__"
    except (TypeError, ValueError):
        pass  # isna() can fail on some types; proceed

    # ── 2. Boolean  (must come before numeric — bool is subclass of int) ───
    if isinstance(v, (bool, np.bool_)):
        return "true" if v else "false"

    # ── 3. Native Python / numpy integers ──────────────────────────────────
    if isinstance(v, (int, np.integer)):
        return str(int(v))

    # ── 4. Native Python / numpy floats ────────────────────────────────────
    if isinstance(v, (float, np.floating)):
        if np.isnan(v):
            return "__NULL__"
        if np.isinf(v):
            return "__INF__"
        if v == int(v):
            return str(int(v))
        return f"{v:g}"

    # ── 5. datetime / date objects (from openpyxl / pandas) ────────────────
    if isinstance(v, (dt_mod.datetime, pd.Timestamp)):
        try:
            return pd.Timestamp(v).date().isoformat()  # → "2024-03-15"
        except Exception:
            pass
    if isinstance(v, dt_mod.date):
        return v.isoformat()

    # ── 6. Everything else: convert to string then normalise ────────────────
    s = str(v).strip()

    # Empty string
    if not s or s.lower() in ("nan", "none", "nat", "null", ""):
        return "__NULL__"

    # Try numeric string  ("100.0", "3,500.00", scientific "1.2e3")
    numeric_s = s.replace(",", "")           # remove thousands separators
    try:
        f = float(numeric_s)
        if np.isnan(f):
            return "__NULL__"
        if np.isinf(f):
            return "__INF__"
        if f == int(f):
            return str(int(f))
        return f"{f:g}"
    except (ValueError, OverflowError):
        pass

    # Try date/datetime string  (many formats)
    _DATE_FMTS = [
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y",
        "%d-%m-%Y", "%m-%d-%Y",
        "%d.%m.%Y", "%Y.%m.%d",
        "%d %b %Y", "%d %B %Y",
        "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S",
    ]
    for fmt in _DATE_FMTS:
        try:
            parsed = dt_mod.datetime.strptime(s, fmt)
            return parsed.date().isoformat()
        except ValueError:
            continue

    # General string: unicode NFC → lowercase → collapse whitespace
    s = unicodedata.normalize("NFC", s)
    s = s.lower()
    s = " ".join(s.split())
    return s


_NULL = "__NULL__"


def _values_agree(a, b) -> bool:
    """True if two cells hold the same value (numbers within a small tolerance)."""
    x, y = _normalize_key_val(a), _normalize_key_val(b)
    if x == y:
        return x != _NULL
    try:
        fx, fy = float(x), float(y)
    except ValueError:
        return False
    return abs(fx - fy) <= max(1e-6, 1e-4 * max(abs(fx), abs(fy)))

def _is_numeric_comparison_col(series: pd.Series) -> bool:
    """
    Returns True if this column looks like a comparison/value field
    (price, amount, quantity) rather than an identity field.
    Heuristic: mostly numeric AND has significant fractional values.
    """
    sample = series.dropna().head(300)
    if sample.empty:
        return False
    num = pd.to_numeric(sample, errors="coerce")
    if num.notna().mean() < 0.7:
        return False           # not mostly numeric — could be a string ID
    has_fraction = (num.dropna() % 1 != 0).mean()
    return bool(has_fraction > 0.1)   # >10% fractional values → comparison field

def _pick_identity_columns(
    df_client: pd.DataFrame,
    df_output: pd.DataFrame,
    primary_client: str,
    primary_output: str,
    mapped_pairs: list[tuple[str, Optional[str]]],
    min_unique: float = 0.2,
    min_overlap: float = 0.5,
    max_cols: int = 5,
) -> list[tuple[str, list[str], list[str]]]:
    """
    Choose mapped columns that can identify a row, best first.

    A column qualifies as an identity column when:
      1. It has enough distinct values in BOTH files (min_unique), AND
      2. It is NOT a numeric comparison field (price, amount, qty), AND
      3. EITHER:
         a. HIGH cardinality (>=40% unique in both files) -> trusted
            unconditionally regardless of cross-file overlap.  Low overlap
            here simply means many rows are MISSING — not that the column
            is expected to differ.  This fixes the case where client has
            23 000 rows and output has 2 000: the unique ID column shows
            ~9% cross-file overlap yet is the best possible identity key.
         b. MEDIUM cardinality (20-40%) -> requires >=50% cross-file overlap
            to guard against accidentally using a field that genuinely
            differs between the two files.

    Returns [(client_col, normalized_client_values, normalized_output_values), ...]
    ordered best identity column first.
    """
    HIGH_CARDINALITY = 0.40   # >=40% unique in both files -> unconditionally trusted
    scored = []
    # #region agent log
    _t_pick = _dbg_time.time()
    # #endregion

    for client_col, output_col in mapped_pairs:
        if not output_col or client_col == primary_client or output_col == primary_output:
            continue
        if client_col not in df_client.columns or output_col not in df_output.columns:
            continue

        unique_c = _cardinality_ratio(df_client[client_col])
        unique_o = _cardinality_ratio(df_output[output_col])
        unique = min(unique_c, unique_o)

        if unique < min_unique:
            continue    # too few distinct values to identify rows

        # Exclude numeric columns with fractional values (prices / amounts)
        if _is_numeric_comparison_col(df_client[client_col]):
            continue

        norm_c = df_client[client_col].map(_normalize_key_val)
        norm_o = df_output[output_col].map(_normalize_key_val)
        present = norm_c[norm_c != _NULL]
        if present.empty:
            continue

        is_high_card = (unique_c >= HIGH_CARDINALITY and unique_o >= HIGH_CARDINALITY)

        if is_high_card:
            # Tier 1 — high-cardinality: bypass overlap check.
            # Sort key: (tier=1, uniqueness) — higher uniqueness = better.
            score_key = (1, unique)
        else:
            # Tier 0 — medium-cardinality: require cross-file overlap.
            overlap = present.isin(set(norm_o)).mean()
            if overlap < min_overlap:
                continue
            bucket = int(overlap * 20) / 20
            score_key = (0, bucket, unique)

        scored.append((score_key, client_col, norm_c.tolist(), norm_o.tolist()))

    scored.sort(key=lambda t: t[0], reverse=True)
    picked = [(name, c, o) for _, name, c, o in scored[:max_cols]]
    # #region agent log
    _agent_dbg("B", "row_matcher.py:_pick_identity_columns", "pick done", {
        "elapsed_ms": int((_dbg_time.time() - _t_pick) * 1000),
        "min_unique": min_unique,
        "n_mapped": len(mapped_pairs),
        "n_scored": len(scored),
        "picked": [n for n, _, _ in picked],
        "n_client": len(df_client),
        "n_output": len(df_output),
    })
    # #endregion
    return picked



def _align_by_sequential_identity(
    df_client: pd.DataFrame,
    df_output: pd.DataFrame,
    primary_client: str,
    primary_output: str,
    mapped_pairs: list[tuple[str, Optional[str]]],
) -> Optional[AlignedDataFrames]:
    """
    Primary-field-only matching. For every client row:
      1. restrict to output rows with the same primary value;
      2. walk identity columns one by one, keeping only output rows whose value
         equals the client's value; stop as soon as exactly one row is left;
      3. a column that would eliminate every candidate is ignored (it is
         expected to differ); a row the best column cannot find is retried
         with the next column, and only reported missing if all fail;
      4. if identity values repeat, candidates are ranked by how many other
         mapped fields agree (ranking only — never used to reject a row).
    Each output row is used at most once. Returns None if no usable identity
    column exists (caller falls back to occurrence matching).
    """
    id_cols: list = []
    for min_unique in (0.2, 0.05):          # relax if rows repeat heavily
        id_cols = _pick_identity_columns(
            df_client, df_output, primary_client, primary_output, mapped_pairs,
            min_unique=min_unique,
        )
        if id_cols:
            break
    if not id_cols:
        return None

    prim_c = df_client[primary_client].map(_normalize_key_val).tolist()
    prim_o = df_output[primary_output].map(_normalize_key_val).tolist()
    c_vals = [c for _, c, _ in id_cols]
    o_vals = [o for _, _, o in id_cols]

    # Pre-normalize ranking columns once. Using raw cells + _values_agree here
    # re-parsed every client×candidate pair (91k pairs × 41 cols in the last run).
    _t_score_norm = _dbg_time.time()
    score_cols = [
        (
            df_client[cc].map(_normalize_key_val).tolist(),
            df_output[oc].map(_normalize_key_val).tolist(),
        )
        for cc, oc in mapped_pairs
        if oc and cc != primary_client and oc != primary_output
        and cc in df_client.columns and oc in df_output.columns
    ]
    _score_norm_ms = int((_dbg_time.time() - _t_score_norm) * 1000)

    n_ids = len(id_cols)

    # #region agent log
    _t_match = _dbg_time.time()
    _max_clients_grp = 0
    _max_pool = 0
    _n_edge_evals = 0
    _n_narrow_scans = 0
    _t_score_ms = 0.0
    _agent_dbg("A", "row_matcher.py:_align_by_sequential_identity", "match start", {
        "n_client": len(df_client),
        "n_output": len(df_output),
        "n_ids": n_ids,
        "id_names": [n for n, _, _ in id_cols],
        "n_score_cols": len(score_cols),
        "score_norm_ms": _score_norm_ms,
    })
    # #endregion

    # One lookup per identity column: (primary, value) -> output positions
    indexes: list[dict[tuple, list[int]]] = []
    for lvl in range(n_ids):
        idx: dict[tuple, list[int]] = {}
        for j, key in enumerate(zip(prim_o, o_vals[lvl])):
            idx.setdefault(key, []).append(j)
        indexes.append(idx)

    consumed: set[int] = set()

    pairs: list[tuple[int, int]] = []
    ambiguous = 0
    pending = list(range(len(df_client)))

    # Level 0 uses the best identity column. Rows it cannot find fall through
    # to the next column, so one differing field never drops a row by itself.
    for lvl in range(n_ids):
        if not pending:
            break
        groups: dict[tuple, list[int]] = {}
        still_pending: list[int] = []
        for i in pending:
            v = c_vals[lvl][i]
            if v == _NULL:
                still_pending.append(i)
            else:
                groups.setdefault((prim_c[i], v), []).append(i)

        for key, clients in groups.items():
            pool = [j for j in indexes[lvl].get(key, []) if j not in consumed]
            if not pool:
                still_pending.extend(clients)
                continue

            # #region agent log
            _max_clients_grp = max(_max_clients_grp, len(clients))
            _max_pool = max(_max_pool, len(pool))
            # #endregion

            # Narrow each client row's candidates with the other identity
            # columns, one at a time; stop once exactly one row is left.
            cands: dict[int, list[int]] = {}
            for i in clients:
                cand = pool
                for k in range(n_ids):
                    if k == lvl:
                        continue
                    if len(cand) <= 1:
                        break
                    v = c_vals[k][i]
                    if v == _NULL:
                        continue
                    # #region agent log
                    _n_narrow_scans += len(cand)
                    # #endregion
                    narrowed = [j for j in cand if o_vals[k][j] == v]
                    if narrowed:      # skip a column that eliminates everything
                        cand = narrowed
                cands[i] = cand

            singles = [c[0] for c in cands.values() if len(c) == 1]
            if len(singles) == len(clients) and len(set(singles)) == len(singles):
                chosen = [(i, cands[i][0]) for i in clients]
            else:
                ambiguous += len(clients)
                # #region agent log
                _n_edges = sum(len(cands[i]) for i in clients)
                _n_edge_evals += _n_edges * max(len(score_cols), 1)
                if _n_edges >= 200:
                    _agent_dbg("A", "row_matcher.py:_align_by_sequential_identity", "large ambiguous group", {
                        "lvl": lvl,
                        "n_clients": len(clients),
                        "n_edges": _n_edges,
                        "n_score_cols": len(score_cols),
                        "pool": len(pool),
                    })
                _ts = _dbg_time.time()
                # #endregion
                edges = []
                for i in clients:
                    for j in cands[i]:
                        score = sum(
                            1 for cr, orr in score_cols
                            if cr[i] != _NULL and cr[i] == orr[j]
                        )
                        edges.append((-score, i, j))
                edges.sort()
                # #region agent log
                _t_score_ms += (_dbg_time.time() - _ts) * 1000
                # #endregion
                used_c: set[int] = set()
                used_o: set[int] = set()
                chosen = []
                for _, i, j in edges:
                    if i in used_c or j in used_o:
                        continue
                    used_c.add(i)
                    used_o.add(j)
                    chosen.append((i, j))
                still_pending.extend(i for i in clients if i not in used_c)

            pairs.extend(chosen)
            consumed.update(j for _, j in chosen)

        pending = still_pending

    pairs.sort()
    unmatched_client = sorted(pending)
    unmatched_output = [j for j in range(len(df_output)) if j not in consumed]
    # #region agent log
    _agent_dbg("A", "row_matcher.py:_align_by_sequential_identity", "match done", {
        "elapsed_ms": int((_dbg_time.time() - _t_match) * 1000),
        "n_pairs": len(pairs),
        "n_pending": len(pending),
        "ambiguous": ambiguous,
        "max_clients_grp": _max_clients_grp,
        "max_pool": _max_pool,
        "n_edge_evals": _n_edge_evals,
        "n_narrow_scans": _n_narrow_scans,
        "t_score_ms": int(_t_score_ms),
    })
    # #endregion

    def _keys(prims, firsts, positions):
        seen: dict[str, int] = {}
        out = []
        for p in positions:
            base = f"{prims[p]}||{firsts[p]}"
            seen[base] = seen.get(base, 0) + 1
            out.append(f"{base}||#{seen[base]}")
        return out

    ci = [i for i, _ in pairs]
    oi = [j for _, j in pairs]

    both = pd.concat(
        [
            df_client.iloc[ci].add_suffix("_client").reset_index(drop=True),
            df_output.iloc[oi].add_suffix("_output").reset_index(drop=True),
        ],
        axis=1,
    )
    both["_row_key_"] = _keys(prim_c, c_vals[0], ci)

    client_only = df_client.iloc[unmatched_client].add_suffix("_client").reset_index(drop=True)
    client_only["_row_key_"] = _keys(prim_c, c_vals[0], unmatched_client)

    output_only = df_output.iloc[unmatched_output].add_suffix("_output").reset_index(drop=True)
    output_only["_row_key_"] = _keys(prim_o, o_vals[0], unmatched_output)

    return AlignedDataFrames(
        merged=both,
        client_only=client_only,
        output_only=output_only,
        inferred_identity_cols=[name for name, _, _ in id_cols],
        ambiguous_rows=ambiguous,
    )


# ---------------------------------------------------------------------------
# Multi-key composite matching with disambiguation
# ---------------------------------------------------------------------------

def _align_composite_with_disambiguation(
    df_client: pd.DataFrame,
    df_output: pd.DataFrame,
    client_key_cols: list[str],
    output_key_cols: list[str],
    mapped_pairs: list[tuple[str, Optional[str]]],
) -> AlignedDataFrames:
    """
    Match rows using a composite key (Field 1 + Field 2 [+ Field 3]).

    For groups where the composite key is unique in both files → direct match.
    For groups where the composite key repeats → progressively narrow using the
    remaining mapped columns (highest cardinality first, same 'skip if zeroes'
    rule as the single-key path).  If still multiple candidates remain, rank by
    total agreement count and accept the best.  If no output row exists for a
    client key → MISSING; extra output rows → EXTRA.

    A cell difference must NOT prevent a row from being matched.  Only the
    selected key columns drive the initial grouping; all mapped columns are
    compared afterwards by the comparator.
    """
    # ── Build normalized base keys ─────────────────────────────────────────
    def _make_base_keys(df: pd.DataFrame, key_cols: list[str]) -> list[str]:
        parts = [df[c].map(_normalize_key_val) for c in key_cols if c in df.columns]
        if not parts:
            return [f"_row_{i}_" for i in range(len(df))]
        result = parts[0]
        for p in parts[1:]:
            result = result + "||" + p
        return result.tolist()

    base_c = _make_base_keys(df_client, client_key_cols)
    base_o = _make_base_keys(df_output, output_key_cols)

    # ── Remaining mapped columns (not used as key cols) ────────────────────
    key_c_set = set(client_key_cols)
    key_o_set = set(output_key_cols)
    extra_pairs = [
        (cc, oc) for cc, oc in mapped_pairs
        if oc and cc not in key_c_set and oc not in key_o_set
        and cc in df_client.columns and oc in df_output.columns
    ]

    # Pre-normalise remaining columns once (avoid repeated per-row calls)
    norm_c = {cc: df_client[cc].map(_normalize_key_val).tolist() for cc, _ in extra_pairs}
    norm_o = {oc: df_output[oc].map(_normalize_key_val).tolist() for _, oc in extra_pairs}

    # Sort remaining pairs by min cardinality (most unique first → best discriminator)
    def _card(vals: list) -> float:
        non_null = [v for v in vals if v != _NULL]
        return len(set(non_null)) / len(non_null) if non_null else 0.0

    extra_pairs_ranked = sorted(
        extra_pairs,
        key=lambda p: min(_card(norm_c[p[0]]), _card(norm_o[p[1]])),
        reverse=True,
    )

    # ── Build output lookup: base_key → [output positions] ────────────────
    out_by_key: dict[str, list[int]] = {}
    for j, key in enumerate(base_o):
        out_by_key.setdefault(key, []).append(j)

    # ── Match client rows ──────────────────────────────────────────────────
    pairs: list[tuple[int, int]] = []
    unmatched_client: list[int] = []
    consumed: set[int] = set()
    ambiguous_count = 0

    # Group client rows by their base key
    cli_by_key: dict[str, list[int]] = {}
    for i, key in enumerate(base_c):
        cli_by_key.setdefault(key, []).append(i)

    for key, cli_idxs in cli_by_key.items():
        all_pool = out_by_key.get(key, [])

        if not all_pool:
            unmatched_client.extend(cli_idxs)
            continue

        for i in cli_idxs:
            pool = [j for j in all_pool if j not in consumed]

            if not pool:
                unmatched_client.append(i)
                continue

            if len(pool) == 1:
                pairs.append((i, pool[0]))
                consumed.add(pool[0])
                continue

            # Progressive narrowing using ranked extra columns
            current = pool
            for cc, oc in extra_pairs_ranked:
                c_val = norm_c[cc][i]
                if c_val == _NULL:
                    continue                       # skip null client value
                narrowed = [j for j in current if norm_o[oc][j] == c_val]
                if narrowed:                      # only keep if not zero candidates
                    current = narrowed
                if len(current) == 1:
                    break

            if len(current) == 1:
                j = current[0]
                if j not in consumed:
                    pairs.append((i, j))
                    consumed.add(j)
                else:
                    unmatched_client.append(i)
                    ambiguous_count += 1
            else:
                # Still multiple → rank by total field agreement, take best
                best_score, best_j = -1, None
                for j in current:
                    if j in consumed:
                        continue
                    score = sum(
                        norm_c[cc][i] == norm_o[oc][j]
                        for cc, oc in extra_pairs_ranked
                        if norm_c[cc][i] != _NULL
                    )
                    if score > best_score:
                        best_score, best_j = score, j

                if best_j is not None:
                    pairs.append((i, best_j))
                    consumed.add(best_j)
                    ambiguous_count += 1
                else:
                    unmatched_client.append(i)
                    ambiguous_count += 1

    unmatched_output = [j for j in range(len(df_output)) if j not in consumed]

    # ── Assemble output DataFrames in the same format as align_rows ────────
    pairs.sort()
    ci = [i for i, _ in pairs]
    oi = [j for _, j in pairs]

    # Unique row key per pair (base key + occurrence within the matched set)
    seen: dict[str, int] = {}
    row_keys = []
    for i in ci:
        k = base_c[i]
        seen[k] = seen.get(k, 0) + 1
        row_keys.append(f"{k}||#{seen[k]}")

    both = pd.concat([
        df_client.iloc[ci].add_suffix("_client").reset_index(drop=True),
        df_output.iloc[oi].add_suffix("_output").reset_index(drop=True),
    ], axis=1)
    both["_row_key_"] = row_keys

    def _build_unmatched_keys(idxs, base_vals, suffix):
        seen2: dict[str, int] = {}
        keys = []
        for idx in idxs:
            k = base_vals[idx]
            seen2[k] = seen2.get(k, 0) + 1
            keys.append(f"{k}||#{seen2[k]}")
        return keys

    client_only = df_client.iloc[unmatched_client].add_suffix("_client").reset_index(drop=True)
    client_only["_row_key_"] = _build_unmatched_keys(unmatched_client, base_c, "c")

    output_only = df_output.iloc[unmatched_output].add_suffix("_output").reset_index(drop=True)
    output_only["_row_key_"] = _build_unmatched_keys(unmatched_output, base_o, "o")

    return AlignedDataFrames(
        merged=both,
        client_only=client_only,
        output_only=output_only,
        inferred_identity_cols=[],
        ambiguous_rows=ambiguous_count,
    )


def build_composite_key(
    df: pd.DataFrame,
    key_cols: list[Optional[str]],
    sort_within_col: Optional[str] = None,
) -> pd.DataFrame:
    """
    Build a composite row key from up to 3 columns.

    key_cols        : [primary_col, secondary_col, id_col] — None entries skipped.
    sort_within_col : optional extra column used to sort rows WITHIN each primary
                      group before assigning occurrence numbers. It is NOT added
                      to the composite key — only used for ordering.

    Steps:
    1. Normalize each key column value with _normalize_key_val()
    2. Sort rows by:
         a. all normalized key columns  (primary group ordering)
         b. then by sort_within_col     (within-group ordering, if provided)
    3. Build _base_key_ by concatenating normalized key values
    4. Assign _occurrence_ counter (1, 2, 3 …) per base key
    5. Build final _row_key_ = _base_key_ + "||#" + _occurrence_

    Returns a copy of df sorted according to the above rules.
    """
    df = df.copy()

    # Filter None / empty slots
    active_cols = [c for c in key_cols if c]

    if active_cols:
        # Step 1: normalize key columns into temp columns
        norm_cols = []
        for col in active_cols:
            tmp = f"_nk_{col}_"
            df[tmp] = df[col].map(_normalize_key_val)
            norm_cols.append(tmp)

        # Step 2a+b: sort by key cols first, then by the within-group sort col
        sort_by = list(norm_cols)
        if sort_within_col and sort_within_col in df.columns:
            tmp_sw = "_nk_sort_within_"
            df[tmp_sw] = df[sort_within_col].map(_normalize_key_val)
            sort_by.append(tmp_sw)
        else:
            tmp_sw = None

        df = df.sort_values(sort_by, kind="mergesort", ignore_index=True)

        # Step 3: build _base_key_ from the key-column normalized values
        if len(norm_cols) == 1:
            df["_base_key_"] = df[norm_cols[0]]
        else:
            df["_base_key_"] = df[norm_cols[0]].str.cat(
                [df[c] for c in norm_cols[1:]], sep="||"
            )

        # Drop temp columns
        cols_to_drop = norm_cols + ([tmp_sw] if tmp_sw else [])
        df.drop(columns=cols_to_drop, inplace=True)
    else:
        # No key cols: use stable row position as key
        df["_base_key_"] = [f"_row_{i}_" for i in range(len(df))]

    # Step 4 & 5: occurrence counter + final key
    df["_occurrence_"] = df.groupby("_base_key_").cumcount() + 1
    df["_row_key_"] = df["_base_key_"] + "||#" + df["_occurrence_"].astype(str)

    return df


# ---------------------------------------------------------------------------
# Row alignment
# ---------------------------------------------------------------------------

def align_rows(
    df_client: pd.DataFrame,
    df_output: pd.DataFrame,
    client_key_cols: list[Optional[str]],
    output_key_cols: list[Optional[str]],
    client_sort_col: Optional[str] = None,
    output_sort_col: Optional[str] = None,
    mapped_column_pairs: Optional[list[tuple[str, Optional[str]]]] = None,
) -> AlignedDataFrames:
    """
    Align rows between two DataFrames using composite keys.

    Parameters
    ----------
    df_client / df_output   : DataFrames with _excel_row_ column
    client_key_cols         : [primary, secondary, id] column names in client
                              (pass None for unused slots)
    output_key_cols         : corresponding column names in output
    client_sort_col         : optional column in client used to sort rows WITHIN
                              each primary group (not added to the key)
    output_sort_col         : corresponding sort column in output file
    mapped_column_pairs     : (client_col, output_col) pairs used to:
                              a) find the matching output row inside each primary
                                 group (single-key path), OR
                              b) disambiguate duplicate composite keys (multi-key path)
    """
    active_client = [c for c in client_key_cols if c]
    active_output = [c for c in output_key_cols if c]

    # #region agent log
    _t_align = _dbg_time.time()
    _agent_dbg("E", "row_matcher.py:align_rows", "start", {
        "n_client": len(df_client),
        "n_output": len(df_output),
        "client_keys": active_client,
        "n_mapped": len(mapped_column_pairs or []),
    })
    # #endregion

    # ── Case 1: Only Field 1 selected → progressive identity matching ──────
    if len(active_client) == 1 and len(active_output) == 1 and mapped_column_pairs:
        aligned = _align_by_sequential_identity(
            df_client, df_output,
            active_client[0], active_output[0],
            mapped_column_pairs,
        )
        if aligned is not None:
            # #region agent log
            _agent_dbg("A", "row_matcher.py:align_rows", "sequential done", {
                "elapsed_ms": int((_dbg_time.time() - _t_align) * 1000),
                "merged": len(aligned.merged),
                "client_only": len(aligned.client_only),
                "output_only": len(aligned.output_only),
                "ambiguous": aligned.ambiguous_rows,
                "id_cols": aligned.inferred_identity_cols,
            })
            # #endregion
            return aligned

    # ── Case 2: Field 2 or 3 also selected → composite key + disambiguation ─
    if len(active_client) >= 2 and mapped_column_pairs:
        aligned = _align_composite_with_disambiguation(
            df_client, df_output,
            active_client, active_output,
            mapped_column_pairs,
        )
        # #region agent log
        _agent_dbg("D", "row_matcher.py:align_rows", "composite done", {
            "elapsed_ms": int((_dbg_time.time() - _t_align) * 1000),
            "merged": len(aligned.merged),
            "client_only": len(aligned.client_only),
            "output_only": len(aligned.output_only),
            "ambiguous": aligned.ambiguous_rows,
        })
        # #endregion
        return aligned

    # ── Case 3: Fallback — occurrence-based composite key ─────────────────
    # (used when no mapped_column_pairs are provided, e.g. legacy callers)
    inferred_labels: list[str] = []
    df_c = build_composite_key(df_client, client_key_cols, sort_within_col=client_sort_col)
    df_o = build_composite_key(df_output, output_key_cols, sort_within_col=output_sort_col)

    merged = pd.merge(
        df_c.add_suffix("_client").rename(columns={"_row_key__client": "_row_key_"}),
        df_o.add_suffix("_output").rename(columns={"_row_key__output": "_row_key_"}),
        on="_row_key_",
        how="outer",
        indicator=True,
    )

    client_only = merged[merged["_merge"] == "left_only"].copy()
    output_only = merged[merged["_merge"] == "right_only"].copy()
    both = merged[merged["_merge"] == "both"].copy()

    for frame in [both, client_only, output_only]:
        if "_row_key_" in frame.columns:
            frame.sort_values("_row_key_", inplace=True, ignore_index=True)

    # #region agent log
    _agent_dbg("E", "row_matcher.py:align_rows", "fallback done", {
        "elapsed_ms": int((_dbg_time.time() - _t_align) * 1000),
        "merged": len(both),
        "client_only": len(client_only),
        "output_only": len(output_only),
    })
    # #endregion
    return AlignedDataFrames(
        merged=both,
        client_only=client_only,
        output_only=output_only,
        inferred_identity_cols=inferred_labels,
    )

