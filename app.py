"""
app.py
------
Streamlit UI for the Excel Reconciliation Tool.

Flow:
  Upload    — Files are uploaded; a raw-row pre-check runs immediately using
              read_excel_raw() (no sheet/header detection, no column mapping).
  Step 2    — User clicks "Analyse Files": read_excel_file() + map_columns()
              run now (sheet detection, header detection, header mapping).
  Step 3    — User confirms ambiguous mappings & key fields
  Step 4    — Set numeric tolerance
  Step 5    — Run comparison
  Step 6    — View summary & download Excel report
"""

from __future__ import annotations

import io
import os
import tempfile
from typing import Optional
import time

import pandas as pd
import streamlit as st

from reconciler.file_reader import read_excel_file, read_excel_raw, ExcelReadResult
from reconciler.header_mapper import (
    map_columns,
    MappingResult,
    ColumnMapping,
    AUTO_ACCEPT_THRESHOLD,
    NEEDS_CONFIRM_THRESHOLD,
    normalize_header,
)
from reconciler.row_matcher import detect_key_fields, align_rows, KeyFieldResult
from reconciler.comparator import compare_dataframes, ComparisonResult
from reconciler.report_generator import generate_report
from reconciler.row_comparison import run_simple_row_match, check_duplicate_rows

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Excel Reconciliation Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.main { background: #0f1117; }

/* Hero header */
.hero-header {
    background: linear-gradient(135deg, #1a1f35 0%, #0d1b2a 50%, #162032 100%);
    border: 1px solid #2a3a5c;
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4);
}
.hero-title {
    font-size: 2rem;
    font-weight: 700;
    background: linear-gradient(90deg, #60a5fa, #a78bfa, #34d399);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin: 0 0 0.4rem 0;
}
.hero-subtitle {
    color: #94a3b8;
    font-size: 1rem;
    margin: 0;
    font-weight: 400;
}

/* Step cards */
.step-card {
    background: #1e2433;
    border: 1px solid #2d3a55;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
    transition: border-color 0.2s;
}
.step-card:hover { border-color: #60a5fa; }
.step-number {
    background: linear-gradient(135deg, #3b82f6, #8b5cf6);
    color: white;
    border-radius: 50%;
    width: 28px; height: 28px;
    display: inline-flex;
    align-items: center; justify-content: center;
    font-weight: 700; font-size: 0.85rem;
    margin-right: 0.5rem;
}

/* Metric cards */
.metric-row { display: flex; gap: 1rem; flex-wrap: wrap; margin: 1rem 0; }
.metric-card {
    background: #1e2433;
    border: 1px solid #2d3a55;
    border-radius: 10px;
    padding: 1rem 1.5rem;
    min-width: 140px;
    flex: 1;
    text-align: center;
}
.metric-value { font-size: 1.8rem; font-weight: 700; }
.metric-label { font-size: 0.78rem; color: #94a3b8; margin-top: 4px; }
.green  { color: #34d399; }
.red    { color: #f87171; }
.yellow { color: #fbbf24; }
.blue   { color: #60a5fa; }
.purple { color: #a78bfa; }

/* Confidence badge */
.badge-high   { background:#166534; color:#bbf7d0; padding:2px 8px; border-radius:999px; font-size:0.75rem; }
.badge-medium { background:#854d0e; color:#fef9c3; padding:2px 8px; border-radius:999px; font-size:0.75rem; }
.badge-low    { background:#7f1d1d; color:#fecaca; padding:2px 8px; border-radius:999px; font-size:0.75rem; }

/* Divider */
.section-divider { border: none; border-top: 1px solid #2d3a55; margin: 1.5rem 0; }

/* Status pills */
.pill-match   { background:#166534; color:#bbf7d0; padding:3px 10px; border-radius:999px; font-size:0.8rem; }
.pill-miss    { background:#7f1d1d; color:#fecaca; padding:3px 10px; border-radius:999px; font-size:0.8rem; }
.pill-missing { background:#713f12; color:#fef9c3; padding:3px 10px; border-radius:999px; font-size:0.8rem; }
.pill-extra   { background:#1e3a5f; color:#bfdbfe; padding:3px 10px; border-radius:999px; font-size:0.8rem; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Session state initialisation
# ---------------------------------------------------------------------------

def _init_state():
    defaults = {
        "step": 1,
        "client_result": None,
        "output_result": None,
        "mapping_result": None,
        "confirmed_mappings": None,
        "key_result": None,
        "output_key_result": None,
        "comparison_result": None,
        "report_bytes": None,
        "report_filename": None,
        "client_file_name": None,
        "output_file_name": None,
        "client_duplicates": None,
        "output_duplicates": None,
        "row_match_result": None,       # legacy row-by-row result
        "simple_match_result": None,    # simple row match result
        "precheck_result": None,        # pre-check tuple result
        "comparison_mode": "advanced",  # 'advanced' or 'simple'
        "precheck_done": False,
        "precheck_time": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _save_uploaded(uploaded) -> str:
    """Save an UploadedFile to a temp file and return the path."""
    suffix = os.path.splitext(uploaded.name)[-1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded.read())
    tmp.flush()
    tmp.close()
    return tmp.name


def _confidence_badge(conf: float) -> str:
    pct = f"{conf*100:.1f}%"
    if conf >= AUTO_ACCEPT_THRESHOLD:
        return f'<span class="badge-high">✓ {pct}</span>'
    if conf >= NEEDS_CONFIRM_THRESHOLD:
        return f'<span class="badge-medium">~ {pct}</span>'
    return f'<span class="badge-low">✗ {pct}</span>'


def _reset_from(step: int):
    """Reset session state from a given step onward."""
    if step <= 2:
        st.session_state.mapping_result = None
        st.session_state.confirmed_mappings = None
    if step <= 3:
        st.session_state.key_result = None
        st.session_state.output_key_result = None
    if step <= 4:
        st.session_state.comparison_result = None
        st.session_state.report_bytes = None
        st.session_state.report_filename = None
        st.session_state.client_duplicates = None
        st.session_state.output_duplicates = None
        st.session_state.row_match_result = None
        st.session_state.simple_match_result = None
        # NOTE: precheck_result and precheck_done are intentionally NOT reset here.
        # The pre-check is independent of the analysis/comparison pipeline and
        # must persist across Analyse Files and Run Comparison clicks.


# ---------------------------------------------------------------------------
# Hero
# ---------------------------------------------------------------------------

st.markdown("""
<div class="hero-header">
  <p class="hero-title">📊 Excel Reconciliation Tool</p>
  <p class="hero-subtitle">
    Upload any two Excel files — the tool automatically maps columns, aligns rows,
    and produces a detailed cell-by-cell comparison report.
  </p>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Step 1 — Upload files
# ---------------------------------------------------------------------------

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
with st.container():
    st.markdown("### 📁 Step 1 — Upload Files")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Client / Reference File**")
        client_upload = st.file_uploader(
            "Upload client file",
            type=["xlsx", "xls", "xlsb"],
            key="client_upload",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown("**Software Output File**")
        output_upload = st.file_uploader(
            "Upload output file",
            type=["xlsx", "xls", "xlsb"],
            key="output_upload",
            label_visibility="collapsed",
        )

    if client_upload and output_upload:
        # Detect if files changed
        if (client_upload.name != st.session_state.client_file_name or
                output_upload.name != st.session_state.output_file_name):
            _reset_from(1)
            st.session_state.client_file_name = client_upload.name
            st.session_state.output_file_name = output_upload.name
            st.session_state.precheck_done = False
            st.session_state.step = 1

        st.success(f"✅ **{client_upload.name}** and **{output_upload.name}** uploaded.")

        # ── Auto pre-check: runs immediately when files are uploaded ─────────
        # Uses raw Excel rows only — no sheet/header detection, no header
        # mapping.  read_excel_raw() picks the best-scoring sheet and returns
        # every row as plain strings without promoting any row to a header.
        # client_result / output_result are NOT set here; they are populated
        # only when the user clicks "Analyse Files".
        if not st.session_state.precheck_done:
            with st.spinner("Running quick row pre-check…"):
                try:
                    # Start timer
                    precheck_start = time.perf_counter()
                    _cp_pc  = _save_uploaded(client_upload)
                    _op_pc  = _save_uploaded(output_upload)
                    _raw_c  = read_excel_raw(_cp_pc)
                    _raw_o  = read_excel_raw(_op_pc)
                    st.session_state.precheck_result = run_simple_row_match(
                        _raw_c, _raw_o
                    )
                     # End timer
                    precheck_end = time.perf_counter()
                    
                    # Store elapsed time
                    st.session_state.precheck_time = precheck_end - precheck_start

                    st.session_state.precheck_done = True
                except Exception as _e:
                    st.warning(f"⚠️ Pre-check skipped: {_e}")
# ============================================================
# PRE-CHECK RESULT — persistent across all steps
# ============================================================

if st.session_state.precheck_result is not None:
    _pc       = st.session_state.precheck_result
    _pc_pct   = _pc["match_pct"]
    _pc_color = "green" if _pc_pct >= 95 else "yellow" if _pc_pct >= 80 else "red"

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("#### 🔍 Pre-Check: Exact Row Match")

    # Show elapsed time (always, regardless of how long it took)
    if st.session_state.precheck_time is not None:
        elapsed = st.session_state.precheck_time
        if elapsed < 60:
            st.info(f"⏱️ Pre-check completed in {elapsed:.2f} seconds")
        else:
            minutes = int(elapsed // 60)
            seconds = elapsed % 60
            st.info(f"⏱️ Pre-check completed in {minutes} min {seconds:.2f} sec")

    st.caption(
        "Every client row is looked up verbatim in the output file using "
        "**exact tuple matching** (whitespace-normalised). "
        "Proceed to **Analyse Files** below for the full column-mapped comparison."
    )

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card">
        <div class="metric-value blue">{_pc['total_human']}</div>
        <div class="metric-label">Client Rows</div>
      </div>
      <div class="metric-card">
        <div class="metric-value blue">{_pc['total_output']}</div>
        <div class="metric-label">Output Rows</div>
      </div>
      <div class="metric-card">
        <div class="metric-value green">{_pc['matched_count']}</div>
        <div class="metric-label">Exact Matches</div>
      </div>
      <div class="metric-card">
        <div class="metric-value red">{_pc['unmatched_count']}</div>
        <div class="metric-label">Unmatched<br><small>(client → not in output)</small></div>
      </div>
      <div class="metric-card">
        <div class="metric-value purple">{_pc['extra_count']}</div>
        <div class="metric-label">Extra<br><small>(output → not in client)</small></div>
      </div>
      <div class="metric-card">
        <div class="metric-value {_pc_color}">{_pc_pct:.1f}%</div>
        <div class="metric-label">Row Match Rate</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    _pct1, _pct2, _pct3 = st.tabs([
        f"✅ Matched ({_pc['matched_count']})",
        f"❌ Unmatched — client NOT in output ({_pc['unmatched_count']})",
        f"➕ Extra — output NOT in client ({_pc['extra_count']})",
    ])
    with _pct1:
        _df_m = pd.DataFrame([
            {"Client Row #": r["human_row"],
             "Output Row #": r["system_rows"][0],
             **{k: v for k, v in r["row_data"].items()}}
            for r in _pc["matched_rows"]
        ])
        if not _df_m.empty:
            st.dataframe(_df_m, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No rows matched exactly.")
    with _pct2:
        _df_u = pd.DataFrame([
            {"Client Row #": r["human_row"],
             **{k: v for k, v in r["row_data"].items()}}
            for r in _pc["unmatched_rows"]
        ])
        if not _df_u.empty:
            st.dataframe(_df_u, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Every client row was found in the output file!")
    with _pct3:
        if not _pc["extra_rows_df"].empty:
            st.dataframe(_pc["extra_rows_df"], use_container_width=True, hide_index=True)
        else:
            st.success("✅ No extra rows in the output file.")
# ---------------------------------------------------------------------------
# Step 2 — Analyse & Map Columns
# ---------------------------------------------------------------------------

if client_upload and output_upload:
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### 🔍 Step 2 — Analyse & Map Columns")


    analyse_col, _ = st.columns([1, 3])
    with analyse_col:
        analyse_btn = st.button("🔎 Analyse Files", type="primary", use_container_width=True)

    if analyse_btn or st.session_state.mapping_result is not None:
        if analyse_btn:
            _reset_from(2)
            with st.spinner("Reading files and detecting columns…"):
                try:
                    client_path = _save_uploaded(client_upload)
                    output_path = _save_uploaded(output_upload)

                    cr = read_excel_file(client_path)
                    orr = read_excel_file(output_path)

                    st.session_state.client_result = cr
                    st.session_state.output_result = orr

                    mr = map_columns(cr.df, orr.df)
                    st.session_state.mapping_result = mr
                    st.session_state.step = 2
                except Exception as e:
                    import traceback
                    st.error(f"❌ Error reading files: {e}")
                    st.code(traceback.format_exc())
                    st.stop()

        cr: ExcelReadResult = st.session_state.client_result
        orr: ExcelReadResult = st.session_state.output_result
        mr: MappingResult = st.session_state.mapping_result

        # File info
        with st.expander("📋 File Detection Details", expanded=False):
            ic1, ic2 = st.columns(2)
            _cr_cols  = [c for c in cr.df.columns  if not c.startswith("_")]
            _orr_cols = [c for c in orr.df.columns if not c.startswith("_")]

            with ic1:
                st.markdown(f"""
                **Client File**
                - Sheet: `{cr.sheet_name}`
                - Header row (0-based): `{cr.header_row_index}`
                - Columns detected: `{len(_cr_cols)}`
                - Data rows: `{len(cr.df)}`
                """)
            with ic2:
                st.markdown(f"""
                **Output File**
                - Sheet: `{orr.sheet_name}`
                - Header row (0-based): `{orr.header_row_index}`
                - Columns detected: `{len(_orr_cols)}`
                - Data rows: `{len(orr.df)}`
                """)

            _row_diff = len(cr.df) - len(orr.df)
            if _row_diff == 0:
                st.success(f"✅ Row count matches — both files have **{len(cr.df)} rows**.")
            else:
                _more = "Client" if _row_diff > 0 else "Output"
                st.warning(
                    f"⚠️ Row count difference: **{abs(_row_diff)} row(s)**"
                    f" — Client: {len(cr.df)}, Output: {len(orr.df)}"
                    f" ({_more} file has more rows)"
                )



        # Unmapped client columns
        if mr.unmapped_client:
            st.warning(f"⚠️ {len(mr.unmapped_client)} client column(s) could not be auto-mapped "
                       f"(score < 50%): "
                       f"`{'`, `'.join(mr.unmapped_client)}`")

        if mr.unmapped_output:
            st.info(f"ℹ️ {len(mr.unmapped_output)} output column(s) had no client match: "
                    f"`{'`, `'.join(mr.unmapped_output)}`")

        # Empty columns — always show if any exist
        _has_empty = mr.empty_client_cols or mr.empty_output_cols
        if _has_empty:
            with st.expander(
                f"🚫 Empty Columns Detected "
                f"({len(mr.empty_client_cols)} client · {len(mr.empty_output_cols)} output)",
                expanded=True,
            ):
                st.markdown(
                    "<small style='color:#94a3b8;'>These columns contain <b>no data</b> "
                    "(all cells are blank or null). They are excluded from auto-mapping "
                    "and comparison but listed here for your awareness.</small>",
                    unsafe_allow_html=True,
                )
                _ec1, _ec2 = st.columns(2)
                with _ec1:
                    st.markdown("**📄 Client File — Empty Columns**")
                    if mr.empty_client_cols:
                        _empty_client_rows = [{"Column Name": c} for c in mr.empty_client_cols]
                        st.dataframe(
                            pd.DataFrame(_empty_client_rows),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.success("✅ No empty columns in client file")
                with _ec2:
                    st.markdown("**📄 Output File — Empty Columns**")
                    if mr.empty_output_cols:
                        _empty_output_rows = [{"Column Name": c} for c in mr.empty_output_cols]
                        st.dataframe(
                            pd.DataFrame(_empty_output_rows),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.success("✅ No empty columns in output file")

        # ── Duplicate-row detection ──────────────────────────────────────────
        _dc_cols = [c for c in cr.df.columns  if not c.startswith("_")]
        _do_cols = [c for c in orr.df.columns if not c.startswith("_")]
        _dup_c = check_duplicate_rows(cr.df[_dc_cols])
        _dup_o = check_duplicate_rows(orr.df[_do_cols])

        _dup_c_count = _dup_c["duplicate_row_count"]
        _dup_o_count = _dup_o["duplicate_row_count"]
        _dup_label = (
            f"{_dup_c_count} client · {_dup_o_count} output"
        )
        with st.expander(
            f"🔁 Duplicate Rows Detected ({_dup_label})",
            expanded=(_dup_c["has_duplicates"] or _dup_o["has_duplicates"]),
        ):
            st.markdown(
                "<small style='color:#94a3b8;'>Rows where <b>every column value</b> is "
                "identical to at least one other row. Duplicates are highlighted here "
                "so you can decide whether to investigate before running comparison.</small>",
                unsafe_allow_html=True,
            )
            _dc1, _dc2 = st.columns(2)

            def _render_dup_panel(col, label: str, dup: dict):
                with col:
                    st.markdown(f"**{label}**")

                    if dup["has_duplicates"]:
                        st.warning(
                            f"⚠️ **{dup['duplicate_row_count']}** "
                            f"duplicate row(s) found "
                            f"({dup['duplicate_occurrence_count']} "
                            f"extra occurrence(s), "
                            f"{dup['unique_rows']} unique rows)"
                            )
                    else:
                        st.success("✅ No duplicate rows found")

            _render_dup_panel(_dc1, "📄 Client File — Duplicates", _dup_c)
            _render_dup_panel(_dc2, "📄 Output File — Duplicates", _dup_o)



if st.session_state.mapping_result is not None:
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### ✏️ Step 3 — Confirm Mappings & Row-Matching Fields")

    mr: MappingResult = st.session_state.mapping_result
    cr: ExcelReadResult = st.session_state.client_result
    orr: ExcelReadResult = st.session_state.output_result

    all_output_cols_norm = ["— skip this column —"] + [
        c for c in orr.df.columns if not c.startswith("_")
    ]

    confirmed: list[ColumnMapping] = list(mr.auto_accepted)

    # Unmapped client columns — allow manual override
    if mr.unmapped_client:
        st.markdown("#### 🔴 Unmapped Client Columns — Manual Mapping (Optional)")
        st.caption(
            "These columns scored below 50% — no confident match was found. "
            "You can manually select the correct output column, or leave as '— skip this column —'."
        )
        for uc in mr.unmapped_client:
            cols = st.columns([3, 3, 4])
            cols[0].markdown(f"**Client:** `{uc}`")
            chosen = cols[1].selectbox(
                f"Map `{uc}` to output column",
                options=all_output_cols_norm,
                index=0,
                key=f"unmap_{uc}",
                label_visibility="collapsed",
            )
            # Find the best-candidate reason for this col
            best_reason = next(
                (m.mapping_reason for m in (mr.needs_confirmation + mr.auto_accepted)
                 if m.client_col == uc),
                ""
            )
            # Also check if there's a cached best-match reason stored on any mapping
            if not best_reason:
                best_reason = "Score < 50% — no confident match found"
            cols[2].caption(best_reason)
            if chosen != "— skip this column —":
                confirmed.append(ColumnMapping(
                    client_col=uc,
                    output_col=chosen,
                    confidence=0.0,
                    status="user_mapped",
                    mapping_reason="Manually mapped by user",
                ))

    st.markdown("---")

    # --- Exclude columns from comparison ---
    st.markdown("#### 🚫 Exclude Columns from Comparison")
    st.caption(
        "Uncheck any columns you do **not** want to compare. "
        "Excluded columns are still used for row matching if selected as a Key Field, "
        "but their cell values will NOT be checked for differences."
    )

    # Build a list of all currently mapped columns to show checkboxes for
    _all_mapped = [m for m in confirmed if m.output_col]
    if _all_mapped:
        _excl_cols = st.columns(min(4, len(_all_mapped)) or 1)
        _excluded: set[str] = set()
        for _i, _m in enumerate(_all_mapped):
            _col_widget = _excl_cols[_i % 4]
            _label = f"`{_m.client_col}`" + (
                f" → `{_m.output_col}`" if _m.client_col != _m.output_col else ""
            )
            _include = _col_widget.checkbox(
                _label,
                value=True,               # default: included
                key=f"excl_{_m.client_col}",
            )
            if not _include:
                _excluded.add(_m.client_col)

        if _excluded:
            st.info(
                f"⏭️ **{len(_excluded)} column(s) excluded from comparison:** "
                + ", ".join(f"`{c}`" for c in sorted(_excluded))
            )

        # Filter confirmed: remove excluded columns
        confirmed = [m for m in confirmed if m.client_col not in _excluded]
    else:
        st.info("No mapped columns yet — map columns above first.")

    st.markdown("---")

    # --- Key field detection ---
    st.markdown("#### 🔑 Row-Matching Key Fields")
    st.caption(
        "Select up to **3 fields** to identify each row. "
        "The tool builds a composite key: **Primary Group → Secondary Sort → Row ID**. "
        "This key is used to match the correct row between the two files."
    )


    client_data_cols = [c for c in cr.df.columns if not c.startswith("_")]
    output_data_cols = [c for c in orr.df.columns if not c.startswith("_")]

    # Detect from client
    kr: KeyFieldResult = detect_key_fields(cr.df)

    none_opt = "— none —"

    # Helper: find mapped output equivalent of a client column
    def _mapped_output_col(client_col: Optional[str]) -> Optional[str]:
        if not client_col:
            return None
        for m in confirmed:
            if m.client_col == client_col and m.output_col:
                return m.output_col
        return None

    def _make_key_row(label_c, label_o, key_c_id, key_o_id,
                      options_c, options_o, default_c, default_o, help_c=""):
        """Render one row of (client selectbox, output selectbox) for a key field."""
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**{label_c}**")
            idx_c = options_c.index(default_c) if default_c in options_c else 0
            sel_c = st.selectbox(
                label_c, options_c, index=idx_c,
                key=key_c_id, label_visibility="collapsed", help=help_c,
            )
        with col2:
            st.markdown(f"**{label_o}**")
            default_o_val = _mapped_output_col(sel_c if sel_c != none_opt else None) or default_o
            idx_o = options_o.index(default_o_val) if default_o_val in options_o else 0
            sel_o = st.selectbox(
                label_o, options_o, index=idx_o,
                key=key_o_id, label_visibility="collapsed",
            )
        return sel_c, sel_o

    opts_c = [none_opt] + client_data_cols
    opts_o = [none_opt] + output_data_cols

    st.markdown("<small>**Client File** &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **Output File**</small>", unsafe_allow_html=True)

    # Field 1 — Primary grouping
    sel_grp_c, sel_grp_o = _make_key_row(
        "🔵 Field 1 — Primary Group (Client)",
        "🔵 Field 1 — Primary Group (Output)",
        "sel_grp_c", "sel_grp_o",
        opts_c, opts_o,
        kr.grouping_col if kr.grouping_col in opts_c else none_opt,
        None,
        help_c=f"Auto-detected: {kr.grouping_col} (score: {kr.grouping_score:.2f})",
    )

    # ↳ Excel-style unique-values filter under Field 1
    if sel_grp_c != none_opt:
        raw_vals_c = sorted(
            cr.df[sel_grp_c].dropna().astype(str).unique().tolist()
        )
        n_total = len(raw_vals_c)
        with st.expander(
            f"🔍 Filter Primary Group values — {n_total} unique value(s) in client file",
            expanded=False,
        ):
            st.caption(
                "Select which primary group values to include in the comparison. "
                "Unselected values will be excluded from both files."
            )
            selected_group_vals = st.multiselect(
                "Select values to include",
                options=raw_vals_c,
                default=raw_vals_c,
                key="grp_multisel",
                label_visibility="collapsed",
            )
            n_sel = len(selected_group_vals)
            if n_sel == n_total:
                st.success(f"✅ All {n_total} values selected")
            elif n_sel == 0:
                st.error("❌ No values selected — comparison will have 0 rows")
            else:
                st.info(f"📌 {n_sel} of {n_total} values selected")
    else:
        selected_group_vals = None  # no filter applied

    # Field 2 — Secondary sort (~11 rows per group)
    sel_sec_c, sel_sec_o = _make_key_row(
        "🟡 Field 2 — Secondary Sort (Client)",
        "🟡 Field 2 — Secondary Sort (Output)",
        "sel_sec_c", "sel_sec_o",
        opts_c, opts_o,
        none_opt, None,
        help_c="A second column that is PART OF the composite key (further narrows the group)",
    )

    # Field 3 — Row identifier (final unique discriminator)
    sel_id_c, sel_id_o = _make_key_row(
        "🟢 Field 3 — Row Identifier (Client)",
        "🟢 Field 3 — Row Identifier (Output)",
        "sel_id_c", "sel_id_o",
        opts_c, opts_o,
        kr.id_col if kr.id_col in opts_c else none_opt,
        None,
        help_c=f"Auto-detected: {kr.id_col} (score: {kr.id_score:.2f})",
    )

    st.info(
        "💡 **How this works:** Rows are matched using **Field 1 + Field 2 + Field 3** as a "
        "composite key. If only Field 1 is set, rows within the same group are matched "
        "by alphabetical order (occurrence #1, #2, #3…). "
        "Add Field 2 or Field 3 to make matching exact."
    )

    # Warn if no unique identifier is set — occurrence-based matching can hide deletions
    _key1_set = sel_grp_c != none_opt
    _key2_set = sel_sec_c != none_opt
    _key3_set = sel_id_c  != none_opt
    if _key1_set and not _key2_set and not _key3_set:
        st.warning(
            "⚠️ **Only Field 1 is set.** Rows within the same group are matched purely by "
            "sorted occurrence order. If a row is deleted from the **middle** of a group, "
            "the rows below it will shift up — the deleted row may appear as a cell difference "
            "rather than a missing row. For accurate deletion detection, set **Field 2** or "
            "**Field 3** to a column that uniquely identifies each row (e.g. product code, date)."
        )

    # Build key-col lists
    client_key_cols = [
        sel_grp_c if sel_grp_c != none_opt else None,
        sel_sec_c if sel_sec_c != none_opt else None,
        sel_id_c  if sel_id_c  != none_opt else None,
    ]
    output_key_cols = [
        sel_grp_o if sel_grp_o != none_opt else None,
        sel_sec_o if sel_sec_o != none_opt else None,
        sel_id_o  if sel_id_o  != none_opt else None,
    ]

    # Guard: must have at least one key field on each side
    if not any(client_key_cols) or not any(output_key_cols):
        st.error(
            "❌ You must select at least one key field (Field 1, 2 or 3) "
            "in both the Client and Output columns to align rows."
        )
        st.stop()

    if not confirmed:
        st.error("❌ No column mappings available. Cannot proceed.")
        st.stop()


# ---------------------------------------------------------------------------
# Step 4 — Tolerance Settings
# ---------------------------------------------------------------------------

if st.session_state.mapping_result is not None:
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### ⚖️ Step 4 — Numeric Tolerance")
    st.caption(
        "tolerance = max(absolute_tolerance, relative_tolerance × max(|client|, |output|)). "
        "Values within tolerance are treated as **MATCH**."
    )

    tc1, tc2 = st.columns(2)
    with tc1:
        abs_tol = st.number_input(
            "Absolute tolerance",
            min_value=0.0,
            value=1e-6,
            format="%.2e",
            step=1e-6,
            help="Minimum absolute difference to flag as mismatch",
        )
    with tc2:
        rel_tol = st.number_input(
            "Relative tolerance",
            min_value=0.0,
            value=1e-4,
            format="%.2e",
            step=1e-5,
            help="Relative tolerance fraction (e.g. 0.0001 = 0.01%)",
        )


# ---------------------------------------------------------------------------
# Step 5 — Run Comparison
# ---------------------------------------------------------------------------

if st.session_state.mapping_result is not None:
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### ▶️ Step 5 — Run Comparison")

    run_col, _ = st.columns([1, 3])
    with run_col:
        run_btn = st.button("🚀 Run Comparison", type="primary", use_container_width=True)

    if run_btn:
        _reset_from(4)
        with st.spinner("Comparing rows…"):
            try:
                # ── Apply primary group filter ────────────────────────────────
                if sel_grp_c != none_opt and selected_group_vals is not None:
                    sel_set = set(selected_group_vals)
                    df_c_use = cr.df[
                        cr.df[sel_grp_c].astype(str).isin(sel_set)
                    ].reset_index(drop=True)
                    if sel_grp_o != none_opt:
                        df_o_use = orr.df[
                            orr.df[sel_grp_o].astype(str).isin(sel_set)
                        ].reset_index(drop=True)
                    else:
                        df_o_use = orr.df
                else:
                    df_c_use = cr.df
                    df_o_use = orr.df

                # ── Key-field alignment + cell-by-cell comparison ───────────────────
                aligned = align_rows(
                    df_c_use, df_o_use,
                    client_key_cols=client_key_cols,
                    output_key_cols=output_key_cols,
                )
                result = compare_dataframes(
                    aligned=aligned,
                    column_mappings=confirmed,
                    df_client_orig=cr.df,
                    df_output_orig=orr.df,
                    abs_tol=abs_tol,
                    rel_tol=rel_tol,
                )
                st.session_state.comparison_result = result
                tmp_dir = tempfile.gettempdir()
                report_path = generate_report(result, output_dir=tmp_dir)
                with open(report_path, "rb") as f:
                    st.session_state.report_bytes = f.read()
                st.session_state.report_filename = os.path.basename(report_path)

                # ── Duplicate row detection ─────────────────────────────────────────
                def _find_full_row_duplicates(df: pd.DataFrame) -> pd.DataFrame:
                    """Return all duplicate rows with a Duplicate Count column."""
                    data_cols = [c for c in df.columns if not c.startswith("_")]
                    df_data = df[data_cols].copy().reset_index(drop=True)
                    df_str = df_data.astype(str)
                    row_tuples = [tuple(r) for r in df_str.values]
                    from collections import Counter
                    counts = Counter(row_tuples)
                    df_data["Duplicate Count"] = [counts[t] for t in row_tuples]
                    df_data = df_data[df_data["Duplicate Count"] > 1].copy()
                    cols = ["Duplicate Count"] + [c for c in df_data.columns if c != "Duplicate Count"]
                    return df_data[cols].reset_index(drop=True)

                st.session_state.client_duplicates = _find_full_row_duplicates(df_c_use)
                st.session_state.output_duplicates = _find_full_row_duplicates(df_o_use)

                st.session_state.step = 5
            except Exception as e:
                st.error(f"❌ Comparison failed: {e}")
                import traceback
                st.code(traceback.format_exc())
                st.stop()


# ---------------------------------------------------------------------------
# Step 6 — Results
# ---------------------------------------------------------------------------

if st.session_state.comparison_result is not None:
    result: ComparisonResult = st.session_state.comparison_result
    s = result.summary

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### 📊 Step 6 — Results")

    # ── Summary metric cards ───────────────────────────────────────────────
    _cell_match_color = 'green' if s.match_pct >= 95 else 'yellow' if s.match_pct >= 80 else 'red'
    _blank_diffs = s.missing_cell_values if hasattr(s, 'missing_cell_values') else 0
    _value_diffs = s.mismatched_cells - _blank_diffs if s.mismatched_cells > _blank_diffs else s.mismatched_cells
    _blank_diffs  = s.mismatched_cells - _value_diffs

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card">
        <div class="metric-value blue">{s.total_client_rows}</div>
        <div class="metric-label">Client Rows</div>
      </div>
      <div class="metric-card">
        <div class="metric-value blue">{s.total_output_rows}</div>
        <div class="metric-label">Output Rows</div>
      </div>
      <div class="metric-card">
        <div class="metric-value green">{s.matched_rows}</div>
        <div class="metric-label">Matched Rows</div>
      </div>
      <div class="metric-card">
        <div class="metric-value red">{s.mismatched_rows}</div>
        <div class="metric-label">Rows with<br>Differences</div>
      </div>
      <div class="metric-card">
        <div class="metric-value yellow">{s.missing_rows}</div>
        <div class="metric-label">Missing Rows<br><small>(in client, not in output)</small></div>
      </div>
      <div class="metric-card">
        <div class="metric-value purple">{s.extra_rows}</div>
        <div class="metric-label">Extra Rows<br><small>(in output, not in client)</small></div>
      </div>
      <div class="metric-card">
        <div class="metric-value {_cell_match_color}">{s.match_pct:.1f}%</div>
        <div class="metric-label">Cell Match Rate</div>
      </div>
      <div class="metric-card">
        <div class="metric-value orange">{s.mismatched_cells}</div>
        <div class="metric-label">Total Cell<br>Differences<br><small>({_value_diffs} value · {_blank_diffs} blank)</small></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Download button ────────────────────────────────────────────────────
    if st.session_state.report_bytes:
        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        dl_col, _ = st.columns([1, 3])
        with dl_col:
            st.download_button(
                label="⬇️ Download Excel Report",
                data=st.session_state.report_bytes,
                file_name=st.session_state.report_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        st.caption(
            f"📄 Report: `{st.session_state.report_filename}` — "
            "contains **Summary**, **Column Mapping**, **Row Comparison**, "
            "**Cell Differences**, **Missing & Extra Rows** sheets."
        )

    # ── Tabs ───────────────────────────────────────────────────────────────
    _c_dups = st.session_state.client_duplicates
    _o_dups = st.session_state.output_duplicates
    _c_dup_count = len(_c_dups) if _c_dups is not None and not _c_dups.empty else 0
    _o_dup_count = len(_o_dups) if _o_dups is not None and not _o_dups.empty else 0

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        f"📋 Row Results ({s.matched_rows + s.mismatched_rows})",
        f"🔬 Cell Differences ({s.mismatched_cells})",
        f"⚠️ Missing Rows ({s.missing_rows}) — in client, NOT in output",
        f"➕ Extra Rows ({s.extra_rows}) — in output, NOT in client",
        f"🔁 Duplicates — Client ({_c_dup_count})",
        f"🔁 Duplicates — Output ({_o_dup_count})",
    ])

    with tab1:
        st.caption("All matched and mismatched rows (rows that exist in both files). Green = match, red = differences found.")
        row_df = pd.DataFrame([vars(rr) for rr in result.row_results]) if result.row_results else pd.DataFrame()
        if not row_df.empty:
            st.dataframe(row_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No row results to display.")

    with tab2:
        st.caption("Every cell that differed between the two files, showing client value vs output value.")
        cell_df = pd.DataFrame([vars(cd) for cd in result.cell_diffs]) if result.cell_diffs else pd.DataFrame()
        if not cell_df.empty:
            st.dataframe(cell_df, use_container_width=True, hide_index=True)
        else:
            st.success("✅ No cell differences — all compared cells match!")

    with tab3:
        st.caption("Rows present in the **client file** but not found in the output file.")
        if not result.missing_rows_df.empty:
            st.dataframe(result.missing_rows_df, use_container_width=True, hide_index=True)
        else:
            st.success("✅ No missing rows.")

    with tab4:
        st.caption("Rows present in the **output file** but not found in the client file.")
        if not result.extra_rows_df.empty:
            st.dataframe(result.extra_rows_df, use_container_width=True, hide_index=True)
        else:
            st.success("✅ No extra rows.")

    with tab5:
        st.caption("🔁 **Duplicate Rows — Client File**: rows where every column value is identical to another row.")
        if _c_dups is not None and not _c_dups.empty:
            st.warning(f"⚠️ **{_c_dup_count} duplicate row(s)** found in the client file.")
            st.dataframe(_c_dups, use_container_width=True, hide_index=True)
        else:
            st.success("✅ No duplicate rows found in the client file.")

    with tab6:
        st.caption("🔁 **Duplicate Rows — Output File**: rows where every column value is identical to another row.")
        if _o_dups is not None and not _o_dups.empty:
            st.warning(f"⚠️ **{_o_dup_count} duplicate row(s)** found in the output file.")
            st.dataframe(_o_dups, use_container_width=True, hide_index=True)
        else:
            st.success("✅ No duplicate rows found in the output file.")

    # Re-run button
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    if st.button("🔄 Start Over", use_container_width=False):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
