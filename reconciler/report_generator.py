"""
report_generator.py
-------------------
Generates a timestamped, colour-coded Excel comparison report with 5 sheets:
  1. Summary
  2. Column Mapping
  3. Row Comparison
  4. Cell Differences
  5. Missing & Extra Rows
"""

from __future__ import annotations

import datetime
import os
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Font,
    PatternFill,
    Border,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from reconciler.comparator import ComparisonResult


# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------

_GREEN  = PatternFill("solid", fgColor="C6EFCE")   # matched
_RED    = PatternFill("solid", fgColor="FFC7CE")   # mismatch / extra
_YELLOW = PatternFill("solid", fgColor="FFEB9C")   # missing / warning
_BLUE   = PatternFill("solid", fgColor="BDD7EE")   # header
_GREY   = PatternFill("solid", fgColor="D9D9D9")   # sub-header
_WHITE  = PatternFill("solid", fgColor="FFFFFF")
_ORANGE = PatternFill("solid", fgColor="FFD966")   # low confidence

_HEADER_FONT = Font(bold=True, color="1F3864")
_TITLE_FONT  = Font(bold=True, size=14, color="1F3864")

_THIN_BORDER_SIDE = Side(style="thin", color="BFBFBF")
_THIN_BORDER = Border(
    left=_THIN_BORDER_SIDE, right=_THIN_BORDER_SIDE,
    top=_THIN_BORDER_SIDE,  bottom=_THIN_BORDER_SIDE,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _write_header_row(ws, row_idx: int, values: list, fill=_BLUE):
    for col_idx, val in enumerate(values, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.fill = fill
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _THIN_BORDER


def _write_data_row(ws, row_idx: int, values: list, fill=None):
    for col_idx, val in enumerate(values, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        if fill:
            cell.fill = fill
        cell.alignment = Alignment(vertical="center", wrap_text=False)
        cell.border = _THIN_BORDER


def _autofit_columns(ws, min_width: int = 10, max_width: int = 50):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)


def _freeze_top_row(ws):
    ws.freeze_panes = ws.cell(row=2, column=1)


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def _build_summary_sheet(ws, result: ComparisonResult):
    s = result.summary
    ws.title = "Summary"

    # Title
    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = "Excel Reconciliation — Summary"
    title_cell.font = _TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = _BLUE
    ws.row_dimensions[1].height = 28

    metrics = [
        ("", "", "", ""),
        ("Row Statistics", "", "", ""),
        ("Client rows", s.total_client_rows, "Output rows", s.total_output_rows),
        ("Matched rows", s.matched_rows, "Mismatched rows", s.mismatched_rows),
        ("Missing rows (client not in output)", s.missing_rows,
         "Extra rows (output not in client)", s.extra_rows),
        ("", "", "", ""),
        ("Cell Statistics", "", "", ""),
        ("Total cells compared", s.total_cells_compared,
         "Cell match %", f"{s.match_pct:.2f}%"),
        ("Matched cells", s.matched_cells,
         "Mismatched cells", s.mismatched_cells),
        ("Missing cell values", s.missing_cell_values, "", ""),
        ("", "", "", ""),
        ("Tolerance Settings", "", "", ""),
        ("Absolute tolerance", s.abs_tolerance, "Relative tolerance", s.rel_tolerance),
    ]

    sub_header_rows = {3, 8, 13}

    for row_offset, row_data in enumerate(metrics, start=2):
        row_idx = row_offset
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = _THIN_BORDER
            if row_offset in sub_header_rows and col_idx == 1:
                cell.font = Font(bold=True, color="1F3864")
                cell.fill = _GREY
            cell.alignment = Alignment(vertical="center")

    _autofit_columns(ws)


def _build_column_mapping_sheet(ws, result: ComparisonResult):
    ws.title = "Column Mapping"
    headers = [
        "Client Column", "Output Column", "Confidence %",
        "Name Score", "Synonym Score", "Type Score",
        "Overlap Score", "Uniqueness Score", "Status",
    ]
    _write_header_row(ws, 1, headers)
    _freeze_top_row(ws)

    all_mappings = result.column_mappings
    for row_idx, m in enumerate(all_mappings, start=2):
        conf_pct = f"{m.confidence * 100:.1f}%"
        status = getattr(m, "status", "auto")
        if m.output_col is None:
            fill = _YELLOW
            status = "unmapped"
        elif m.confidence >= 0.80:
            fill = _GREEN
        elif m.confidence >= 0.50:
            fill = _ORANGE
        else:
            fill = _RED

        _write_data_row(ws, row_idx, [
            m.client_col,
            m.output_col or "— unmapped —",
            conf_pct,
            f"{m.name_score:.2f}",
            f"{m.synonym_score:.2f}",
            f"{m.type_score:.2f}",
            f"{m.overlap_score:.2f}",
            f"{m.unique_score:.2f}",
            status.upper(),
        ], fill=fill)

    _autofit_columns(ws)


def _build_row_comparison_sheet(ws, result: ComparisonResult):
    ws.title = "Row Comparison"
    headers = [
        "Row Key", "Client Excel Row", "Output Excel Row",
        "Status", "Mismatched Columns",
    ]
    _write_header_row(ws, 1, headers)
    _freeze_top_row(ws)

    status_fill = {
        "MATCH":    _GREEN,
        "MISMATCH": _RED,
        "MISSING":  _YELLOW,
        "EXTRA":    _ORANGE,
    }

    for row_idx, rr in enumerate(result.row_results, start=2):
        fill = status_fill.get(rr.status, _WHITE)
        _write_data_row(ws, row_idx, [
            rr.row_key,
            rr.client_excel_row or "",
            rr.output_excel_row or "",
            rr.status,
            ", ".join(rr.mismatch_columns) if rr.mismatch_columns else "",
        ], fill=fill)

    _autofit_columns(ws)


def _build_cell_differences_sheet(ws, result: ComparisonResult):
    ws.title = "Cell Differences"
    headers = [
        "Row Key",
        "Client Excel Row", "Output Excel Row",
        "Column",
        "Client Value", "Output Value",
        "Numeric Diff", "Diff %",
        "Issue",
    ]
    _write_header_row(ws, 1, headers)
    _freeze_top_row(ws)

    issue_fill = {
        "MISMATCH":       _RED,
        "MISSING_CLIENT": _YELLOW,
        "MISSING_OUTPUT": _YELLOW,
    }

    for row_idx, cd in enumerate(result.cell_diffs, start=2):
        fill = issue_fill.get(cd.issue, _RED)
        ndiff = f"{cd.numeric_diff:.6g}" if cd.numeric_diff is not None else ""
        dpct  = f"{cd.diff_pct:.4f}%" if cd.diff_pct is not None else ""
        _write_data_row(ws, row_idx, [
            cd.row_key,
            cd.client_excel_row or "",
            cd.output_excel_row or "",
            cd.column_name,
            cd.client_value,
            cd.output_value,
            ndiff,
            dpct,
            cd.issue,
        ], fill=fill)

    _autofit_columns(ws)


def _build_missing_extra_sheet(ws, result: ComparisonResult):
    ws.title = "Missing & Extra Rows"

    current_row = 1

    def _write_section(title: str, df: pd.DataFrame, fill):
        nonlocal current_row
        # Section title
        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True, size=12)
        ws.cell(row=current_row, column=1).fill = _GREY
        current_row += 1

        if df.empty:
            ws.cell(row=current_row, column=1, value="(none)").font = Font(italic=True)
            current_row += 2
            return

        # Drop internal columns for display
        display_df = df.drop(
            columns=[c for c in df.columns if c.startswith("_")],
            errors="ignore",
        )

        # Header
        _write_header_row(ws, current_row, display_df.columns.tolist(), fill=_BLUE)
        current_row += 1

        for _, row in display_df.iterrows():
            _write_data_row(ws, current_row, row.tolist(), fill=fill)
            current_row += 1

        current_row += 1  # blank spacer row

    _write_section(
        f"Missing Rows ({len(result.missing_rows_df)}) — In Client, Not in Output",
        result.missing_rows_df, _YELLOW,
    )
    _write_section(
        f"Extra Rows ({len(result.extra_rows_df)}) — In Output, Not in Client",
        result.extra_rows_df, _ORANGE,
    )

    _autofit_columns(ws)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_report(
    result: ComparisonResult,
    output_dir: str = ".",
    base_name: str = "reconciliation_report",
) -> str:
    """
    Write a timestamped Excel report and return the file path.

    Parameters
    ----------
    result      : ComparisonResult from comparator.compare_dataframes()
    output_dir  : directory to write the report into
    base_name   : prefix for the filename

    Returns
    -------
    Absolute path to the generated .xlsx file
    """
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{base_name}_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Build each sheet
    _build_summary_sheet(wb.create_sheet(), result)
    _build_column_mapping_sheet(wb.create_sheet(), result)
    _build_row_comparison_sheet(wb.create_sheet(), result)
    _build_cell_differences_sheet(wb.create_sheet(), result)
    _build_missing_extra_sheet(wb.create_sheet(), result)

    wb.save(filepath)
    return os.path.abspath(filepath)
