import pandas as pd
from pathlib import Path
from datetime import datetime


# ---------------------------------------------------------------------------
# Duplicate-row detection  (called by app.py after Analyse Files)
# ---------------------------------------------------------------------------

def check_duplicate_rows(df: pd.DataFrame) -> dict:
    """
    Detect completely duplicated rows in a DataFrame.

    Returns only duplicate statistics, not the duplicate rows.
    """
    total_rows = len(df)

    duplicate_mask = df.duplicated(keep=False)

    duplicate_row_count = int(duplicate_mask.sum())

    unique_rows = int(df.drop_duplicates().shape[0])

    duplicate_occurrence_count = total_rows - unique_rows

    return {
        "has_duplicates": duplicate_row_count > 0,
        "duplicate_row_count": duplicate_row_count,
        "duplicate_occurrence_count": duplicate_occurrence_count,
        "unique_rows": unique_rows,
    }
# ---------------------------------------------------------------------------
# DataFrame-based pre-check  (called by app.py — no file I/O)
# ---------------------------------------------------------------------------

def run_simple_row_match(
    human_df: pd.DataFrame,
    system_df: pd.DataFrame,
    selected_columns: list | None = None,
) -> dict:
    """
    Tuple-based row pre-check: compare every client row against the output
    file using exact whole-row matching after normalisation.

    Returns a dict with:
        matched_rows   – list of {human_row, system_rows, row_data}
        unmatched_rows – list of {human_row, row_data}  (missing from output)
        extra_rows_df  – pd.DataFrame  (output rows not matched by any client row)
        total_human    – int
        total_output   – int
        matched_count  – int
        unmatched_count– int
        extra_count    – int
        match_pct      – float  (0–100)
        used_columns   – list[str]
        human_columns  – list[str]
        system_columns – list[str]
    """

    human_df  = human_df.copy()
    system_df = system_df.copy()
    human_df.columns  = [str(c).strip() for c in human_df.columns]
    system_df.columns = [str(c).strip() for c in system_df.columns]

    human_columns  = list(human_df.columns)
    system_columns = list(system_df.columns)

    use_human_cols  = selected_columns if selected_columns else human_columns
    use_system_cols = [c for c in use_human_cols if c in system_columns]
    if not use_system_cols:
        use_system_cols = system_columns
        use_human_cols  = human_columns

    def _norm(v):
        return "" if pd.isna(v) else str(v).strip()

    human_norm  = human_df[use_human_cols].map(_norm)
    system_norm = system_df[use_system_cols].map(_norm)

    # Build lookup: tuple → [df-index, …]
    sys_lookup: dict[tuple, list[int]] = {}
    for idx, row in system_norm.iterrows():
        sys_lookup.setdefault(tuple(row.tolist()), []).append(int(idx))

    consumed: set[int] = set()
    matched_rows:   list[dict] = []
    unmatched_rows: list[dict] = []

    for h_idx, row in human_norm.iterrows():
        key         = tuple(row.tolist())
        human_excel = int(h_idx) + 2
        row_data    = {c: human_df.iloc[int(h_idx)][c] for c in human_columns}
        candidates  = sys_lookup.get(key, [])

        chosen = next((i for i in candidates if i not in consumed), None)
        if chosen is not None:
            consumed.add(chosen)
            matched_rows.append({
                "human_row":   human_excel,
                "system_rows": [chosen + 2],
                "row_data":    row_data,
            })
        else:
            unmatched_rows.append({"human_row": human_excel, "row_data": row_data})

    # Extra output rows (not consumed by any client match)
    extra = [
        {c: system_df.iloc[i][c] for c in system_columns}
        for i in range(len(system_df)) if i not in consumed
    ]
    extra_rows_df = pd.DataFrame(extra) if extra else pd.DataFrame(columns=system_columns)

    total   = len(human_df)
    matched = len(matched_rows)
    unmatched = len(unmatched_rows)

    return {
        "matched_rows":    matched_rows,
        "unmatched_rows":  unmatched_rows,
        "extra_rows_df":   extra_rows_df,
        "total_human":     total,
        "total_output":    len(system_df),
        "matched_count":   matched,
        "unmatched_count": unmatched,
        "extra_count":     len(extra),
        "match_pct":       (matched / total * 100) if total else 0.0,
        "used_columns":    use_human_cols,
        "human_columns":   human_columns,
        "system_columns":  system_columns,
    }


# ---------------------------------------------------------------------------
# File-path-based standalone script  (unchanged)
# ---------------------------------------------------------------------------


def compare_human_system_excel_rows(
    human_excel_path: str,
    system_excel_path: str,
    output_folder: str = "output"
) -> str:
    """
    Compare every row in Human Excel against System Excel.

    Human Excel:
        The reference file. Every Human row is checked.

    System Excel:
        The file searched for matching Human rows.

    A row is considered a match only when all column values
    are equal after normalization.
    """

    # --------------------------------------------------------
    # READ EXCEL FILES
    # --------------------------------------------------------

    human_df = pd.read_excel(
        human_excel_path,
        header=0
    )

    system_df = pd.read_excel(
        system_excel_path,
        header=0
    )

    # --------------------------------------------------------
    # NORMALIZE COLUMN HEADERS
    # --------------------------------------------------------

    human_df.columns = [
        str(column).strip()
        for column in human_df.columns
    ]

    system_df.columns = [
        str(column).strip()
        for column in system_df.columns
    ]

    # --------------------------------------------------------
    # CHECK COLUMNS
    # --------------------------------------------------------

    human_columns = list(human_df.columns)
    system_columns = list(system_df.columns)

    # if human_columns != system_columns:

    #     missing_in_system = sorted(
    #         set(human_columns) - set(system_columns)
    #     )

    #     missing_in_human = sorted(
    #         set(system_columns) - set(human_columns)
    #     )

    #     raise ValueError(
    #         "Excel column structure is different.\n"
    #         f"Missing in System: {missing_in_system}\n"
    #         f"Missing in Human: {missing_in_human}"
    #     )

    # --------------------------------------------------------
    # NORMALIZE CELL VALUES
    # --------------------------------------------------------

    def normalize_value(value):

        # Handle NaN / None
        if pd.isna(value):
            return ""

        # Convert to string
        value = str(value)

        # Remove leading/trailing spaces
        value = value.strip()

        # Treat empty/space values as empty
        if value == "":
            return ""

        return value

    # Create normalized copies
    human_normalized = human_df.map(normalize_value)
    system_normalized = system_df.map(normalize_value)

    # --------------------------------------------------------
    # CREATE SYSTEM ROW LOOKUP
    # --------------------------------------------------------

    system_row_lookup = {}

    for system_index, row in system_normalized.iterrows():

        row_tuple = tuple(row.tolist())

        if row_tuple not in system_row_lookup:
            system_row_lookup[row_tuple] = []

        system_row_lookup[row_tuple].append(
            system_index + 2
        )
        # +2 because Excel row 1 is header
    print(system_row_lookup)
    # --------------------------------------------------------
    # COMPARE HUMAN ROWS
    # --------------------------------------------------------

    matched_rows = []
    unmatched_rows = []

    for human_index, row in human_normalized.iterrows():

        row_tuple = tuple(row.tolist())

        human_excel_row = human_index + 2

        system_matches = system_row_lookup.get(
            row_tuple,
            []
        )
        print(system_matches)

        if system_matches:

            matched_rows.append({
                "human_row": human_excel_row,
                "system_rows": system_matches
            })

        else:

            # Store complete Human row
            row_data = {
                column: human_df.iloc[human_index][column]
                for column in human_columns
            }

            unmatched_rows.append({
                "human_row": human_excel_row,
                "data": row_data
            })

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    total_human_rows = len(human_df)

    matched_count = len(matched_rows)

    unmatched_count = len(unmatched_rows)

    match_percentage = (
        (matched_count / total_human_rows) * 100
        if total_human_rows > 0
        else 0
    )

    # --------------------------------------------------------
    # CREATE MARKDOWN REPORT
    # --------------------------------------------------------

    md_content = []

    md_content.append(
        "# Human vs System Excel Row Comparison\n"
    )

    md_content.append(
        f"**Generated On:** "
        f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    )

    md_content.append("---\n")

    # --------------------------------------------------------
    # FILE INFORMATION
    # --------------------------------------------------------

    md_content.append("## 1. File Information\n")

    md_content.append("| Type | Excel File |")
    md_content.append("|---|---|")

    md_content.append(
        f"| Human | `{Path(human_excel_path).name}` |"
    )

    md_content.append(
        f"| System | `{Path(system_excel_path).name}` |"
    )

    md_content.append("")

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    md_content.append("## 2. Comparison Summary\n")

    md_content.append(
        "| Metric | Result |"
    )

    md_content.append(
        "|---|---:|"
    )

    md_content.append(
        f"| Total Human Rows | `{total_human_rows}` |"
    )

    md_content.append(
        f"| Matched Rows | `{matched_count}` |"
    )

    md_content.append(
        f"| Unmatched Rows | `{unmatched_count}` |"
    )

    md_content.append(
        f"| Match Percentage | `{match_percentage:.2f}%` |"
    )

    md_content.append("")

    # --------------------------------------------------------
    # OVERALL STATUS
    # --------------------------------------------------------

    if unmatched_count == 0:

        md_content.append(
            "### ✅ ALL HUMAN ROWS FOUND IN SYSTEM\n"
        )

    else:

        md_content.append(
            "### ❌ SOME HUMAN ROWS NOT FOUND IN SYSTEM\n"
        )

    # --------------------------------------------------------
    # MATCHED ROWS
    # --------------------------------------------------------

    md_content.append("## 3. Matched Rows\n")

    if matched_rows:

        md_content.append(
            "| Human Excel Row | Matching System Row(s) |"
        )

        md_content.append(
            "|---:|---|"
        )

        for match in matched_rows:

            system_rows_text = ", ".join(
                str(row_number)
                for row_number in match["system_rows"]
            )

            md_content.append(
                f"| `{match['human_row']}` | "
                f"`{system_rows_text}` |"
            )

    else:

        md_content.append(
            "No matching rows were found."
        )

    md_content.append("")

    # --------------------------------------------------------
    # UNMATCHED ROWS
    # --------------------------------------------------------

    md_content.append(
        "## 4. Human Rows Not Found in System\n"
    )

    if unmatched_rows:

        md_content.append(
            f"**Total Unmatched Rows:** "
            f"`{unmatched_count}`\n"
        )

        # for item in unmatched_rows:

        #     md_content.append(
        #         f"### Human Excel Row {item['human_row']}\n"
        #     )

        #     md_content.append("| Column | Value |")
        #     md_content.append("|---|---|")

        #     for column, value in item["data"].items():

        #         if pd.isna(value):
        #             value = ""

        #         value = str(value).strip()

        #         md_content.append(
        #             f"| `{column}` | `{value}` |"
        #         )

        #     md_content.append("")

    else:

        md_content.append(
            "### ✅ No Unmatched Human Rows\n"
        )

        md_content.append(
            "Every Human Excel row was found in the System Excel."
        )

    # --------------------------------------------------------
    # SAVE REPORT
    # --------------------------------------------------------

    now = datetime.now()

    date_folder = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%Y%m%d_%H%M%S")

    output_dir = (
        Path(output_folder)
        / date_folder
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    md_path = (
        output_dir
        / f"human_vs_system_row_comparison_{timestamp}.md"
    )

    md_path.write_text(
        "\n".join(md_content),
        encoding="utf-8"
    )

    return str(md_path)
 
 
if __name__ == "__main__":

    human_excel = r"C:\Users\anush\OneDrive\Documents\project_folders\excel_comparison\output (16) (1).xlsx"
    system_excel = r"C:\Users\anush\OneDrive\Documents\project_folders\excel_comparison\DSS 202603 Final (1).xlsb"

    output = compare_human_system_excel_rows(
        human_excel_path=human_excel,
        system_excel_path=system_excel
    )

    print("Comparison completed.")
    print("Report saved at:", output)