# Excel Reconciliation Tool

A Streamlit application that compares two Excel files cell-by-cell and produces
a colour-coded Excel report highlighting every difference.

---

## Features

- **Auto column mapping** - matches columns between two files using a 5-factor
  scoring system (name similarity, synonyms, data type, value overlap, cardinality)
- **Smart row alignment** - builds a composite key from up to 3 user-selected
  fields to align rows across files, even when in different orders
- **Full data-type normalisation** - handles integers, floats, scientific notation,
  dates, booleans, and strings uniformly before comparison
- **Numeric tolerance** - configurable absolute and relative tolerance for
  floating-point comparisons
- **Exclude columns** - selectively skip columns from cell comparison while still
  using them for row matching
- **Group filtering** - Excel-style multiselect to narrow comparison to specific
  primary groups
- **Detailed results** - row-level status (MATCH / MISMATCH / MISSING / EXTRA)
  and cell-level diffs with numeric difference and percentage
- **Downloadable report** - 5-sheet colour-coded Excel workbook

---

## Installation and Setup

### Prerequisites

- Python 3.11+

### Step 1 - Create a virtual environment

```bash
python -m venv env
.\env\Scripts\activate
```

### Step 2 - Install dependencies

```bash
pip install -r requirements.txt
```

### Step 3 - Run the app

```bash
streamlit run app.py
```

The app opens automatically at http://localhost:8501

---

## Usage

| Step | Action |
|------|--------|
| 1 | Upload the **Client** file and the **Output** file (.xlsx, .xls, or .xlsb) |
| 2 | Click **Analyse Headers** - columns are auto-mapped |
| 3 | Confirm ambiguous mappings, manually map unmapped columns, exclude columns you do not want compared, and set 3 row-matching key fields |
| 4 | Set numeric tolerance (default: 1e-6 absolute, 0.01% relative) |
| 5 | Click **Run Comparison** |
| 6 | Review the results and download the Excel report |

### Row-Matching Key Fields

| Field | Role |
|-------|------|
| Field 1 - Primary Group | Groups rows into broad categories (e.g. distributor, region) |
| Field 2 - Secondary Sort | Narrows within each group |
| Field 3 - Row Identifier | Final unique discriminator per row (e.g. product code, date) |

> Note: If only Field 1 is set, rows are matched by sorted occurrence order.
> For accurate deletion detection, always set Field 2 or Field 3 to a column
> that uniquely identifies each row within a group.

---

## Output Report Sheets

| Sheet | Contents |
|-------|----------|
| Summary | Total counts, match %, tolerances used |
| Column Mapping | All client to output column mappings with confidence scores |
| Row Comparison | Every row with MATCH / MISMATCH / MISSING / EXTRA status |
| Cell Differences | Each differing cell: column, client value, output value, numeric diff, diff % |
| Missing and Extra Rows | Full row data for unmatched rows |

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| streamlit | >= 1.35 | Web UI framework |
| pandas | >= 2.0 | DataFrame operations |
| numpy | >= 1.26 | Vectorized numeric comparison |
| openpyxl | >= 3.1 | Read/write .xlsx, report formatting |
| xlrd | >= 2.0 | Read legacy .xls files |
| pyxlsb | >= 1.0 | Read binary .xlsb files |
| rapidfuzz | >= 3.6 | Fuzzy string matching for column name scoring |

---

## Project Structure

```
excel_comparison/
├── app.py               Streamlit UI - all 6 workflow steps
├── requirements.txt     Python dependencies
└── reconciler/
    ├── __init__.py
    ├── file_reader.py       Auto-detects best sheet and header row
    ├── header_mapper.py     5-factor column similarity scoring
    ├── row_matcher.py       Composite key building and row alignment
    ├── comparator.py        Vectorized cell-by-cell comparison
    └── report_generator.py  5-sheet colour-coded Excel report
```
