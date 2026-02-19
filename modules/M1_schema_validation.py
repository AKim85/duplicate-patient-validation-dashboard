"""
Module 1 — Load & Validate Schema (Data Readiness)

Purpose:
    Validate the structure and basic data types of the supplier-provided
    Duplicate Patient Summary before any downstream processing.

MVP link:
    MVP #1–2 (load data; schema and type validation)
"""

import pandas as pd
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Border, Side

# -----------------------------------------------------
# Configuration
# -----------------------------------------------------

FILE_PATH = Path("DUPLICATE_PATIENTS.xlsx")
INPUT_SHEET = "Input_DuplicatePatientSummary"
OUTPUT_SHEET = "M1_SchemaValidation"

EXPECTED_COLUMNS = [
    "ID 1",
    "ID 2",
    "Patient 1",
    "Patient 2",
    "Identifier 1",
    "Identifier 2",
    "Recent Edit",
    "Matches",
    "Forename Match",
    "Surname Match",
    "Soundex Match",
    "Sex Match",
    "DOB Match",
    "PostCode Match",
    "Address Match",
]

# -----------------------------------------------------
# Load input
# -----------------------------------------------------

xls = pd.ExcelFile(FILE_PATH)

if INPUT_SHEET not in xls.sheet_names:
    raise ValueError(
        f"Expected sheet '{INPUT_SHEET}' not found. "
        f"Available sheets: {xls.sheet_names}"
    )

df = pd.read_excel(FILE_PATH, sheet_name=INPUT_SHEET)

# -----------------------------------------------------
# Validation checks
# -----------------------------------------------------

# 1. Schema validation
missing_columns = [c for c in EXPECTED_COLUMNS if c not in df.columns]

# 2. Date parsing
invalid_dates = 0
if "Recent Edit" in df.columns:
    parsed_dates = pd.to_datetime(df["Recent Edit"], errors="coerce")
    invalid_dates = parsed_dates.isna().sum()

# 3. Numeric checks
numeric_columns = [
    c for c in EXPECTED_COLUMNS if c.endswith("Match") or c == "Matches"
]

non_numeric_counts = {}
for col in numeric_columns:
    if col in df.columns:
        coerced = pd.to_numeric(df[col], errors="coerce")
        non_numeric_counts[col] = coerced.isna().sum()

# -----------------------------------------------------
# Determine PASS / FAIL
# -----------------------------------------------------

status = "PASS" if len(missing_columns) == 0 else "FAIL"

# -----------------------------------------------------
# Build validation summary
# -----------------------------------------------------

summary_rows = [
    {
        "Check": "Missing columns",
        "Result": len(missing_columns),
        "Details": ", ".join(missing_columns),
    },
    {
        "Check": "Invalid dates (Recent Edit)",
        "Result": invalid_dates,
        "Details": "",
    },
]

for col, count in non_numeric_counts.items():
    summary_rows.append(
        {
            "Check": f"Non-numeric values in {col}",
            "Result": count,
            "Details": "",
        }
    )

summary_df = pd.DataFrame(summary_rows)

# -----------------------------------------------------
# Write output to Excel (formatted)
# -----------------------------------------------------

with pd.ExcelWriter(
    FILE_PATH,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace",
) as writer:
    summary_df.to_excel(
        writer,
        sheet_name=OUTPUT_SHEET,
        index=False,
    )

    ws = writer.book[OUTPUT_SHEET]

    # Insert STATUS row
    ws.insert_rows(1)
    ws["A1"] = "STATUS"
    ws["B1"] = status

    # Style STATUS cell
    ws["B1"].font = Font(bold=True)

    if status == "PASS":
        ws["B1"].fill = PatternFill(
            start_color="C6EFCE",
            end_color="C6EFCE",
            fill_type="solid",
        )
        ws["B1"].font = Font(bold=True, color="008000")
    else:
        ws["B1"].fill = PatternFill(
            start_color="FFC7CE",
            end_color="FFC7CE",
            fill_type="solid",
        )
        ws["B1"].font = Font(bold=True, color="FF0000")

    # Border around STATUS
    ws["B1"].border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Bold header row (Check / Result / Details)
    for cell in ws[2]:
        cell.font = Font(bold=True)

    # Auto-size column A
    max_length = max(len(str(cell.value)) for cell in ws["A"] if cell.value)
    ws.column_dimensions["A"].width = max_length + 5

# -----------------------------------------------------
# Preview output inside Colab (data only)
# -----------------------------------------------------

preview_df = pd.read_excel(FILE_PATH, sheet_name=OUTPUT_SHEET)
preview_df

# -----------------------------------------------------
# End of Module 1
# -----------------------------------------------------
