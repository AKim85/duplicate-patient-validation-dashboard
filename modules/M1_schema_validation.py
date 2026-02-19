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

# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------

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

# ---------------------------------------------------------------------
# Load input
# ---------------------------------------------------------------------

xls = pd.ExcelFile(FILE_PATH)

if INPUT_SHEET not in xls.sheet_names:
    raise ValueError(
        f"Expected sheet '{INPUT_SHEET}' not found. "
        f"Available sheets: {xls.sheet_names}"
    )

df = pd.read_excel(FILE_PATH, sheet_name=INPUT_SHEET)

# ---------------------------------------------------------------------
# Validation checks
# ---------------------------------------------------------------------

# 1. Schema check
missing_columns = [c for c in EXPECTED_COLUMNS if c not in df.columns]

# 2. Date parsing check
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

# ---------------------------------------------------------------------
# Determine PASS / FAIL
# ---------------------------------------------------------------------

status = "PASS" if len(missing_columns) == 0 else "FAIL"

# ---------------------------------------------------------------------
# Build validation summary
# ---------------------------------------------------------------------

summary_rows = []

summary_rows.append({
    "Check": "Missing columns",
    "Result": len(missing_columns),
    "Details": ", ".join(missing_columns),
})

summary_rows.append({
    "Check": "Invalid dates (Recent Edit)",
    "Result": invalid_dates,
    "Details": "",
})

for col, count in non_numeric_counts.items():
    summary_rows.append({
        "Check": f"Non-numeric values in {col}",
        "Result": count,
        "Details": "",
    })

summary_df = pd.DataFrame(summary_rows)

# ---------------------------------------------------------------------
# Write output (without modifying source data)
# ---------------------------------------------------------------------

with pd.ExcelWriter(
    FILE_PATH,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    summary_df.to_excel(
        writer,
        sheet_name=OUTPUT_SHEET,
        index=False
    )

    worksheet = writer.book[OUTPUT_SHEET]
    worksheet.insert_rows(1)
    worksheet["A1"] = "STATUS"
    worksheet["B1"] = status

# ---------------------------------------------------------------------
# End of Module 1
# ---------------------------------------------------------------------
