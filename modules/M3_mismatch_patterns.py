"""
Module 3 – Mismatch Patterns for Scores 6 and 5

Purpose:
- Explain why near‑match duplicate records fail to reach auto‑merge confidence
- Identify which demographic fields differ at MatchScores 5 and 6

MVP link:
- In scope #4: Explain near‑miss behaviour

Acceptance Criteria:
- Focus only on MatchScores 5 and 6
- Summarise mismatched demographic fields (e.g. Postcode, DOB, Name)
- Output must be interpretable by non‑technical reviewers
- Derived outputs only (no mutation of source data)
"""

# =====================================================
# 0. Enable Google Colab preview (safe locally)
# =====================================================
try:
    from google.colab import files
    from IPython.display import display
    IN_COLAB = True
except ImportError:
    IN_COLAB = False


# =====================================================
# 1. Imports
# =====================================================
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference


# =====================================================
# 2. File configuration
# =====================================================
SOURCE_FILE = "DUPLICATE_PATIENTS.xlsx"
SOURCE_SHEET = "Input_DuplicatePatientSummary"
TARGET_SHEET = "M3_MismatchPatterns"


# =====================================================
# 3. Load source data (read‑only)
# =====================================================
df = pd.read_excel(
    SOURCE_FILE,
    sheet_name=SOURCE_SHEET,
    engine="openpyxl"
)

df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]


# =====================================================
# 4. Restrict to near‑matches (MatchScore 5 and 6 only)
# =====================================================
near = df[df["matches"].isin([5, 6])].copy()


# =====================================================
# 5. Identify demographic match indicators
# Assumption: boolean fields ending with '_match'
# =====================================================
match_cols = [c for c in near.columns if c.endswith("_match")]


# =====================================================
# 6. Summarise mismatches by score (counts + percentages)
# =====================================================
rows = []

for score in [6, 5]:
    subset = near[near["matches"] == score]
    total = len(subset)

    for col in match_cols:
        mismatch_count = (~subset[col].astype(bool)).sum()

        if mismatch_count > 0:
            rows.append({
                "Score": score,
                "Field Name": col.replace("_match", "").replace("_", " ").title(),
                "Mismatch Count": int(mismatch_count),
                "Mismatch %": round((mismatch_count / total) * 100, 1)
            })

result_df = pd.DataFrame(rows)


# =====================================================
# 7. Pivot to side‑by‑side format (Score 6 vs Score 5)
# =====================================================
pivot = result_df.pivot_table(
    index="Field Name",
    columns="Score",
    values=["Mismatch Count", "Mismatch %"],
    fill_value=0
)

pivot.columns = [f"Score {c[1]} – {c[0]}" for c in pivot.columns]
pivot = pivot.reset_index()


# =====================================================
# 8. Write outputs to Excel
# =====================================================
wb = load_workbook(SOURCE_FILE)
ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb.create_sheet(TARGET_SHEET)
ws.delete_rows(1, ws.max_row)

bold = Font(bold=True)

# Title
ws["A1"] = "Mismatch Drivers for Near‑Match Duplicate Records (Scores 5 vs 6)"
ws["A1"].font = bold

# Table headers
for col_idx, header in enumerate(pivot.columns, start=1):
    ws.cell(row=2, column=col_idx, value=header).font = bold

# Table rows
for r_idx, row in enumerate(pivot.itertuples(index=False), start=3):
    for c_idx, value in enumerate(row, start=1):
        ws.cell(row=r_idx, column=c_idx, value=value)


# =====================================================
# 9. Visualisation – mismatch percentage by field
# =====================================================
chart = BarChart()
chart.title = "Mismatch Percentage by Field (Score 6 vs Score 5)"
chart.y_axis.title = "Mismatch %"
chart.x_axis.title = "Demographic Field"

max_row = ws.max_row

# Add only percentage columns to chart
percentage_columns = [
    i + 1 for i, h in enumerate(pivot.columns)
    if "Mismatch %" in h
]

for col in percentage_columns:
    data = Reference(ws, min_col=col, min_row=2, max_row=max_row)
    chart.add_data(data, titles_from_data=True)

categories = Reference(ws, min_col=1, min_row=3, max_row=max_row)
chart.set_categories(categories)

chart.height = 10
chart.width = 18

ws.add_chart(chart, f"A{max_row + 3}")


# =====================================================
# 10. Save workbook
# =====================================================
wb.save(SOURCE_FILE)


# =====================================================
# 11. Google Colab preview
# =====================================================
if IN_COLAB:
    display(pivot)
    files.download(SOURCE_FILE)


# -----------------------------------------------------
# End of Module 3
# -----------------------------------------------------
