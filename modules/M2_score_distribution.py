"""
Module 2 Profile Match‑Score Distribution

Purpose:
- Profile distribution of duplicate records by match score (0–7)
- Identify patient IDs participating in multiple duplicate relationships
- Support discussion on safe merge thresholds and human review selection

Acceptance Criteria covered:
- All match scores 0–7 included (zero‑filled)
- Pivot‑ready outputs
- No mutation of source data
- Per‑ID duplicate count and average match score
- Ranked to surface highest‑risk IDs
- Conditional formatting for merge decisions
- Visual distribution of merge outcomes
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
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import PieChart, Reference


# =====================================================
# 2. File configuration
# =====================================================
SOURCE_FILE = "DUPLICATE_PATIENTS.xlsx"
SOURCE_SHEET = "Input_DuplicatePatientSummary"
TARGET_SHEET = "M2_MatchScoreDistribution"


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
# 4. Match score distribution (0–7, zero‑filled)
# =====================================================
score_index = pd.DataFrame({"match_score": range(0, 8)})

score_dist = (
    df.groupby("matches")
      .size()
      .reset_index(name="count")
      .rename(columns={"matches": "match_score"})
)

score_dist = (
    score_index
    .merge(score_dist, on="match_score", how="left")
    .fillna(0)
)

score_dist["count"] = score_dist["count"].astype(int)


# =====================================================
# 5. ID participation (duplicate relationships per patient)
# =====================================================
id_pairs = pd.concat(
    [
        df[["id_1", "matches"]].rename(columns={"id_1": "patient_id"}),
        df[["id_2", "matches"]].rename(columns={"id_2": "patient_id"})
    ],
    ignore_index=True
)

id_risk = (
    id_pairs
    .groupby("patient_id")
    .agg(
        duplicate_pair_count=("matches", "count"),
        average_matchscore=("matches", "mean")
    )
    .reset_index()
    .sort_values("duplicate_pair_count", ascending=False)
)


# =====================================================
# 6. Clinical merge decision logic
# =====================================================
def merge_decision(avg_score):
    if avg_score >= 6.5:
        return "Safe for auto merge"
    elif avg_score >= 5:
        return "Human review required"
    else:
        return "Do not merge"


id_risk["merge_decision"] = id_risk["average_matchscore"].apply(merge_decision)


# =====================================================
# 7. Write outputs to Excel
# =====================================================
wb = load_workbook(SOURCE_FILE)
ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb.create_sheet(TARGET_SHEET)
ws.delete_rows(1, ws.max_row)

bold = Font(bold=True)

# --- Left table: Match score distribution ---
ws["A1"] = "Match Score Distribution"
ws["A1"].font = bold
ws["A2"] = "Match Score"
ws["B2"] = "Count"
ws["A2"].font = ws["B2"].font = bold

for i, r in enumerate(score_dist.itertuples(index=False), start=3):
    ws[f"A{i}"] = r.match_score
    ws[f"B{i}"] = r.count

# --- Right table: ID risk + merge decision ---
headers = ["Patient ID", "Duplicate Pair Count", "Average Match Score", "Merge Decision"]
for col, h in zip(["D", "E", "F", "G"], headers):
    ws[f"{col}2"] = h
    ws[f"{col}2"].font = bold

for i, r in enumerate(id_risk.itertuples(index=False), start=3):
    ws[f"D{i}"] = r.patient_id
    ws[f"E{i}"] = r.duplicate_pair_count
    ws[f"F{i}"] = round(r.average_matchscore, 2)
    ws[f"G{i}"] = r.merge_decision


# =====================================================
# 8. Conditional formatting for merge decisions
# =====================================================
last_row = ws.max_row

ws.conditional_formatting.add(
    f"G3:G{last_row}",
    FormulaRule(formula=['G3="Safe for auto merge"'], font=Font(color="006100"))
)

ws.conditional_formatting.add(
    f"G3:G{last_row}",
    FormulaRule(formula=['G3="Human review required"'], font=Font(color="9C6500"))
)

ws.conditional_formatting.add(
    f"G3:G{last_row}",
    FormulaRule(formula=['G3="Do not merge"'], font=Font(color="9C0006"))
)


# =====================================================
# 9. Summary table for chart
# =====================================================
summary = id_risk["merge_decision"].value_counts().reset_index()
summary.columns = ["Decision", "Count"]

ws["J2"] = "Decision"
ws["K2"] = "Count"
ws["J2"].font = ws["K2"].font = bold

for i, r in enumerate(summary.itertuples(index=False), start=3):
    ws[f"J{i}"] = r.Decision
    ws[f"K{i}"] = r.Count


# =====================================================
# 10. Pie chart (merge decision distribution)
# =====================================================
pie = PieChart()
pie.title = "Record Merge Decision Distribution"

labels = Reference(ws, min_col=10, min_row=3, max_row=2 + len(summary))
data = Reference(ws, min_col=11, min_row=2, max_row=2 + len(summary))

pie.add_data(data, titles_from_data=True)
pie.set_categories(labels)
pie.height = 9
pie.width = 9

ws.add_chart(pie, "J6")


# =====================================================
# 11. Save workbook
# =====================================================
wb.save(SOURCE_FILE)


# =====================================================
# 12. Google Colab preview
# =====================================================
if IN_COLAB:
    display(score_dist)
    display(id_risk.head(20))
    files.download(SOURCE_FILE)


# -----------------------------------------------------
# End of Module 2
# -----------------------------------------------------
