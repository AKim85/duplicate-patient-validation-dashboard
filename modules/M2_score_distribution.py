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


import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font


SOURCE_FILE = "DUPLICATE_PATIENTS.xlsx"
SOURCE_SHEET = "Input_DuplicatePatientSummary"
TARGET_SHEET = "M2_MatchScoreDistribution"


# ------------------------------------------------------------------
# Load source data (read‑only)
# ------------------------------------------------------------------
df = pd.read_excel(
    SOURCE_FILE,
    sheet_name=SOURCE_SHEET,
    engine="openpyxl"
)

df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]


# ------------------------------------------------------------------
# Score Distribution (0–7, zero‑filled)
# ------------------------------------------------------------------
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


# ------------------------------------------------------------------
# ID Participation (duplicate relationships per patient)
# ------------------------------------------------------------------
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


# ------------------------------------------------------------------
# Clinical merge decision logic (simple, explicit)
# ------------------------------------------------------------------
def merge_decision(avg_score):
    if avg_score >= 6.5:
        return "Safe for auto merge"
    elif avg_score >= 5:
        return "Human review required"
    else:
        return "Do not merge"


id_risk["merge_decision"] = id_risk["average_matchscore"].apply(merge_decision)


# ------------------------------------------------------------------
# Write outputs into a single tab (side‑by‑side)
# ------------------------------------------------------------------
wb = load_workbook(SOURCE_FILE)
ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb.create_sheet(TARGET_SHEET)
ws.delete_rows(1, ws.max_row)

bold = Font(bold=True)

# Left table: Match Score Distribution
ws["A1"] = "Match Score Distribution"
ws["A1"].font = bold
ws["A2"] = "Match Score"
ws["B2"] = "Count"
ws["A2"].font = bold
ws["B2"].font = bold

for i, r in enumerate(score_dist.itertuples(index=False), start=3):
    ws[f"A{i}"] = r.match_score
    ws[f"B{i}"] = r.count


# Right table: ID Risk Concentration + Clinical Decision
ws["D1"] = "ID Risk Concentration"
ws["D1"].font = bold
ws["D2"] = "Patient ID"
ws["E2"] = "Duplicate Pair Count"
ws["F2"] = "Average Match Score"
ws["G2"] = "Merge Decision"

for cell in ["D2", "E2", "F2", "G2"]:
    ws[cell].font = bold

for i, r in enumerate(id_risk.itertuples(index=False), start=3):
    ws[f"D{i}"] = r.patient_id
    ws[f"E{i}"] = r.duplicate_pair_count
    ws[f"F{i}"] = round(r.average_matchscore, 2)
    ws[f"G{i}"] = r.merge_decision


wb.save(SOURCE_FILE)


# ------------------------------------------------------------------
# Google Colab preview
# ------------------------------------------------------------------
if IN_COLAB:
    print("Match score distribution")
    display(score_dist)

    print("Highest‑risk patient IDs")
    display(id_risk.head(20))

    files.download(SOURCE_FILE)
