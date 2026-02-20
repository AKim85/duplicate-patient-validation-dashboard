"""
Module 2 Profile Match‑Score Distribution
Purpose:
- Profile distribution of duplicate records by match score (0–7)
- Identify patient IDs participating in multiple duplicate relationships
- Support discussion on safe merge thresholds

Acceptance Criteria covered:
- All match scores 0–7 included (zero-filled)
- Pivot-ready outputs
- No mutation of source data
- Per‑ID duplicate count and average match score
- Ranked to surface highest-risk IDs
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

SOURCE_FILE = "DUPLICATE_PATIENTS.xlsx"
TARGET_SHEET = "M2_MatchScoreDistribution"

# ------------------------------------------------------------------
# Load source data (read-only)
# ------------------------------------------------------------------
df = pd.read_excel(SOURCE_FILE, sheet_name=0, engine="openpyxl")
df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

# ------------------------------------------------------------------
# Score Distribution (0–7, zero-filled)
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
# Write both outputs into a single tab (side-by-side)
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

# Right table: ID Risk Concentration
ws["D1"] = "ID Risk Concentration"
ws["D1"].font = bold
ws["D2"] = "Patient Id"
ws["E2"] = "Duplicate Pair Count"
ws["F2"] = "Average Matchscore"

for cell in ["D2", "E2", "F2"]:
    ws[cell].font = bold

for i, r in enumerate(id_risk.itertuples(index=False), start=3):
    ws[f"D{i}"] = r.patient_id
    ws[f"E{i}"] = r.duplicate_pair_count
    ws[f"F{i}"] = round(r.average_matchscore, 2)

wb.save(SOURCE_FILE)
