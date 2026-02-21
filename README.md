# Duplicate Patient Validation Dashboard (DT402B)

This repository contains the modular Python logic for an explainable “Duplicate Patient Review” dashboard used to support migration governance decisions (threshold safety, mismatch explainability, triage/workload estimation, and auditability). 

## What this produces (Excel tab outputs)

**Input (supplier raw tab):**
- `DuplicatePatientsSummary` 

**Output tabs created/updated by this dashboard:**
- `M1_SchemaValidation`
- `M2_ScoreDistribution`
- `M2b_EntityClustering`  *(from `Input_DuplicatePatientSummary` if present)*
- `M3_MismatchPatterns` *(aggregate only)*
- `M4_TriageCategorisation`
- `M5_Audit` 

## Repo structure (keep it simple)

- `/modules/` = one file per module (M1–M5 + M2b), aligned to the Excel output tabs. 
  - `M1_schema_validation.py`
  - `M2_score_distribution.py`
  - `M2b_entity_clustering.py`
  - `M3_mismatch_patterns.py`
  - `M4_triage_categorisation.py`
  - `M5_audit.py` 

## Dependencies

Create a `requirements.txt` in the repo root with:
- pandas
- numpy
- openpyxl
- pillow
- matplotlib

## Google Colab workflow (cell blocks)

The notebook version follows these stages: 
- C1 Install dependencies + safe plotting backend
- C2 Set file paths + preview controls
- C3 Imports + helper functions (schema, parsing, safe previews)
- C4 Detect sheet names + map required columns
- C5 M1 Schema Validation + preview (halts later modules on failure)
- C6 Create output workbook + write M1 + M5 (Audit base)
- C7 M2 Score Distribution + chart + preview
- C8 M2b Entity Clustering + long tail + Pareto chart + previews
- C9 M3 Mismatch Patterns (aggregate-only) + stacked chart + preview
- C10 M4 Triage Categorisation + donut chart + previews
- C11 Finalise Audit + save workbook
- C12 Download the processed workbook 
## Important (data handling)

Do **NOT** commit supplier spreadsheets or processed output workbooks to GitHub.
Keep only code + documentation in this repository. 
