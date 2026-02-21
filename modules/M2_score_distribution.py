# modules/M2_score_distribution.py
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as XLImage

from modules.helpers import write_df_to_sheet, save_fig_to_png

def run(df_std: pd.DataFrame, wb_out, out_dir: str = ".", tab_name: str = "M2_ScoreDistribution"):
    """
    Creates MatchScore distribution table + chart.
    Expects df_std to contain column: MatchScore_std
    Writes to Excel tab: M2_ScoreDistribution
    Returns: score_table, img_path
    """
    scores = [4, 5, 6, 7]
    freq = df_std["MatchScore_std"].value_counts().reindex(scores, fill_value=0).sort_index()
    total = int(freq.sum())

    score_table = pd.DataFrame({
        "MatchScore": scores,
        "Count": freq.values,
        "Percentage": np.round((freq.values / total) * 100, 2) if total else 0
    })

    if tab_name in wb_out.sheetnames:
        ws = wb_out[tab_name]
    else:
        ws = wb_out.create_sheet(tab_name)

    write_df_to_sheet(ws, score_table, index=False)

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.bar(score_table["MatchScore"].astype(str), score_table["Count"], color="#4F81BD")
    ax.set_xlabel("MatchScore (4–7)")
    ax.set_ylabel("Frequency (record pairs)")
    ax.set_title("MatchScore Frequency Distribution")
    for i, v in enumerate(score_table["Count"]):
        ax.text(i, v, str(int(v)), ha="center", va="bottom", fontsize=8)
    plt.tight_layout()

    img_path = os.path.join(out_dir, "m2_score_distribution.png")
    save_fig_to_png(fig, img_path)

    ws.add_image(XLImage(img_path), "E2")
    return score_table, img_path
