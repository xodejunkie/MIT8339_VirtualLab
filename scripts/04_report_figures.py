#!/usr/bin/env python3
"""
MIT 8339 Virtual Lab - Report figures (matplotlib)
Renders analytical figures used in the Word/PDF report.
"""
import json
import sys
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import SEGMENTS, FIGS, KPIS_JSON

FIGS.mkdir(parents=True, exist_ok=True)
plt.rcParams["axes.unicode_minus"] = False

BLUE, RED, AMBER, GREEN = "#1B2A4A", "#C0392B", "#F2994A", "#2E8B57"

seg = lambda name: pd.read_csv(SEGMENTS / f"{name}.csv")


def style_ax(ax):
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color("#BBBBBB")
    ax.tick_params(colors="#444444", labelsize=10)
    ax.yaxis.grid(True, color="#E5E5E5", lw=0.8)
    ax.set_axisbelow(True)


def bar_fig(tbl, cat, val, title, fname, highlight=None, figsize=(8.2, 4.4), fmt="{:.0%}"):
    fig, ax = plt.subplots(figsize=figsize, constrained_layout=True)
    colors = [RED if (highlight and str(c) == highlight) else BLUE for c in tbl[cat]]
    bars = ax.bar(tbl[cat].astype(str), tbl[val], color=colors, width=0.62)
    for b, v in zip(bars, tbl[val]):
        ax.annotate(
            fmt.format(v), (b.get_x() + b.get_width() / 2, v),
            ha="center", va="bottom", fontsize=10, color="#333333",
            xytext=(0, 2), textcoords="offset points",
        )
    ax.set_title(title, fontsize=13, fontweight="bold", color="#1A1A1A", pad=10)
    ax.set_ylabel("Attrition rate", fontsize=11, color="#444444")
    ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
    ax.set_ylim(0, tbl[val].max() * 1.22)
    style_ax(ax)
    fig.savefig(FIGS / fname, dpi=170, facecolor="white")
    plt.close(fig)
    print("[FIG]", fname)


tf = seg("by_trans_frequency")
bar_fig(
    tf, "Transaction_Frequency", "Attrition_Rate",
    "Attrition rate by transaction frequency - the engagement cliff",
    "fig1_transaction_cliff.png", highlight="Low",
)

ut = seg("by_utilisation")
bar_fig(
    ut, "Utilisation_Band", "Attrition_Rate",
    "Attrition rate by utilisation band - dormant cards are the risk pocket",
    "fig2_utilisation_bands.png", highlight="0% (Non-revolver)",
)

pr = seg("by_products")
bar_fig(
    pr, "Total_Relationship_Count", "Attrition_Rate",
    "Attrition rate by number of products held (relationship depth)",
    "fig3_products_held.png",
)

rb = seg("by_risk_band")
fig, ax = plt.subplots(figsize=(8.2, 4.6), constrained_layout=True)
bars = ax.bar(rb["Churn_Risk_Band"], rb["Customers"], color=[GREEN, AMBER, RED], width=0.5)
for b, v, r_ in zip(bars, rb["Customers"], rb["Attrition_Rate"]):
    ax.annotate(
        f"{v:,} customers\nattrition {r_:.1%}",
        (b.get_x() + b.get_width() / 2, v),
        ha="center", va="bottom", fontsize=10.5, color="#1A1A1A",
        xytext=(0, 3), textcoords="offset points",
    )
ax.set_title(
    "Rule-based retention-risk bands - population size vs realised attrition",
    fontsize=13, fontweight="bold", color="#1A1A1A", pad=10,
)
ax.set_ylabel("Customers", fontsize=11, color="#444444")
ax.set_ylim(0, rb["Customers"].max() * 1.30)
style_ax(ax)
fig.savefig(FIGS / "fig4_risk_bands.png", dpi=170, facecolor="white")
plt.close(fig)
print("[FIG] fig4_risk_bands.png")

with open(KPIS_JSON) as f:
    kj = json.load(f)
odds = kj["model"]["odds_ratios"]
feats = sorted(odds, key=lambda k: abs(1 - odds[k]))
vals = [odds[k] for k in feats]
fig, ax = plt.subplots(figsize=(8.6, 4.8), constrained_layout=True)
colors = [RED if v > 1 else BLUE for v in vals]
ax.barh(feats, vals, color=colors, height=0.55)
ax.axvline(1.0, color="#999999", lw=1.2, ls="--")
for i, v in enumerate(vals):
    ax.annotate(
        f"{v:.2f}", (v, i), ha="left", va="center", fontsize=10,
        color="#333333", xytext=(4, 0), textcoords="offset points",
    )
ax.set_title(
    f"Logistic regression - odds ratios per 1 SD (holdout AUC {kj['model']['auc']})",
    fontsize=12.5, fontweight="bold", color="#1A1A1A", pad=10,
)
ax.set_xlabel("Odds ratio (standardised features)", fontsize=11, color="#444444")
ax.spines[["top", "right"]].set_visible(False)
ax.spines[["left", "bottom"]].set_color("#BBBBBB")
ax.xaxis.grid(True, color="#E5E5E5", lw=0.8)
ax.set_axisbelow(True)
ax.tick_params(colors="#444444", labelsize=10)
fig.savefig(FIGS / "fig5_odds_ratios.png", dpi=170, facecolor="white")
plt.close(fig)
print("[FIG] fig5_odds_ratios.png")

mia = seg("by_months_inactive")
fig, ax = plt.subplots(figsize=(8.2, 4.2), constrained_layout=True)
ax.plot(mia["Months_Inactive_12_mon"], mia["Attrition_Rate"], marker="o", color=BLUE, lw=2.2)
ax.fill_between(mia["Months_Inactive_12_mon"], mia["Attrition_Rate"], color=BLUE, alpha=0.08)
ax.set_title(
    "Attrition rate by months inactive in the last 12 months",
    fontsize=13, fontweight="bold", color="#1A1A1A", pad=10,
)
ax.set_xlabel("Months inactive (12-month window)", fontsize=11, color="#444444")
ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
style_ax(ax)
fig.savefig(FIGS / "fig6_months_inactive.png", dpi=170, facecolor="white")
plt.close(fig)
print("[FIG] fig6_months_inactive.png")
print("DONE - figures in", FIGS)
