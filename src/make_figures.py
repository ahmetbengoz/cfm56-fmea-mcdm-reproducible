"""Regenerate manuscript figures.

Run after src/analysis.py:
    python src/make_figures.py
"""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyArrowPatch

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "figures"
RES = ROOT / "results"
FIG.mkdir(exist_ok=True)

plt.rcParams.update({"font.size": 10, "font.family": "DejaVu Sans"})

# Figure 1: workflow.
fig, ax = plt.subplots(figsize=(10, 5.5))
ax.axis("off")
boxes = [
    (0.05, 0.72, "Public evidence inputs\nAviationDB SDR + NTSB reports"),
    (0.37, 0.72, "Data construction\nkeyword protocol + taxonomy"),
    (0.69, 0.72, "FMEA matrix\nO, S and D"),
    (0.05, 0.42, "Occurrence indicators\nraw and period-adjusted counts"),
    (0.37, 0.42, "Evidence-linked scoring\nseverity + detection rubric"),
    (0.69, 0.42, "Weighting module\nCRITIC + entropy + equal"),
    (0.20, 0.12, "Ranking module\nTOPSIS + VIKOR"),
    (0.58, 0.12, "Robustness assessment\ncorrelation + perturbation"),
]
for x, y, text in boxes:
    rect = Rectangle((x, y), 0.24, 0.12, linewidth=1.2, edgecolor="black", facecolor="white")
    ax.add_patch(rect)
    ax.text(x + 0.12, y + 0.06, text, ha="center", va="center", fontsize=9)

def arrow(x1, y1, x2, y2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="->", mutation_scale=12, linewidth=1.0, color="black"))
arrow(0.29,0.78,0.37,0.78); arrow(0.61,0.78,0.69,0.78)
arrow(0.81,0.72,0.81,0.54); arrow(0.69,0.48,0.61,0.48); arrow(0.37,0.48,0.29,0.48)
arrow(0.49,0.42,0.32,0.24); arrow(0.81,0.42,0.70,0.24); arrow(0.44,0.12,0.58,0.18)
plt.tight_layout()
plt.savefig(FIG / "figure1_pipeline.png", dpi=300, bbox_inches="tight")
plt.close()

# Figure 2: weights.
weights = pd.read_csv(RES / "weight_comparison.csv")
fig, ax = plt.subplots(figsize=(8, 4.8))
x = range(len(weights))
width = 0.25
ax.bar([i - width for i in x], weights["CRITIC"], width, label="CRITIC")
ax.bar(x, weights["Entropy"], width, label="Entropy")
ax.bar([i + width for i in x], weights["Equal"], width, label="Equal")
ax.set_xticks(list(x))
ax.set_xticklabels(weights["Criterion"])
ax.set_ylim(0, 0.50)
ax.set_ylabel("Weight")
ax.set_title("Criterion weights under alternative weighting specifications")
ax.legend(frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.10))
ax.grid(axis="y", alpha=0.25)
plt.tight_layout()
plt.savefig(FIG / "figure2_weights.png", dpi=300, bbox_inches="tight")
plt.close()

# Figure 3: period-adjusted occurrence sensitivity.
period = pd.read_csv(RES / "period_adjusted_sensitivity.csv").sort_values("TOPSIS_period_adjusted_rank").head(10)
fig, ax = plt.subplots(figsize=(9, 5.2))
plot = period.iloc[::-1]
ax.barh(plot["Failure category"], plot["O_period_adjusted"])
ax.set_xlabel("Occurrence indicator normalized by annualized SDR count")
ax.set_title("Observation-period-adjusted occurrence sensitivity")
ax.grid(axis="x", alpha=0.25)
plt.tight_layout()
plt.savefig(FIG / "figure3_period_adjusted_sensitivity.png", dpi=300, bbox_inches="tight")
plt.close()

# Figure 4: perturbation summary.
pert = pd.read_csv(RES / "perturbation_summary.csv").sort_values("Mean rank").head(10)
fig, ax = plt.subplots(figsize=(9, 5.2))
plot = pert.iloc[::-1]
ax.barh(plot["Failure category"], plot["Mean rank"])
ax.set_xlabel("Mean rank under 10,000 S/D perturbations; lower is higher priority")
ax.set_title("Perturbation robustness of highest-priority failure categories")
ax.grid(axis="x", alpha=0.25)
plt.tight_layout()
plt.savefig(FIG / "figure4_perturbation.png", dpi=300, bbox_inches="tight")
plt.close()

print("Figures regenerated in figures/.")
