"""Regenerate all manuscript figures from analysis outputs.

Run after ``python src/analysis.py`` from the repository root.
Both PNG previews and 300-dpi TIFF submission files are written to figures/.
"""

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Rectangle
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
FIGURES = ROOT / "figures"
RESULTS = ROOT / "results"
FIGURES.mkdir(exist_ok=True)

plt.rcParams.update({
    "font.size": 10,
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
})


def save_figure(fig: plt.Figure, stem: str) -> None:
    fig.savefig(FIGURES / f"{stem}.png", dpi=300, bbox_inches="tight")
    fig.savefig(FIGURES / f"{stem}.tif", dpi=300, bbox_inches="tight", pil_kwargs={"compression": "tiff_lzw"})
    plt.close(fig)


def figure_pipeline() -> None:
    fig, ax = plt.subplots(figsize=(10, 5.5))
    ax.axis("off")
    boxes = [
        (0.05, 0.68, "Public evidence\nSDR query outputs + investigations"),
        (0.38, 0.68, "Frozen analytical inputs\nquery counts + scoring rationale"),
        (0.71, 0.68, "FMEA-informed matrix\nwindow-adjusted O, S and D"),
        (0.71, 0.28, "Weighting and ranking\nCRITIC / entropy / equal\nTOPSIS + VIKOR"),
        (0.38, 0.28, "Robustness and benchmark\nO×S×D + raw sensitivity\n+ perturbation"),
        (0.05, 0.28, "Maintenance-review queue\nrecord retrieval\n+ engineering follow-up"),
    ]
    for x, y, label in boxes:
        ax.add_patch(Rectangle((x, y), 0.24, 0.12, linewidth=1.2, edgecolor="black", facecolor="white"))
        ax.text(x + 0.12, y + 0.06, label, ha="center", va="center", fontsize=8.5)

    def arrow(x1: float, y1: float, x2: float, y2: float) -> None:
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="->", mutation_scale=12, linewidth=1.0, color="black"))

    arrow(0.29, 0.74, 0.38, 0.74)
    arrow(0.62, 0.74, 0.71, 0.74)
    arrow(0.83, 0.68, 0.83, 0.40)
    arrow(0.71, 0.34, 0.62, 0.34)
    arrow(0.38, 0.34, 0.29, 0.34)
    fig.tight_layout()
    save_figure(fig, "figure1_screening_pipeline")


def figure_weights() -> None:
    weights = pd.read_csv(RESULTS / "weight_comparison.csv")
    fig, ax = plt.subplots(figsize=(8, 4.8))
    x = np.arange(len(weights))
    width = 0.25
    for offset, method, color in [(-width, "CRITIC", "#0072B2"), (0, "Entropy", "#E69F00"), (width, "Equal", "#009E73")]:
        bars = ax.bar(x + offset, weights[method], width, label=method, color=color)
        ax.bar_label(bars, fmt="%.3f", padding=3, fontsize=8)
    ax.set_xticks(x)
    ax.set_xticklabels(["Occurrence proxy", "Severity", "Detection"])
    ax.set_ylim(0, max(0.5, weights[["CRITIC", "Entropy", "Equal"]].to_numpy().max() + 0.08))
    ax.set_ylabel("Weight")
    ax.legend(frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.12))
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    save_figure(fig, "figure2_primary_criterion_weights")


def figure_primary_ranking() -> None:
    ranking = pd.read_csv(RESULTS / "ranking_results.csv").sort_values("TOPSIS_CRITIC_rank").head(10)
    plot = ranking.iloc[::-1]
    fig, ax = plt.subplots(figsize=(9, 5.2))
    bars = ax.barh(plot["Failure category"], plot["TOPSIS_CRITIC_score"], color="#0072B2")
    labels = [f"{score:.3f}  (rank {rank})" for score, rank in zip(plot["TOPSIS_CRITIC_score"], plot["TOPSIS_CRITIC_rank"])]
    ax.bar_label(bars, labels=labels, padding=4, fontsize=8)
    ax.set_xlabel("Window-adjusted CRITIC-TOPSIS score")
    ax.set_xlim(0, min(1.0, plot["TOPSIS_CRITIC_score"].max() + 0.18))
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    save_figure(fig, "figure3_primary_window_adjusted_ranking")


def figure_perturbation() -> None:
    perturbation = pd.read_csv(RESULTS / "perturbation_summary.csv").sort_values("Mean rank").head(10)
    plot = perturbation.iloc[::-1]
    fig, ax = plt.subplots(figsize=(9, 5.2))
    bars = ax.barh(
        plot["Failure category"],
        plot["Mean rank"],
        xerr=plot["Std. rank"],
        color="#0072B2",
        error_kw={"elinewidth": 1.0, "capsize": 3},
    )
    ax.bar_label(bars, labels=[f"{value:.2f}" for value in plot["Mean rank"]], padding=4, fontsize=8)
    ax.set_xlabel("Mean rank under 10,000 S/D perturbations (lower = higher priority)")
    ax.set_xlim(0, max(12.5, float((plot["Mean rank"] + plot["Std. rank"]).max()) + 1.0))
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    save_figure(fig, "figure4_primary_score_perturbation")


def main() -> None:
    figure_pipeline()
    figure_weights()
    figure_primary_ranking()
    figure_perturbation()
    print("Figures regenerated in figures/ as PNG and 300-dpi TIFF files.")


if __name__ == "__main__":
    main()
