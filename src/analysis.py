"""Reproduce the FMEA-MCDM analysis for the CFM aero-engine SDR screening study.

Run from the repository root:
    python src/analysis.py
"""
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
RESULTS = ROOT / "results"
RESULTS.mkdir(exist_ok=True)


def parse_period(period: str):
    period = str(period)
    if "-" in period:
        start, end = period.split("-")
        start, end = int(start), int(end)
    else:
        start = end = int(period)
    return start, end, end - start + 1


def minmax(matrix):
    matrix = np.asarray(matrix, dtype=float)
    denom = matrix.max(axis=0) - matrix.min(axis=0)
    denom = np.where(denom == 0, 1, denom)
    return (matrix - matrix.min(axis=0)) / denom


def critic_weights(X):
    std = X.std(axis=0, ddof=1)
    corr = np.corrcoef(X, rowvar=False)
    information = std * np.sum(1 - corr, axis=1)
    return information / information.sum()


def entropy_weights(X):
    P = X / np.where(X.sum(axis=0) == 0, 1, X.sum(axis=0))
    m = X.shape[0]
    entropy = -(P * np.log(P + 1e-12)).sum(axis=0) / np.log(m)
    diversification = 1 - entropy
    return diversification / diversification.sum()


def topsis(X, weights):
    V = X * weights
    ideal = V.max(axis=0)
    anti_ideal = V.min(axis=0)
    dplus = np.linalg.norm(V - ideal, axis=1)
    dminus = np.linalg.norm(V - anti_ideal, axis=1)
    closeness = dminus / (dplus + dminus)
    return dplus, dminus, closeness


def vikor(X, weights, v=0.5):
    fstar = X.max(axis=0)
    fminus = X.min(axis=0)
    denom = np.where(fstar - fminus == 0, 1, fstar - fminus)
    gaps = (fstar - X) / denom
    S = (weights * gaps).sum(axis=1)
    R = (weights * gaps).max(axis=1)
    Q = v * (S - S.min()) / (S.max() - S.min()) + (1 - v) * (R - R.min()) / (R.max() - R.min())
    return S, R, Q


df = pd.read_csv(DATA / "decision_matrix.csv")
if "Observation years" not in df.columns:
    parsed = df["Observation period"].apply(parse_period)
    df["Start year"] = [p[0] for p in parsed]
    df["End year"] = [p[1] for p in parsed]
    df["Observation years"] = [p[2] for p in parsed]
if "Annualized count" not in df.columns:
    df["Annualized count"] = df["Occurrence count"] / df["Observation years"]
if "O_period_adjusted" not in df.columns:
    df["O_period_adjusted"] = df["Annualized count"] / df["Annualized count"].max()
if "RPN_period_adjusted" not in df.columns:
    df["RPN_period_adjusted"] = df["O_period_adjusted"] * df["S"] * df["D"]

# Write normalized data back for transparency.
df.to_csv(DATA / "decision_matrix.csv", index=False)
df[[
    "Failure category", "Category type", "SDR keyword", "Occurrence count", "Observation period",
    "Start year", "End year", "Observation years", "Annualized count", "O", "O_period_adjusted",
    "S", "D", "RPN", "RPN_period_adjusted"
]].to_csv(DATA / "decision_matrix_period_adjusted.csv", index=False)

# Main analysis using raw-count occurrence indicator.
X = minmax(df[["O", "S", "D"]].values)
w_critic = critic_weights(X)
w_entropy = entropy_weights(X)
w_equal = np.ones(3) / 3
weights = pd.DataFrame({
    "Criterion": ["Occurrence", "Severity", "Detection"],
    "CRITIC": w_critic,
    "Entropy": w_entropy,
    "Equal": w_equal,
})
weights.to_csv(RESULTS / "weight_comparison.csv", index=False)

res = df.copy()
res["RPN_rank"] = res["RPN"].rank(ascending=False, method="min").astype(int)
for name, weights_vector in [("CRITIC", w_critic), ("Entropy", w_entropy), ("Equal", w_equal)]:
    dplus, dminus, score = topsis(X, weights_vector)
    S, R, Q = vikor(X, weights_vector, v=0.5)
    res[f"TOPSIS_{name}_Dplus"] = dplus
    res[f"TOPSIS_{name}_Dminus"] = dminus
    res[f"TOPSIS_{name}_score"] = score
    res[f"TOPSIS_{name}_rank"] = pd.Series(score).rank(ascending=False, method="min").astype(int).values
    res[f"VIKOR_{name}_S"] = S
    res[f"VIKOR_{name}_R"] = R
    res[f"VIKOR_{name}_Q"] = Q
    res[f"VIKOR_{name}_rank"] = pd.Series(Q).rank(ascending=True, method="min").astype(int).values
res.to_csv(RESULTS / "ranking_results.csv", index=False)

ranking_cols = [
    "RPN_rank", "TOPSIS_CRITIC_rank", "VIKOR_CRITIC_rank",
    "TOPSIS_Entropy_rank", "VIKOR_Entropy_rank",
    "TOPSIS_Equal_rank", "VIKOR_Equal_rank",
]
res[ranking_cols].corr(method="spearman").to_csv(RESULTS / "spearman_rank_correlation.csv")
res[ranking_cols].corr(method="kendall").to_csv(RESULTS / "kendall_rank_correlation.csv")

rank_stability = pd.DataFrame({"Rank": [1, 2, 3, 4, 5]})
for col in ["TOPSIS_CRITIC_rank", "TOPSIS_Entropy_rank", "TOPSIS_Equal_rank", "VIKOR_CRITIC_rank", "VIKOR_Entropy_rank", "VIKOR_Equal_rank"]:
    rank_stability[col.replace("_rank", "")] = res.sort_values(col).head(5)["Failure category"].tolist()
rank_stability.to_csv(RESULTS / "top5_rank_stability.csv", index=False)

# Sensitivity analysis using observation-period-adjusted occurrence.
Xp = minmax(df[["O_period_adjusted", "S", "D"]].values)
w_period = critic_weights(Xp)
dplus, dminus, score = topsis(Xp, w_period)
S, R, Q = vikor(Xp, w_period, v=0.5)
period = df[[
    "Failure category", "Occurrence count", "Observation period", "Observation years", "Annualized count",
    "O", "O_period_adjusted", "S", "D", "RPN", "RPN_period_adjusted",
]].copy()
period["CRITIC_period_adjusted_occurrence_weight"] = w_period[0]
period["CRITIC_period_adjusted_severity_weight"] = w_period[1]
period["CRITIC_period_adjusted_detection_weight"] = w_period[2]
period["TOPSIS_period_adjusted_score"] = score
period["TOPSIS_period_adjusted_rank"] = pd.Series(score).rank(ascending=False, method="min").astype(int).values
period["VIKOR_period_adjusted_Q"] = Q
period["VIKOR_period_adjusted_rank"] = pd.Series(Q).rank(ascending=True, method="min").astype(int).values
period["RPN_period_adjusted_rank"] = period["RPN_period_adjusted"].rank(ascending=False, method="min").astype(int)
period.sort_values("TOPSIS_period_adjusted_rank").to_csv(RESULTS / "period_adjusted_sensitivity.csv", index=False)

# Perturbation analysis of S and D scores.
rng = np.random.default_rng(42)
N = 10000
rank_records = np.empty((N, len(df)), dtype=int)
for t in range(N):
    perturbed_s = np.clip(df["S"].values + rng.integers(-1, 2, len(df)), 1, 10)
    perturbed_d = np.clip(df["D"].values + rng.integers(-1, 2, len(df)), 1, 10)
    criteria = np.column_stack([df["O"].values, perturbed_s, perturbed_d]).astype(float)
    Xp = minmax(criteria)
    wp = critic_weights(Xp)
    _, _, score = topsis(Xp, wp)
    rank_records[t] = pd.Series(score).rank(ascending=False, method="min").astype(int).values

perturbation = pd.DataFrame({"Failure category": df["Failure category"]})
perturbation["Mean rank"] = rank_records.mean(axis=0)
perturbation["Std. rank"] = rank_records.std(axis=0)
perturbation["Best rank"] = rank_records.min(axis=0)
perturbation["Worst rank"] = rank_records.max(axis=0)
perturbation["Top-1 frequency"] = (rank_records == 1).mean(axis=0)
perturbation["Top-3 frequency"] = (rank_records <= 3).mean(axis=0)
perturbation["Top-5 frequency"] = (rank_records <= 5).mean(axis=0)
perturbation.sort_values("Mean rank").to_csv(RESULTS / "perturbation_summary.csv", index=False)

print("Analysis complete. Outputs written to data/ and results/.")
