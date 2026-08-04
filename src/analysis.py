"""Reproduce the CFM aero-engine maintenance-review screening analysis.

The primary occurrence proxy is the keyword-query count divided by the
inclusive searchable observation window. Raw query counts are retained as a
secondary sensitivity specification. Neither quantity is an exposure-based
failure rate or a count of unique events.

Run from the repository root:
    python src/analysis.py
"""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
RESULTS = ROOT / "results"
RESULTS.mkdir(exist_ok=True)

SEED = 42
N_PERTURBATIONS = 10_000


def parse_period(period: str) -> tuple[int, int, int]:
    """Return start year, end year and inclusive number of calendar years."""
    value = str(period).strip()
    if "-" in value:
        start_text, end_text = value.split("-", maxsplit=1)
        start, end = int(start_text), int(end_text)
    else:
        start = end = int(value)
    if end < start:
        raise ValueError(f"Invalid observation period: {period}")
    return start, end, end - start + 1


def minmax(matrix: np.ndarray) -> np.ndarray:
    """Column-wise min-max normalization with constant-column protection."""
    values = np.asarray(matrix, dtype=float)
    denominator = values.max(axis=0) - values.min(axis=0)
    denominator = np.where(denominator == 0, 1.0, denominator)
    return (values - values.min(axis=0)) / denominator


def critic_weights(matrix: np.ndarray) -> np.ndarray:
    """CRITIC weights from contrast intensity and inter-criterion conflict."""
    values = np.asarray(matrix, dtype=float)
    standard_deviation = values.std(axis=0, ddof=1)
    correlation = np.corrcoef(values, rowvar=False)
    information = standard_deviation * np.sum(1 - correlation, axis=1)
    if np.isclose(information.sum(), 0):
        return np.ones(values.shape[1]) / values.shape[1]
    return information / information.sum()


def entropy_weights(matrix: np.ndarray) -> np.ndarray:
    """Entropy weights using zero-safe proportions."""
    values = np.asarray(matrix, dtype=float)
    column_totals = values.sum(axis=0)
    proportions = np.divide(
        values,
        column_totals,
        out=np.zeros_like(values),
        where=column_totals != 0,
    )
    log_terms = np.zeros_like(proportions)
    positive = proportions > 0
    log_terms[positive] = proportions[positive] * np.log(proportions[positive])
    entropy = -log_terms.sum(axis=0) / np.log(values.shape[0])
    diversification = 1 - entropy
    if np.isclose(diversification.sum(), 0):
        return np.ones(values.shape[1]) / values.shape[1]
    return diversification / diversification.sum()


def topsis(matrix: np.ndarray, weights: np.ndarray) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """TOPSIS distance and closeness values for risk-increasing criteria."""
    weighted = np.asarray(matrix, dtype=float) * np.asarray(weights, dtype=float)
    ideal = weighted.max(axis=0)
    anti_ideal = weighted.min(axis=0)
    distance_positive = np.linalg.norm(weighted - ideal, axis=1)
    distance_negative = np.linalg.norm(weighted - anti_ideal, axis=1)
    denominator = distance_positive + distance_negative
    closeness = np.divide(
        distance_negative,
        denominator,
        out=np.zeros_like(distance_negative),
        where=denominator != 0,
    )
    return distance_positive, distance_negative, closeness


def vikor(matrix: np.ndarray, weights: np.ndarray, v: float = 0.5) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """VIKOR S, R and Q values for risk-increasing criteria."""
    values = np.asarray(matrix, dtype=float)
    best = values.max(axis=0)
    worst = values.min(axis=0)
    criterion_range = np.where(best - worst == 0, 1.0, best - worst)
    weighted_gaps = np.asarray(weights, dtype=float) * ((best - values) / criterion_range)
    group_utility = weighted_gaps.sum(axis=1)
    individual_regret = weighted_gaps.max(axis=1)

    s_range = group_utility.max() - group_utility.min()
    r_range = individual_regret.max() - individual_regret.min()
    s_scaled = np.zeros_like(group_utility) if np.isclose(s_range, 0) else (
        (group_utility - group_utility.min()) / s_range
    )
    r_scaled = np.zeros_like(individual_regret) if np.isclose(r_range, 0) else (
        (individual_regret - individual_regret.min()) / r_range
    )
    index = v * s_scaled + (1 - v) * r_scaled
    return group_utility, individual_regret, index


def rank_descending(values: np.ndarray | pd.Series) -> np.ndarray:
    return pd.Series(values).rank(ascending=False, method="min").astype(int).to_numpy()


def rank_ascending(values: np.ndarray | pd.Series) -> np.ndarray:
    return pd.Series(values).rank(ascending=True, method="min").astype(int).to_numpy()


def load_decision_data() -> pd.DataFrame:
    """Build the complete decision matrix from immutable query and scoring inputs."""
    query = pd.read_csv(DATA / "query_counts.csv")
    scoring = pd.read_csv(DATA / "scoring.csv")
    decision = query.merge(scoring, on="Failure category", how="inner", validate="one_to_one")

    if len(decision) != len(query) or len(decision) != len(scoring):
        raise ValueError("Query and scoring inputs do not contain the same failure categories")
    if decision["Occurrence count"].lt(0).any():
        raise ValueError("Occurrence counts must be non-negative")
    if not decision["S"].between(1, 10).all() or not decision["D"].between(1, 10).all():
        raise ValueError("Severity and detection scores must be within 1-10")

    periods = decision["Observation period"].map(parse_period)
    decision["Start year"] = [item[0] for item in periods]
    decision["End year"] = [item[1] for item in periods]
    decision["Observation years"] = [item[2] for item in periods]
    decision["Window-adjusted count"] = (
        decision["Occurrence count"] / decision["Observation years"]
    )
    decision["O_raw"] = decision["Occurrence count"] / decision["Occurrence count"].max()
    decision["O_window_adjusted"] = (
        decision["Window-adjusted count"] / decision["Window-adjusted count"].max()
    )
    decision["Baseline_window"] = decision["O_window_adjusted"] * decision["S"] * decision["D"]
    decision["Baseline_raw"] = decision["O_raw"] * decision["S"] * decision["D"]

    column_order = [
        "Failure category", "Category type", "SDR keyword", "Occurrence count",
        "Observation period", "Start year", "End year", "Observation years",
        "Window-adjusted count", "O_window_adjusted", "O_raw", "S", "D",
        "Baseline_window", "Baseline_raw", "Scoring rationale",
    ]
    return decision[column_order]


def evaluate_specification(decision: pd.DataFrame, occurrence_column: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Compute baseline, TOPSIS and VIKOR results for one occurrence proxy."""
    normalized = minmax(decision[[occurrence_column, "S", "D"]].to_numpy())
    weight_sets = {
        "CRITIC": critic_weights(normalized),
        "Entropy": entropy_weights(normalized),
        "Equal": np.ones(3) / 3,
    }
    weights = pd.DataFrame(
        {
            "Criterion": ["Occurrence proxy", "Severity", "Detection"],
            **weight_sets,
        }
    )

    result = decision.copy()
    baseline_column = "Baseline_window" if occurrence_column == "O_window_adjusted" else "Baseline_raw"
    result["Baseline_rank"] = rank_descending(result[baseline_column])
    for name, weight_vector in weight_sets.items():
        dplus, dminus, closeness = topsis(normalized, weight_vector)
        s_value, r_value, q_value = vikor(normalized, weight_vector, v=0.5)
        result[f"TOPSIS_{name}_Dplus"] = dplus
        result[f"TOPSIS_{name}_Dminus"] = dminus
        result[f"TOPSIS_{name}_score"] = closeness
        result[f"TOPSIS_{name}_rank"] = rank_descending(closeness)
        result[f"VIKOR_{name}_S"] = s_value
        result[f"VIKOR_{name}_R"] = r_value
        result[f"VIKOR_{name}_Q"] = q_value
        result[f"VIKOR_{name}_rank"] = rank_ascending(q_value)

    if not np.allclose(weights[["CRITIC", "Entropy", "Equal"]].sum(axis=0), 1.0):
        raise AssertionError("Criterion weights do not sum to one")
    return weights, result


def perturb_primary(decision: pd.DataFrame) -> pd.DataFrame:
    """Perturb author-assigned S/D scores under the primary occurrence proxy."""
    rng = np.random.default_rng(SEED)
    rank_records = np.empty((N_PERTURBATIONS, len(decision)), dtype=int)
    occurrence = decision["O_window_adjusted"].to_numpy()

    for simulation in range(N_PERTURBATIONS):
        severity = np.clip(decision["S"].to_numpy() + rng.integers(-1, 2, len(decision)), 1, 10)
        detection = np.clip(decision["D"].to_numpy() + rng.integers(-1, 2, len(decision)), 1, 10)
        normalized = minmax(np.column_stack([occurrence, severity, detection]))
        weights = critic_weights(normalized)
        _, _, score = topsis(normalized, weights)
        rank_records[simulation] = rank_descending(score)

    perturbation = pd.DataFrame({"Failure category": decision["Failure category"]})
    perturbation["Mean rank"] = rank_records.mean(axis=0)
    perturbation["Std. rank"] = rank_records.std(axis=0)
    perturbation["Best rank"] = rank_records.min(axis=0)
    perturbation["Worst rank"] = rank_records.max(axis=0)
    perturbation["Top-1 frequency"] = (rank_records == 1).mean(axis=0)
    perturbation["Top-3 frequency"] = (rank_records <= 3).mean(axis=0)
    perturbation["Top-5 frequency"] = (rank_records <= 5).mean(axis=0)
    return perturbation.sort_values(["Mean rank", "Failure category"]).reset_index(drop=True)


def main() -> None:
    decision = load_decision_data()
    decision.to_csv(DATA / "decision_matrix.csv", index=False)

    primary_weights, primary = evaluate_specification(decision, "O_window_adjusted")
    raw_weights, raw = evaluate_specification(decision, "O_raw")
    primary_weights.to_csv(RESULTS / "weight_comparison.csv", index=False)
    raw_weights.to_csv(RESULTS / "raw_count_weight_comparison.csv", index=False)
    primary.to_csv(RESULTS / "ranking_results.csv", index=False)

    ranking_columns = [
        "Baseline_rank", "TOPSIS_CRITIC_rank", "VIKOR_CRITIC_rank",
        "TOPSIS_Entropy_rank", "VIKOR_Entropy_rank",
        "TOPSIS_Equal_rank", "VIKOR_Equal_rank",
    ]
    primary[ranking_columns].corr(method="spearman").to_csv(
        RESULTS / "spearman_rank_correlation.csv"
    )
    primary[ranking_columns].corr(method="kendall").to_csv(
        RESULTS / "kendall_rank_correlation.csv"
    )

    top_five = pd.DataFrame({"Rank": [1, 2, 3, 4, 5]})
    for column in ranking_columns[1:]:
        top_five[column.replace("_rank", "")] = (
            primary.sort_values([column, "Failure category"]).head(5)["Failure category"].tolist()
        )
    top_five.to_csv(RESULTS / "top5_rank_stability.csv", index=False)

    benchmark = primary[
        [
            "Failure category", "Category type", "Baseline_window", "Baseline_rank",
            "TOPSIS_CRITIC_score", "TOPSIS_CRITIC_rank", "VIKOR_CRITIC_Q", "VIKOR_CRITIC_rank",
        ]
    ].copy()
    benchmark["TOPSIS_minus_baseline_rank"] = (
        benchmark["TOPSIS_CRITIC_rank"] - benchmark["Baseline_rank"]
    )
    benchmark["VIKOR_minus_baseline_rank"] = (
        benchmark["VIKOR_CRITIC_rank"] - benchmark["Baseline_rank"]
    )
    benchmark.sort_values(["TOPSIS_CRITIC_rank", "Failure category"]).to_csv(
        RESULTS / "benchmark_comparison.csv", index=False
    )

    sensitivity = primary[
        ["Failure category", "O_window_adjusted", "TOPSIS_CRITIC_score", "TOPSIS_CRITIC_rank", "VIKOR_CRITIC_Q", "VIKOR_CRITIC_rank"]
    ].rename(
        columns={
            "TOPSIS_CRITIC_score": "Window_TOPSIS_score",
            "TOPSIS_CRITIC_rank": "Window_TOPSIS_rank",
            "VIKOR_CRITIC_Q": "Window_VIKOR_Q",
            "VIKOR_CRITIC_rank": "Window_VIKOR_rank",
        }
    )
    raw_slice = raw[
        ["Failure category", "O_raw", "TOPSIS_CRITIC_score", "TOPSIS_CRITIC_rank", "VIKOR_CRITIC_Q", "VIKOR_CRITIC_rank"]
    ].rename(
        columns={
            "TOPSIS_CRITIC_score": "Raw_TOPSIS_score",
            "TOPSIS_CRITIC_rank": "Raw_TOPSIS_rank",
            "VIKOR_CRITIC_Q": "Raw_VIKOR_Q",
            "VIKOR_CRITIC_rank": "Raw_VIKOR_rank",
        }
    )
    sensitivity = sensitivity.merge(raw_slice, on="Failure category", validate="one_to_one")
    sensitivity["TOPSIS_rank_shift_raw_minus_window"] = (
        sensitivity["Raw_TOPSIS_rank"] - sensitivity["Window_TOPSIS_rank"]
    )
    sensitivity["VIKOR_rank_shift_raw_minus_window"] = (
        sensitivity["Raw_VIKOR_rank"] - sensitivity["Window_VIKOR_rank"]
    )
    sensitivity.sort_values(["Window_TOPSIS_rank", "Failure category"]).to_csv(
        RESULTS / "raw_count_sensitivity.csv", index=False
    )

    perturbation = perturb_primary(decision)
    perturbation.to_csv(RESULTS / "perturbation_summary.csv", index=False)

    manifest = {
        "primary_occurrence_proxy": "keyword-query count divided by inclusive searchable observation years",
        "secondary_occurrence_proxy": "raw keyword-query count",
        "interpretation": "screening indicators; not unique-event counts, exposure-normalized rates or reliability estimates",
        "perturbation_seed": SEED,
        "perturbation_runs": N_PERTURBATIONS,
        "vikor_v": 0.5,
    }
    (RESULTS / "analysis_manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n", encoding="utf-8"
    )

    print("Analysis complete. Primary and sensitivity outputs written to results/.")


if __name__ == "__main__":
    main()
