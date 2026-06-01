# Evidence-Informed Failure Prioritization of CFM Aero-Engine SDRs Using FMEA-MCDM

This repository contains the reproducibility package for the manuscript:

**Evidence-Informed Failure Prioritization of CFM Aero-Engine Service Difficulty Reports Using FMEA-MCDM and Robustness Analysis**

## Purpose

The package reproduces the screening-level FMEA-MCDM analysis used to prioritize CFM aero-engine Service Difficulty Report (SDR) keyword categories. It includes query-count inputs, the decision matrix, severity/detection scoring rationale, weighting results, ranking outputs, period-adjusted occurrence sensitivity, rank-correlation outputs, perturbation outputs and generated figures.

## Important interpretation note

The occurrence values are **keyword-query occurrence indicators** from the AviationDB/FAA SDR query interface. They are not unique event counts, not model-specific reliability estimates and not failure rates. The period-adjusted sensitivity analysis annualizes keyword counts only to test the effect of unequal observation windows; it does not replace exposure normalization using engine-hours, flight cycles or fleet denominators.

## Repository structure

```text
data/
  query_counts.csv
  decision_matrix.csv
  decision_matrix_period_adjusted.csv
results/
  weight_comparison.csv
  ranking_results.csv
  top5_rank_stability.csv
  spearman_rank_correlation.csv
  kendall_rank_correlation.csv
  period_adjusted_sensitivity.csv
  perturbation_summary.csv
figures/
  figure1_pipeline.png
  figure2_weights.png
  figure3_period_adjusted_sensitivity.png
  figure4_perturbation.png
src/
  analysis.py
  make_figures.py
docs/
  supplementary_material.docx
requirements.txt
README.md
```

## Reproduction

From the repository root:

```bash
pip install -r requirements.txt
python src/analysis.py
python src/make_figures.py
```

The analysis script writes updated CSV outputs to `results/` and the figure script regenerates the manuscript figures in `figures/`.

## Data sources

The public source inputs are:

- AviationDB/FAA Service Difficulty Report query interface.
- National Transportation Safety Board investigation/report material for CFM56-7B fan-blade separation cases.

## Version status

This is the archived reproducibility package associated with the manuscript. The Zenodo DOI is: [DOI].
