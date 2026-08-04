# CFM Aero-Engine Maintenance-Review Signal Prioritization

Reproducibility package for the manuscript:

**Prioritizing CFM Aero-Engine Maintenance-Review Signals from Public Service Difficulty Reports: An FMEA-Informed Robustness Study**

## Study purpose

This package converts public Service Difficulty Report (SDR) keyword-query indicators into a transparent maintenance-review screening matrix. It compares a multiplicative FMEA-style baseline with CRITIC, entropy and equal weighting and with TOPSIS and VIKOR ranking. It also reports rank agreement, raw-count sensitivity and 10,000 score perturbations.

The workflow is **FMEA-informed screening**, not a conventional component-level FMEA and not a reliability-estimation model. The labels intentionally include event, component/system and condition/symptom signals because public SDR narratives are operational records rather than a uniform failure-mode taxonomy.

## Interpretation boundary

- Query counts are keyword matches, not deduplicated events.
- The primary occurrence proxy divides each query count by its inclusive searchable observation window.
- Raw query counts are retained only as a secondary sensitivity specification.
- Neither occurrence specification is normalized by engine-hours, flight cycles, fleet size or model-specific exposure.
- Severity and detection scores are author-assigned ordinal screening values supported by a stated rubric and evidence rationale.
- Outputs indicate where record-level engineering review should begin; they do not prescribe inspection intervals or maintenance actions.

## Repository structure

```text
data/
  query_counts.csv             immutable query-count input
  scoring.csv                  severity/detection inputs and rationale
  decision_matrix.csv          regenerated analytical matrix
results/
  analysis_manifest.json
  benchmark_comparison.csv
  ranking_results.csv
  raw_count_sensitivity.csv
  raw_count_weight_comparison.csv
  weight_comparison.csv
  top5_rank_stability.csv
  spearman_rank_correlation.csv
  kendall_rank_correlation.csv
  perturbation_summary.csv
figures/
  figure1_screening_pipeline.(png|tif)
  figure2_primary_criterion_weights.(png|tif)
  figure3_primary_window_adjusted_ranking.(png|tif)
  figure4_primary_score_perturbation.(png|tif)
src/
  analysis.py
  make_figures.py
tests/
  test_analysis.py
docs/
  EFA_FMEA_MCDM_supplementary_material.docx
requirements.txt
README.md
```

## Reproduction

Use Python 3.10 or later from the repository root:

```bash
python -m pip install -r requirements.txt
python src/analysis.py
python -m unittest discover -s tests -v
python src/make_figures.py
```

`analysis.py` rebuilds the decision matrix from `query_counts.csv` and `scoring.csv`; it does not use previously generated results as inputs. The random seed and number of perturbations are recorded in `results/analysis_manifest.json`.

## Public evidence sources

- AviationDB/FAA Service Difficulty Report query interface: https://aviationdb.com/Aviation/SdrQuery.shtm
- NTSB investigation DCA16FA217: https://www.ntsb.gov/investigations/pages/DCA16FA217.aspx
- NTSB accident report AAR-19/03: https://www.ntsb.gov/investigations/AccidentReports/Reports/AAR1903.pdf

## Archive

The previously registered archive identifier is https://doi.org/10.5281/zenodo.20492171. Users should rely on the GitHub repository for the current window-adjusted primary specification until the archive is refreshed to the same release.

## License

Code and package documentation are provided under the repository's MIT License. Public-source data remain subject to their originating providers' terms and citation requirements.
