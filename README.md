# cfm56-fmea-mcdm-reproducible

Reproducible artifacts for a data-driven **FMEA–MCDM** case study on the **CFM56-7B** engine (Boeing 737).

## What this repository provides
- `analysis.py`: end-to-end Python script that regenerates:
  - `outputs/dataset_and_results.xlsx`
  - publication-ready figures (`outputs/figures/`, PNG, 300 dpi)
- `outputs/`: generated artifacts (Excel + figures) uploaded for transparency and peer verification.

## How to reproduce (local)
```bash
pip install -r requirements.txt
python analysis.py

Data sources (high-level)

NTSB accident/incident investigation reports are used to ground severity and detection logic.

FAA SDR keyword-frequency proxies are used to construct an occurrence index and normalize it by the maximum category count.

Repository structure

analysis.py

requirements.txt

outputs/

docs/
