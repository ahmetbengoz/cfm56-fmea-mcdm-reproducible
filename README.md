# cfm56-fmea-mcdm-reproducible

Reproducible artifacts for a data-driven FMEA–MCDM case study on the CFM56-7B engine (Boeing 737).

## Overview
This repository provides a fully reproducible implementation of a data-driven FMEA–MCDM framework using real-world aviation data sources.

## Contents
- `analysis.py`: End-to-end Python script
- `requirements.txt`: Dependencies
- `outputs/`: Generated results (Excel + figures)
- `docs/`: Additional documentation

## Reproducibility
All results reported in the manuscript can be regenerated using:

```bash
pip install -r requirements.txt
python analysis.py

Outputs

The following artifacts are generated:
dataset_and_results.xlsx
Figure1_Workflow.png
Figure2_RankingComparison.png
Figure3_CRITIC_Weights.png
Figure4_Sensitivity_wS.png

Data Sources
NTSB accident reports (2016, 2019)
FAA Service Difficulty Reports (SDR)

License
MIT License
