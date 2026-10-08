# TATA Forage: Delinquency Analysis

Python analysis and project deliverables for exploring customer delinquency patterns in the TATA Forage project.

## Certification

Verified by Forage (User Verification Code: 6a460628515132a0f9c5930e).

## Project contents

- `main.py`: validates data and explores delinquency rates, numeric associations, and payment history.
- `view_distinct_values.py`: displays distinct values and missing-value counts for each dataset column.
- `task2.py`: placeholder for planning a delinquency prediction model.
- `delinquency_results/`: exported aggregate analysis reports.
- `Final_Output/`: written reports, model plan, and presentation materials.

## Setup

Use Python 3 with the analysis dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
pip install pandas numpy openpyxl
```

Place `Delinquency_prediction_dataset.xlsx` in the project directory. The dataset and `Forage_files/` are excluded from Git and must be supplied locally.

## Run the analysis

```bash
python main.py Delinquency_prediction_dataset.xlsx --output delinquency_results
python view_distinct_values.py Delinquency_prediction_dataset.xlsx --max 20
```

The main analysis also accepts CSV files. Set `--month-order oldest-first` or `--month-order newest-first` only after confirming the payment history chronology. Use `--min-group` to adjust the minimum group size flag (default: 20).

## Analysis outputs

The analysis writes data quality and validation reports, delinquency rates by group with 95% confidence intervals, numeric comparisons, and Spearman associations. Missing or invalid values remain missing rather than being imputed.

Results describe observed associations. This repository does not yet implement a trained prediction model; verify that payment history precedes the outcome before using it for modeling.
