# Financial Payments Statistics Pipeline

A end-to-end payments statistics production pipeline built in Python, R and SQL, modelled on the analytical framework used by the ECB Payments Statistics team.

> **Disclaimer:** This pipeline uses synthetic data calibrated to approximate real payment statistics for Germany as published by the ECB. Volumes and values are realistic in order of magnitude but may slightly vary from official published figures. This dataset is intended for demonstration purposes only.

---

## Overview

This project simulates a quarterly payments statistics production workflow covering data ingestion, validation, transformation, outlier detection, SQL analysis and reporting — mirroring the core tasks of a Research Analyst in the ECB Payments Statistics team.

**Country:** Germany (DE)  
**Period:** 2019-Q1 to 2024-Q4  
**Instruments:** Credit Transfer, Direct Debit, Card Payment, E-Money, Cheque  

---

## Repository Structure

```
financial-statistics-pipeline/
│
├── generate_data.py                    ← generates synthetic raw datasets
├── pipeline.py                         ← full end-to-end pipeline (7 steps)
├── statistical_summary.py              ← overall statistical summary
├── queries.sql                         ← YoY and QoQ SQL queries
├── running_total.sql                   ← running total by country
├── time_series_chart.R                 ← quarterly time series visualisation
├── outlier_boxplot.R                   ← outlier detection boxplot
├── payments_statistics_datasets.xlsx   ← raw synthetic dataset
└── payments_quarterly_report.xlsx      ← pipeline output report
```

---

## Pipeline Steps

The pipeline (`pipeline.py`) runs 7 sequential steps:

| Step | Name | Description |
|------|------|-------------|
| 1 | **Ingest** | Load datasets, schema validation, null handling and imputation, statistical summary |
| 2 | **Validate** | Business rules, referential integrity, data quality assessment |
| 3 | **Transform** | Data joins, derived variables, QoQ and YoY calculations |
| 4 | **Analyse** | Outlier detection (Z-score, IQR), quarter-on-quarter spike detection, country trends |
| 5 | **Report** | ECB-style quarterly summary tables saved to Excel |
| 6 | **SQL** | Business queries via sqlite3 — instrument growth 2019 vs 2024 |
| 7 | **Test** | Unit tests, reconciliation checks, output consistency validation |

---

## Output Report

The pipeline produces `payments_quarterly_report.xlsx` with 6 sheets:

| Sheet | Description |
|-------|-------------|
| Payments Statistics | Quarterly report — volume, value, avg value, QoQ, YoY, outlier flag |
| Data Quality Assessment | Country-level DQA — missing values, completeness, validity |
| Outlier Report | IQR-flagged observations with bounds and instrument detail |
| Country Indicators | Aggregated indicators — card share, cross-border share, dominant instrument |
| Reconciliation | Raw file vs pipeline totals — 19/19 countries matched |
| Instrument Growth 2019-2024 | SQL query result — which instrument grew most over 5 years |
| Metadata | Definitions, methodology notes, rounding explanation |

---

## Data Quality Approach

Missing values are handled using a two-stage strategy:

- **Numeric fields** (`total_value_eur_mn`, `number_of_transactions`) → imputed with group median per instrument and transaction type
- **Key fields** (`reporting_country`, `quarter`, `payment_instrument`) → rows dropped if null as identifiers cannot be imputed

This mirrors real ECB statistical production practice where missing national central bank submissions are estimated from historical series until the actual figure is received.

---

## Outlier Detection

Two methods are applied:

| Method | Logic | Used for |
|--------|-------|---------|
| Z-score | Flags \|z\| > 3 within instrument series | Outlier Flag column in main report |
| IQR | Flags values beyond 1.5 × IQR from Q1/Q3 | Outlier Report sheet and boxplot chart |

Z-score is computed entirely in the backend — only the binary flag (0/1) is exposed in the output, consistent with statistical publication standards.

---

## SQL Queries

Three SQL queries run via Python's built-in `sqlite3` — no external database required:

- **YoY Growth** — absolute and percentage change vs same quarter one year prior
- **QoQ Growth** — absolute and percentage change vs immediately preceding quarter
- **Running Total** — cumulative transaction value over time
- **Instrument Ranking** — top 3 instruments by volume per quarter (2024)
- **Instrument Growth** — 2019 vs 2024 comparison using CTEs

---

## Visualisations (R)

| Script | Chart | Description |
|--------|-------|-------------|
| `time_series_chart.R` | Line chart | Total payment value by instrument, 2019-Q1 to 2024-Q4 |
| `outlier_boxplot.R` | Boxplot | Value distribution per instrument with IQR outliers highlighted in red |

---

## Getting Started

**1. Install Python dependencies:**
```bash
pip install pandas openpyxl numpy
```

**2. Install R dependencies (run once in R):**
```r
install.packages(c("readxl", "ggplot2", "dplyr", "scales", "ggrepel"))
```

**3. Generate synthetic dataset:**
```bash
python3 generate_data.py
```

**4. Run full pipeline:**
```bash
python3 pipeline.py
```

**5. Run statistical summary:**
```bash
python3 statistical_summary.py
```

**6. Generate charts (update paths in scripts first):**
```bash
Rscript time_series_chart.R
Rscript outlier_boxplot.R
```

---

## Technical Stack

| Language | Usage |
|----------|-------|
| Python | Data pipeline, validation, outlier detection, reporting |
| SQL | Analytical queries via sqlite3 |
| R | Statistical visualisations (ggplot2) |

---

## Validation Results

```
Business rules : 17 passed / 0 failed
Unit tests     : 14 passed / 0 failed
Reconciliation : 1/1 countries matched
```

---

*Built as a technical portfolio project demonstrating payments statistics production skills aligned with the ECB Research Analyst role.*
