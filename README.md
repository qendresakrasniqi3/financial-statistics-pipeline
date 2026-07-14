# Financial Payments Statistics Pipeline — Germany (DE)

A end-to-end payments statistics production pipeline built in Python, R and SQL, modelled on the analytical framework used by the payments statistics production.

> **Disclaimer:** This pipeline uses synthetic data generated to approximate real payment statistics for Germany as published online. Volumes and values are calibrated to realistic orders of magnitude but numbers may slightly vary from official published figures. This dataset is intended for demonstration purposes only.

---

## Overview

This project simulates a quarterly payments statistics production workflow for Germany, covering data ingestion, validation, transformation, outlier detection, SQL analysis and reporting — 

| | |
|---|---|
| **Country** | Germany (DE) — largest euro area economy |
| **Period** | 2019-Q1 to 2024-Q4 |
| **Instruments** | Credit Transfer, Direct Debit, Card Payment, E-Money, Cheque |
| **Transaction Types** | Domestic, Cross-Border Intra-EU, Cross-Border Extra-EU |

---

## Repository Structure

```
financial-statistics-pipeline/
│
├── generate_data.py                    ← generates synthetic raw datasets
├── pipeline.py                         ← full end-to-end pipeline (7 steps)
├── time_series_chart.R                 ← quarterly time series visualisation
├── outlier_boxplot.R                   ← outlier detection boxplot
├── payments_statistics_datasets.xlsx   ← raw synthetic dataset
├── payments_quarterly_report.xlsx      ← pipeline output report
├── germany_payments_time_series.jpg    ← time series chart output
├── germany_outlier_boxplot.jpg         ← outlier detection boxplot output
└── README.md                           ← project documentation
```

---

## Pipeline Steps

The pipeline (`pipeline.py`) runs 7 sequential steps:

| Step | Name | Description |
|------|------|-------------|
| 1 | **Ingest** | Load datasets, schema validation, null detection and imputation, statistical summary |
| 2 | **Validate** | Business rules, referential integrity, data quality assessment |
| 3 | **Transform** | Data joins, derived variables, QoQ and YoY calculations |
| 4 | **Analyse** | Outlier detection (Z-score backend, IQR report), quarter-on-quarter spike detection |
| 5 | **Report** | Germany payments statistical summary tables saved to Excel |
| 6 | **SQL** | Business query via sqlite3 — instrument growth 2019 vs 2024 |
| 7 | **Test** | Unit tests, reconciliation checks, output consistency validation |

---

## Output Report

The pipeline produces `payments_quarterly_report.xlsx` with 6 sheets:

| Sheet | Description |
|-------|-------------|
| Payments Statistics | Quarterly report — volume, value, avg value, QoQ, YoY, outlier flag |
| Data Quality Assessment | Country-level DQA — missing values, completeness, imputed, dropped |
| Outlier Report | IQR-flagged observations across all transaction types with bounds |
| Country Indicators | Aggregated indicators — card share, cross-border share, dominant instrument |
| Reconciliation | Raw file vs pipeline totals — source of truth check |
| Instrument Growth 2019-2024 | SQL query result — which instrument grew most over 5 years |
| Metadata | Definitions, methodology notes, rounding explanation |

---

## Data Quality Approach

Missing values are handled using a two-stage strategy:

- **Numeric fields** (`total_value_eur_mn`, `number_of_transactions`) → imputed with median of same country + instrument group
- **Key fields** (`reporting_country`, `quarter`, `payment_instrument`) → rows dropped if null — identifiers cannot be imputed

This mirrors standard statistical production practice where missing values are estimated from historical series until the actual figure is received.

---

## Outlier Detection

Two methods are applied:

| Method | Logic | Used for |
|--------|-------|---------|
| Z-score | Flags \|z\| > 3 within instrument + transaction type series | Outlier Flag column in Payments Statistics sheet |
| IQR | Flags values beyond 1.5 × IQR from Q1/Q3 | Outlier Report sheet and boxplot chart |

**Note:** The Outlier Report covers all transaction types (domestic + cross-border). The boxplot chart filters to domestic transactions only for visual clarity. Both use the same IQR method — results may differ because they compare against different reference distributions.

Z-score is computed entirely in the backend — only the binary flag (0/1) is exposed in the output, consistent with statistical publication standards.

---

## SQL Queries

Queries run via Python's built-in `sqlite3` — no external database required.

**Inside pipeline.py (Step 6):**

| Query | Description |
|-------|-------------|
| Instrument Growth 2019 vs 2024 | Pure SQL using CTEs — which instrument grew most over 5 years? Results saved as a sheet in the Excel report |

---

## Visualisations (R)

| Script | Chart | Description |
|--------|-------|-------------|
| `time_series_chart.R` | Line chart | Total payment value by instrument, 2019-Q1 to 2024-Q4, domestic only |
| `outlier_boxplot.R` | Boxplot | Value distribution per instrument with IQR outliers highlighted in red, domestic only |

---

## Statistical Summary

Step 1 of the pipeline prints a statistical summary of transaction values:

```
Min, Max, Mean, Median, Std Dev — overall
Mean, Min, Max, Median — by payment instrument
```

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

**5. Generate charts (update paths in scripts first):**
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
Unit tests   : 14 passed / 0 failed
Reconciliation : 1/1 countries matched
```

---

*Built as a technical portfolio project demonstrating payments statistics production skills aligned with the payments statistics research analyst role.*
