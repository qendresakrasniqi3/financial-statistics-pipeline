# =============================================================================
# generate_data.py
# Generates synthetic payments statistics datasets for Germany (DE).
# Data is calibrated to approximate real ECB payments statistics for Germany.
#
# DISCLAIMER:
#   This is synthetic data intended for demonstration purposes only.
#   Volumes and values are calibrated to realistic orders of magnitude
#   but numbers may slightly vary from official published figures.
#
# Run this first before pipeline.py:
#   python3 generate_data.py
#
# Output: payments_statistics_datasets.xlsx with 4 sheets:
#   - payments_transactions_quarterly
#   - psp_reference_data
#   - card_payments_detail
#   - validation_rules
# =============================================================================

import pandas as pd
import numpy as np
from datetime import date, timedelta
import random

np.random.seed(42)
random.seed(42)

OUTPUT_PATH = "/mnt/c/Users/qendresa.krasniqi/Downloads/payments_statistics_datasets.xlsx"

# --- CONFIG ---
# Germany only — largest euro area economy and ECB payments statistics reporter
COUNTRY             = "DE"
PAYMENT_INSTRUMENTS = ["credit_transfer","direct_debit","card_payment","e-money","cheque"]
QUARTERS            = [f"{y}Q{q}" for y in range(2019, 2025) for q in range(1, 5)]
CARD_TYPES          = ["debit","credit","prepaid"]
TRANSACTION_TYPES   = ["domestic","cross_border_intra_eu","cross_border_extra_eu"]


# =============================================================================
# 1. PAYMENT TRANSACTIONS QUARTERLY
# Germany-calibrated volumes and values based on ECB published statistics:
#   - Credit transfers: ~600 mn per quarter, avg ~3,500 EUR (retail + corporate)
#   - Direct debits:    ~3 bn per quarter,  avg ~200 EUR  (utilities, subscriptions)
#   - Card payments:    ~4 bn per quarter,  avg ~55 EUR   (retail POS + online)
#   - E-money:          ~100 mn per quarter, avg ~25 EUR  (digital wallets)
#   - Cheques:          ~1 mn per quarter,  avg ~1,500 EUR (declining instrument)
# =============================================================================

rows = []
for quarter in QUARTERS:
    for instrument in PAYMENT_INSTRUMENTS:
        for txn_type in TRANSACTION_TYPES:

            # Base volumes calibrated to Germany quarterly figures
            base_volume = {
                "credit_transfer": np.random.randint(550_000_000,  650_000_000),
                "direct_debit":    np.random.randint(2_800_000_000, 3_200_000_000),
                "card_payment":    np.random.randint(3_500_000_000, 4_500_000_000),
                "e-money":         np.random.randint(80_000_000,   120_000_000),
                "cheque":          np.random.randint(800_000,      1_200_000),
            }[instrument]

            # Scale cross-border volumes down (domestic dominates)
            if txn_type == "cross_border_intra_eu":
                base_volume = int(base_volume * np.random.uniform(0.05, 0.10))
            elif txn_type == "cross_border_extra_eu":
                base_volume = int(base_volume * np.random.uniform(0.01, 0.03))

            # Upward trend over time — card and e-money growing faster
            year = int(quarter[:4])
            trend = {
                "credit_transfer": 1 + (year - 2019) * 0.02,
                "direct_debit":    1 + (year - 2019) * 0.02,
                "card_payment":    1 + (year - 2019) * 0.06,
                "e-money":         1 + (year - 2019) * 0.15,
                "cheque":          1 - (year - 2019) * 0.08,  # declining
            }[instrument]

            volume = int(base_volume * trend * np.random.uniform(0.97, 1.03))
            volume = max(volume, 1)

            # Average transaction value in EUR
            avg_value = {
                "credit_transfer": np.random.uniform(3_000, 4_000),
                "direct_debit":    np.random.uniform(150,   250),
                "card_payment":    np.random.uniform(45,    65),
                "e-money":         np.random.uniform(15,    35),
                "cheque":          np.random.uniform(1_000, 2_000),
            }[instrument]

            total_value_eur_mn = round(volume * avg_value / 1_000_000, 2)

            # Inject ~1% outliers for outlier detection testing
            if np.random.random() < 0.01:
                volume             = int(volume * np.random.uniform(3, 8))
                total_value_eur_mn = round(total_value_eur_mn * np.random.uniform(3, 8), 2)

            rows.append({
                "reporting_country":      COUNTRY,
                "quarter":                quarter,
                "payment_instrument":     instrument,
                "transaction_type":       txn_type,
                "number_of_transactions": volume,
                "total_value_eur_mn":     total_value_eur_mn,
            })

df_transactions = pd.DataFrame(rows)

# --- Inject ~1% nulls to simulate real-world missing data ---
null_idx_value  = df_transactions.sample(frac=0.01, random_state=1).index
null_idx_volume = df_transactions.sample(frac=0.01, random_state=2).index
df_transactions.loc[null_idx_value,  "total_value_eur_mn"]     = np.nan
df_transactions.loc[null_idx_volume, "number_of_transactions"] = np.nan

print(f"1. transactions — {len(df_transactions):,} rows | "
      f"nulls injected: {df_transactions.isnull().sum().sum()}")


# =============================================================================
# 2. PSP REFERENCE DATA — Germany-based PSPs
# =============================================================================
psp_types = ["credit_institution","e-money_institution","payment_institution","post_office"]
n_psps    = 50  # Germany has ~50 major PSPs reporting to Bundesbank
psp_ids   = [f"PSP_DE_{str(i).zfill(3)}" for i in range(1, n_psps+1)]

df_psps = pd.DataFrame({
    "psp_id":          psp_ids,
    "psp_name":        [f"DE_Provider_{i}" for i in range(1, n_psps+1)],
    "psp_type":        np.random.choice(psp_types, n_psps, p=[0.60, 0.10, 0.25, 0.05]),
    "home_country":    COUNTRY,
    "license_date":    [
        (date(2000, 1, 1) + timedelta(days=int(np.random.randint(0, 8000)))).isoformat()
        for _ in range(n_psps)
    ],
    "is_active":        np.random.choice([True, False], n_psps, p=[0.94, 0.06]),
    "psd2_licensed":    np.random.choice([True, False], n_psps, p=[0.85, 0.15]),
    "reporting_agent":  np.random.choice([True, False], n_psps, p=[0.40, 0.60]),
})
print(f"2. psps         — {len(df_psps):,} rows")


# =============================================================================
# 3. CARD PAYMENTS DETAIL — Germany
# =============================================================================
half_years = [f"{y}H{h}" for y in range(2019, 2025) for h in [1, 2]]
card_rows  = []
for period in half_years:
    for card_type in CARD_TYPES:
        for txn_type in TRANSACTION_TYPES:
            volume = np.random.randint(500_000_000, 2_000_000_000)
            value  = round(volume * np.random.uniform(45, 65) / 1_000_000, 2)
            card_rows.append({
                "reporting_country":        COUNTRY,
                "period":                   period,
                "card_type":                card_type,
                "transaction_type":         txn_type,
                "at_terminal_type":         random.choice(["POS","ATM","online","contactless"]),
                "number_of_cards_issued_mn":round(np.random.uniform(10, 50), 2),
                "number_of_transactions":   volume,
                "total_value_eur_mn":       value,
            })

df_cards = pd.DataFrame(card_rows)
print(f"3. cards        — {len(df_cards):,} rows")


# =============================================================================
# 4. VALIDATION RULES
# =============================================================================
validation_rules = [
    {"rule_id":"VR001","dataset":"transactions","field":"reporting_country",
     "rule_type":"reference_check","description":"Must be DE","severity":"critical"},
    {"rule_id":"VR002","dataset":"transactions","field":"payment_instrument",
     "rule_type":"allowed_values","description":"Must be from allowed instrument list","severity":"critical"},
    {"rule_id":"VR003","dataset":"transactions","field":"transaction_type",
     "rule_type":"allowed_values","description":"Must be domestic or cross_border","severity":"critical"},
    {"rule_id":"VR004","dataset":"transactions","field":"number_of_transactions",
     "rule_type":"non_negative","description":"Must be >= 0","severity":"critical"},
    {"rule_id":"VR005","dataset":"transactions","field":"total_value_eur_mn",
     "rule_type":"non_negative","description":"Must be >= 0","severity":"critical"},
    {"rule_id":"VR006","dataset":"transactions","field":"quarter",
     "rule_type":"format","description":"Must match YYYYQn pattern","severity":"critical"},
    {"rule_id":"VR007","dataset":"psps","field":"psp_id",
     "rule_type":"uniqueness","description":"PSP ID must be unique","severity":"critical"},
    {"rule_id":"VR008","dataset":"psps","field":"home_country",
     "rule_type":"reference_check","description":"Must be DE","severity":"critical"},
    {"rule_id":"VR009","dataset":"cards","field":"number_of_cards_issued_mn",
     "rule_type":"range","description":"Must be between 0 and 500","severity":"warning"},
    {"rule_id":"VR010","dataset":"transactions","field":"total_value_eur_mn",
     "rule_type":"outlier_zscore","description":"Z-score > 3 flags statistical outlier","severity":"warning"},
]
df_rules = pd.DataFrame(validation_rules)
print(f"4. rules        — {len(df_rules):,} rows")


# =============================================================================
# SAVE TO EXCEL
# =============================================================================
with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    df_transactions.to_excel(writer, sheet_name="payments_transactions_quarterly", index=False)
    df_psps.to_excel(        writer, sheet_name="psp_reference_data",              index=False)
    df_cards.to_excel(       writer, sheet_name="card_payments_detail",            index=False)
    df_rules.to_excel(       writer, sheet_name="validation_rules",                index=False)

print(f"\nSaved → {OUTPUT_PATH}")
print("Run pipeline.py next to process the data.")
