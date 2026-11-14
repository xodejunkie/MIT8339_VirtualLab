#!/usr/bin/env python3
"""
MIT 8339 Virtual Lab - Task 2: Data Cleaning and Transformation
Dataset: Credit Card Customers / BankChurners (Kaggle, sakshigoyal7/credit-card-customers)

Pipeline: load -> profile (before) -> clean/standardise -> derive analytical fields
          -> profile (after) -> export cleaned dataset + quality log.
"""
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import RAW_CSV, CLEANED_CSV, CLEANING_LOG, QUALITY_CHECKS

pd.set_option("display.width", 200)

# ---------------------------------------------------------------- 1. LOAD
raw = pd.read_csv(RAW_CSV)
rows_before, cols_before = raw.shape
print(f"[LOAD] raw shape: {rows_before} rows x {cols_before} cols")

# ---------------------------------------------------------------- 2. PROFILE (BEFORE)
before_profile = {
    "rows": int(rows_before),
    "cols": int(cols_before),
    "duplicate_clientnum": int(raw["CLIENTNUM"].duplicated().sum()),
    "full_duplicate_rows": int(raw.duplicated().sum()),
    "missing_cells_total": int(raw.isna().sum().sum()),
    "missing_by_col": {k: int(v) for k, v in raw.isna().sum().items() if v > 0},
    "dtypes": {k: str(v) for k, v in raw.dtypes.items()},
}

unknown_labels = {}
for col in ["Education_Level", "Marital_Status", "Income_Category"]:
    n = int((raw[col] == "Unknown").sum())
    unknown_labels[col] = n
print(f"[PROFILE] Unknown masked labels: {unknown_labels}")


def iqr_outliers(s: pd.Series):
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    return int(((s < lo) | (s > hi)).sum()), float(lo), float(hi)


outlier_vars = [
    "Credit_Limit", "Total_Revolving_Bal", "Total_Trans_Amt",
    "Total_Trans_Ct", "Avg_Utilization_Ratio", "Customer_Age",
    "Months_on_book", "Total_Amt_Chng_Q4_Q1", "Total_Ct_Chng_Q4_Q1",
]
outliers_before = {v: iqr_outliers(raw[v]) for v in outlier_vars}
print("[PROFILE] IQR outlier counts (count, low_fence, high_fence):")
for k, v in outliers_before.items():
    print(f"   {k:26s} {v}")

# ---------------------------------------------------------------- 3. CLEAN
df = raw.copy()

nb_cols = [c for c in df.columns if c.startswith("Naive_Bayes_Classifier")]
df = df.drop(columns=nb_cols)
print(f"[CLEAN] dropped leaked NB columns: {len(nb_cols)}")

df.columns = [c.strip() for c in df.columns]

cat_cols = [c for c in df.columns if df[c].dtype == object or str(df[c].dtype) == "string"]
for c in cat_cols:
    df[c] = df[c].astype(str).str.strip()

dups = int(df["CLIENTNUM"].duplicated().sum())
if dups:
    df = df.drop_duplicates(subset="CLIENTNUM", keep="first")
print(f"[CLEAN] duplicate CLIENTNUM rows removed: {dups}")

assert (df["Customer_Age"].between(18, 100)).all(), "implausible age found"
num_money = ["Credit_Limit", "Total_Revolving_Bal", "Total_Trans_Amt", "Avg_Open_To_Buy"]
for c in num_money:
    assert (df[c] >= 0).all(), f"negative value found in {c}"

df["Is_Attrited"] = (df["Attrition_Flag"] == "Attrited Customer").astype(int)

# ---------------------------------------------------------------- 4. DERIVE ANALYTICAL FIELDS
df["Age_Band"] = pd.cut(
    df["Customer_Age"],
    bins=[25, 34, 44, 54, 64, 200],
    labels=["26-34", "35-44", "45-54", "55-64", "65+"],
)

income_map = {
    "Less than $40K": "Low (<$40K)",
    "$40K - $60K": "Lower-Middle ($40-60K)",
    "$60K - $80K": "Middle ($60-80K)",
    "$80K - $120K": "Upper-Middle ($80-120K)",
    "$120K +": "High (>$120K)",
    "Unknown": "Not Declared",
}
df["Income_Segment"] = df["Income_Category"].map(income_map)


def util_band(r):
    if r == 0:
        return "0% (Non-revolver)"
    if r <= 0.10:
        return "0-10% (Low)"
    if r <= 0.30:
        return "10-30% (Moderate)"
    if r <= 0.50:
        return "30-50% (Elevated)"
    return ">50% (High)"


df["Utilisation_Band"] = df["Avg_Utilization_Ratio"].apply(util_band)

# Activity bands share the cut-offs used by risk rule R3: Low under 45,
# Moderate 45-67, High 68-81, Very High above 81. On this dataset 45, 67 and 81
# are also the quartiles of Total_Trans_Ct.
q1, q2, q3 = df["Total_Trans_Ct"].quantile([0.25, 0.5, 0.75])


def txn_freq(ct):
    if ct < 45:
        return "Low"
    if ct <= 67:
        return "Moderate"
    if ct <= 81:
        return "High"
    return "Very High"


df["Transaction_Frequency"] = df["Total_Trans_Ct"].apply(txn_freq)
print(
    f"[DERIVE] transaction bands: Low<45, Moderate 45-67, High 68-81, Very High>81 "
    f"(quartile refs Q1={q1:.0f}, Q2={q2:.0f}, Q3={q3:.0f})"
)


def tercile(s):
    return pd.qcut(s, 3, labels=[1, 2, 3]).astype(int)


value_score = (
    tercile(df["Total_Trans_Amt"])
    + tercile(df["Credit_Limit"])
    + tercile(df["Total_Relationship_Count"])
)
value_map = {
    3: "Bronze", 4: "Bronze", 5: "Silver", 6: "Silver",
    7: "Gold", 8: "Platinum", 9: "Platinum",
}
df["Customer_Value_Class"] = value_score.map(value_map)

print(
    "[DERIVE] fields added: Age_Band, Income_Segment, Utilisation_Band, "
    "Transaction_Frequency, Customer_Value_Class, Is_Attrited"
)

# ---------------------------------------------------------------- 5. PROFILE (AFTER)
after_profile = {
    "rows": int(df.shape[0]),
    "cols": int(df.shape[1]),
    "missing_cells_total": int(df.isna().sum().sum()),
    "duplicate_clientnum": int(df["CLIENTNUM"].duplicated().sum()),
}

qc_rows = [
    ["Rows", rows_before, df.shape[0], "No rows removed - duplicates absent"],
    ["Columns", cols_before, df.shape[1], f"Dropped {len(nb_cols)} leaked Naive Bayes columns; added 6 analytical fields"],
    ["Missing cells (NaN)", before_profile["missing_cells_total"], after_profile["missing_cells_total"], "None at both stages; 'Unknown' labels treated as explicit category"],
    ["Duplicate CLIENTNUM", before_profile["duplicate_clientnum"], after_profile["duplicate_clientnum"], "None at both stages"],
    ["Education_Level 'Unknown'", unknown_labels["Education_Level"], int((df["Education_Level"] == "Unknown").sum()), "Retained as explicit 'Not Declared' style category - imputation would fabricate data"],
    ["Marital_Status 'Unknown'", unknown_labels["Marital_Status"], int((df["Marital_Status"] == "Unknown").sum()), "Retained as explicit category"],
    ["Income_Category 'Unknown'", unknown_labels["Income_Category"], int((df["Income_Category"] == "Unknown").sum()), "Retained and mapped to 'Not Declared' income segment"],
    ["Credit_Limit IQR outliers", outliers_before["Credit_Limit"][0], outliers_before["Credit_Limit"][0], "Retained - genuine premium-card customers"],
    ["Total_Trans_Amt IQR outliers", outliers_before["Total_Trans_Amt"][0], outliers_before["Total_Trans_Amt"][0], "Retained - genuine heavy transactors"],
    ["Avg_Utilization_Ratio IQR outliers", outliers_before["Avg_Utilization_Ratio"][0], outliers_before["Avg_Utilization_Ratio"][0], "Retained - key risk signal, do not cap"],
    ["Customer_Age range", f"{raw['Customer_Age'].min()}-{raw['Customer_Age'].max()}", f"{df['Customer_Age'].min()}-{df['Customer_Age'].max()}", "Plausible range, no correction needed"],
]
qc = pd.DataFrame(qc_rows, columns=["Check", "Before", "After", "Treatment / Justification"])
qc.to_csv(QUALITY_CHECKS, index=False)

log = {
    "dataset": "Credit Card Customers (BankChurners) - Kaggle sakshigoyal7/credit-card-customers",
    "mirror_source": "https://raw.githubusercontent.com/EfemenaIK/Bank-Customer-Churn/main/BankChurners.csv",
    "before": before_profile,
    "unknown_labels": unknown_labels,
    "outliers_iqr": {k: {"count": v[0], "low": v[1], "high": v[2]} for k, v in outliers_before.items()},
    "columns_dropped": nb_cols,
    "columns_added": [
        "Is_Attrited", "Age_Band", "Income_Segment", "Utilisation_Band",
        "Transaction_Frequency", "Customer_Value_Class",
    ],
    "transaction_quartiles": {"Q1": float(q1), "Q2": float(q2), "Q3": float(q3)},
    "after": after_profile,
    "attrition_rate_overall": float(df["Is_Attrited"].mean()),
}
with open(CLEANING_LOG, "w") as f:
    json.dump(log, f, indent=2, default=str)

df.to_csv(CLEANED_CSV, index=False)
print(f"[SAVE] cleaned dataset -> {CLEANED_CSV} ({df.shape[0]} rows x {df.shape[1]} cols)")
print(f"[SAVE] quality log -> {CLEANING_LOG} ; QC table -> {QUALITY_CHECKS}")
print(f"[CHECK] attrition rate: {df['Is_Attrited'].mean():.4f} ({df['Is_Attrited'].sum()} attrited / {len(df)})")
print("\n[VALUE CLASS] distribution:")
print(df["Customer_Value_Class"].value_counts().to_string())
print("\n[QC TABLE]")
print(qc.to_string(index=False))
