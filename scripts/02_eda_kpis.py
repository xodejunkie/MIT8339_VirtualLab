#!/usr/bin/env python3
"""
MIT 8339 Virtual Lab - Task 3: Exploratory Analysis, KPI Development,
Risk Scoring and Model Validation
"""
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    classification_report,
    confusion_matrix,
    f1_score,
    precision_score,
    recall_score,
    roc_auc_score,
)
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import CLEANED_CSV, SCORED_CSV, SEGMENTS, KPIS_JSON

df = pd.read_csv(CLEANED_CSV)
N = len(df)


def attrition_table(group_col, order=None):
    g = df.groupby(group_col, observed=True).agg(
        Customers=("CLIENTNUM", "count"),
        Attrited=("Is_Attrited", "sum"),
        Attrition_Rate=("Is_Attrited", "mean"),
        Avg_Credit_Limit=("Credit_Limit", "mean"),
        Avg_Utilisation=("Avg_Utilization_Ratio", "mean"),
        Avg_Trans_Amt=("Total_Trans_Amt", "mean"),
        Avg_Trans_Ct=("Total_Trans_Ct", "mean"),
        Avg_Products=("Total_Relationship_Count", "mean"),
    ).reset_index()
    if order:
        g[group_col] = pd.Categorical(g[group_col], categories=order, ordered=True)
        g = g.sort_values(group_col)
    g["Attrition_Rate"] = g["Attrition_Rate"].round(4)
    for c in ["Avg_Credit_Limit", "Avg_Utilisation", "Avg_Trans_Amt", "Avg_Trans_Ct", "Avg_Products"]:
        g[c] = g[c].round(2)
    return g


# ------------------------------------------------------------ 1. KPIs
# Engaged = High + Very High transaction-frequency bands (aligned with report / dashboard)
engaged_mask = df["Transaction_Frequency"].isin(["High", "Very High"])
engaged_ratio = float(engaged_mask.mean())
dormant_ratio = float((df["Avg_Utilization_Ratio"] == 0).mean())

kpis = {
    "attrition_rate": float(df["Is_Attrited"].mean()),
    "attrited_customers": int(df["Is_Attrited"].sum()),
    "total_customers": N,
    "engaged_customer_ratio": engaged_ratio,
    "card_dormancy_rate": dormant_ratio,
    "avg_credit_limit": float(df["Credit_Limit"].mean()),
    "median_credit_limit": float(df["Credit_Limit"].median()),
    "avg_utilisation_ratio": float(df["Avg_Utilization_Ratio"].mean()),
    "total_trans_amount_portfolio": float(df["Total_Trans_Amt"].sum()),
    "avg_trans_amount_per_customer": float(df["Total_Trans_Amt"].mean()),
    "avg_trans_ct_per_customer": float(df["Total_Trans_Ct"].mean()),
    "avg_products_per_customer": float(df["Total_Relationship_Count"].mean()),
    "avg_revolving_balance": float(df["Total_Revolving_Bal"].mean()),
}
print("=== KPIs (overall) ===")
for k, v in kpis.items():
    print(f"  {k:32s} {v:,.4f}" if isinstance(v, float) else f"  {k:32s} {v:,}")

split = df.groupby("Attrition_Flag").agg(
    n=("CLIENTNUM", "count"),
    avg_credit_limit=("Credit_Limit", "mean"),
    avg_utilisation=("Avg_Utilization_Ratio", "mean"),
    avg_trans_amt=("Total_Trans_Amt", "mean"),
    avg_trans_ct=("Total_Trans_Ct", "mean"),
    avg_products=("Total_Relationship_Count", "mean"),
    avg_revolving_bal=("Total_Revolving_Bal", "mean"),
    avg_months_inactive=("Months_Inactive_12_mon", "mean"),
    avg_contacts=("Contacts_Count_12_mon", "mean"),
    avg_ct_chg_q4q1=("Total_Ct_Chng_Q4_Q1", "mean"),
    avg_amt_chg_q4q1=("Total_Amt_Chng_Q4_Q1", "mean"),
).round(3)
print("\n=== Attrited vs Existing (means) ===")
print(split.to_string())

num_vars = [
    "Credit_Limit", "Avg_Utilization_Ratio", "Total_Trans_Amt", "Total_Trans_Ct",
    "Total_Relationship_Count", "Total_Revolving_Bal", "Months_Inactive_12_mon",
    "Contacts_Count_12_mon", "Total_Ct_Chng_Q4_Q1", "Customer_Age", "Months_on_book",
]
att = df[df["Is_Attrited"] == 1]
exi = df[df["Is_Attrited"] == 0]
tests = {}
for v in num_vars:
    t, p = stats.ttest_ind(att[v], exi[v], equal_var=False)
    tests[v] = {
        "attrited_mean": round(float(att[v].mean()), 3),
        "existing_mean": round(float(exi[v].mean()), 3),
        "difference": round(float(att[v].mean() - exi[v].mean()), 3),
        "welch_t": round(float(t), 2),
        "p_value": float(p),
    }
print("\n=== Welch t-tests (attrited vs existing) ===")
for v, r in tests.items():
    sig = "p<0.001" if r["p_value"] < 0.001 else f"p={r['p_value']:.4f}"
    print(
        f"  {v:26s} attr={r['attrited_mean']:>9.2f} exist={r['existing_mean']:>9.2f} "
        f"diff={r['difference']:>9.2f} t={r['welch_t']:>7.2f} {sig}"
    )

chi2_results = {}
for c in ["Card_Category", "Income_Category", "Education_Level", "Marital_Status", "Gender"]:
    ct = pd.crosstab(df[c], df["Is_Attrited"])
    chi2, p, dof, _ = stats.chi2_contingency(ct)
    chi2_results[c] = {"chi2": round(float(chi2), 2), "p_value": float(p), "dof": int(dof)}
print("\n=== Chi-square tests ===")
for c, r in chi2_results.items():
    print(f"  {c:18s} chi2={r['chi2']:>8.2f} p={r['p_value']:.4f}")

# ------------------------------------------------------------ 2. Segment tables
tables = {
    "by_income": attrition_table(
        "Income_Segment",
        order=[
            "Low (<$40K)", "Lower-Middle ($40-60K)", "Middle ($60-80K)",
            "Upper-Middle ($80-120K)", "High (>$120K)", "Not Declared",
        ],
    ),
    "by_card": attrition_table("Card_Category", order=["Blue", "Silver", "Gold", "Platinum"]),
    "by_age_band": attrition_table("Age_Band", order=["26-34", "35-44", "45-54", "55-64", "65+"]),
    "by_trans_frequency": attrition_table(
        "Transaction_Frequency", order=["Low", "Moderate", "High", "Very High"]
    ),
    "by_utilisation": attrition_table(
        "Utilisation_Band",
        order=[
            "0% (Non-revolver)", "0-10% (Low)", "10-30% (Moderate)",
            "30-50% (Elevated)", ">50% (High)",
        ],
    ),
    "by_products": attrition_table("Total_Relationship_Count"),
    "by_months_inactive": attrition_table("Months_Inactive_12_mon"),
    "by_contacts": attrition_table("Contacts_Count_12_mon"),
    "by_value_class": attrition_table(
        "Customer_Value_Class", order=["Bronze", "Silver", "Gold", "Platinum"]
    ),
}
for name, t in tables.items():
    t.to_csv(SEGMENTS / f"{name}.csv", index=False)

print("\n=== Attrition rate by key segments ===")
for name in [
    "by_trans_frequency", "by_products", "by_utilisation", "by_months_inactive",
    "by_contacts", "by_value_class", "by_income", "by_card", "by_age_band",
]:
    t = tables[name]
    col = t.columns[0]
    print(f"\n-- {name} --")
    print(t[[col, "Customers", "Attrition_Rate"]].to_string(index=False))

watch = df.pivot_table(
    index="Income_Segment", columns="Card_Category",
    values="Is_Attrited", aggfunc="mean", observed=True,
).round(3)
watch_cnt = df.pivot_table(
    index="Income_Segment", columns="Card_Category",
    values="CLIENTNUM", aggfunc="count", observed=True,
)
watch.to_csv(SEGMENTS / "watch_income_card_rate.csv")
watch_cnt.to_csv(SEGMENTS / "watch_income_card_count.csv")

# ------------------------------------------------------------ 3. Rule-based risk score
def score_row(r):
    s = 0
    u = r.Avg_Utilization_Ratio
    s += 2 if u == 0 else (1 if u <= 0.10 else 0)
    s += 2 if r.Months_Inactive_12_mon >= 4 else (1 if r.Months_Inactive_12_mon >= 2 else 0)
    s += 2 if r.Total_Trans_Ct < 45 else (1 if r.Total_Trans_Ct <= 67 else 0)
    s += 2 if r.Total_Ct_Chng_Q4_Q1 < 0.75 else (1 if r.Total_Ct_Chng_Q4_Q1 <= 1.0 else 0)
    s += 2 if r.Total_Relationship_Count <= 2 else (1 if r.Total_Relationship_Count == 3 else 0)
    s += 2 if r.Contacts_Count_12_mon >= 4 else (1 if r.Contacts_Count_12_mon >= 2 else 0)
    return s


df["Risk_Score"] = df.apply(score_row, axis=1)
df["Risk_Score_Pct"] = (df["Risk_Score"] / 12 * 100).round(1)

for cuts in [(4, 7), (3, 6), (4, 8)]:
    lo, hi = cuts
    b = pd.cut(df["Risk_Score"], bins=[-0.1, lo, hi, 12], labels=["Low", "Medium", "High"])
    ar = df.groupby(b, observed=True)["Is_Attrited"].mean()
    cnts = b.value_counts()
    print(
        f"\ncuts (Low<= {lo}, Med {lo+1}-{hi}, High >{hi}): "
        + " ".join(f"{k}={v:.3f}(n={cnts.get(k, 0)})" for k, v in ar.items())
    )

df["Churn_Risk_Band"] = pd.cut(
    df["Risk_Score"], bins=[-0.1, 4, 7, 12], labels=["Low", "Medium", "High"]
)

risk = df.groupby("Churn_Risk_Band", observed=True).agg(
    Customers=("CLIENTNUM", "count"),
    Attrited=("Is_Attrited", "sum"),
    Attrition_Rate=("Is_Attrited", "mean"),
    Avg_Credit_Limit=("Credit_Limit", "mean"),
    Avg_Utilisation=("Avg_Utilization_Ratio", "mean"),
    Avg_Trans_Ct=("Total_Trans_Ct", "mean"),
    Avg_Products=("Total_Relationship_Count", "mean"),
).round(4).reset_index()
risk["Share_of_Base"] = (risk["Customers"] / N).round(4)
risk.to_csv(SEGMENTS / "by_risk_band.csv", index=False)
print("\n=== Risk band summary (Low<=4, Medium 5-7, High>=8) ===")
print(risk.to_string(index=False))

high = df[df["Churn_Risk_Band"] == "High"]
print(
    f"\nHigh-risk customers: {len(high)} ({len(high)/N:.1%} of base); "
    f"capture of all attrited: {high['Is_Attrited'].sum()}/{df['Is_Attrited'].sum()} "
    f"({high['Is_Attrited'].sum()/df['Is_Attrited'].sum():.1%})"
)
print("High-risk by value class:", high["Customer_Value_Class"].value_counts().to_dict())
print("High-risk attrition rate by value class:")
print(
    df[df.Churn_Risk_Band == "High"]
    .groupby("Customer_Value_Class", observed=True)["Is_Attrited"]
    .mean()
    .round(3)
    .to_string()
)

# ------------------------------------------------------------ 4. Logistic regression (validation)
feat = [
    "Total_Relationship_Count", "Months_Inactive_12_mon", "Contacts_Count_12_mon",
    "Credit_Limit", "Total_Revolving_Bal", "Avg_Utilization_Ratio",
    "Total_Trans_Amt", "Total_Trans_Ct", "Total_Ct_Chng_Q4_Q1", "Total_Amt_Chng_Q4_Q1",
]
X, y = df[feat], df["Is_Attrited"]
X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.30, stratify=y, random_state=42)
scaler = StandardScaler().fit(X_tr)
model = LogisticRegression(max_iter=1000, random_state=42).fit(scaler.transform(X_tr), y_tr)
proba = model.predict_proba(scaler.transform(X_te))[:, 1]
y_pred = (proba >= 0.5).astype(int)
auc = roc_auc_score(y_te, proba)
acc = accuracy_score(y_te, y_pred)
# Attrited class (=1) metrics — accuracy alone is weak under class imbalance
prec = precision_score(y_te, y_pred, pos_label=1, zero_division=0)
rec = recall_score(y_te, y_pred, pos_label=1, zero_division=0)
f1 = f1_score(y_te, y_pred, pos_label=1, zero_division=0)
cm = confusion_matrix(y_te, y_pred, labels=[0, 1])
tn, fp, fn, tp = (int(x) for x in cm.ravel())
coefs = pd.Series(model.coef_[0], index=feat)
odds = np.exp(coefs)
print("\n=== Logistic regression (30% stratified holdout — internal validation only) ===")
print(f"AUC={auc:.3f}  accuracy={acc:.3f}")
print(f"Attrited class: precision={prec:.3f}  recall={rec:.3f}  F1={f1:.3f}")
print("Confusion matrix [rows=actual Existing/Attrited; cols=pred Existing/Attrited]:")
print(f"  TN={tn}  FP={fp}")
print(f"  FN={fn}  TP={tp}")
print(classification_report(y_te, y_pred, target_names=["Existing", "Attrited"], digits=3))
print("Standardised coefficients (odds ratios):")
for f in coefs.abs().sort_values(ascending=False).index:
    print(f"  {f:26s} beta={coefs[f]:+.3f}  OR={odds[f]:.2f}")

kpis["high_risk_share"] = float((df["Churn_Risk_Band"] == "High").mean())
kpis["high_risk_customers"] = int((df["Churn_Risk_Band"] == "High").sum())

model_out = {
    "features": feat,
    "auc": round(float(auc), 3),
    "accuracy": round(float(acc), 3),
    "precision_attrited": round(float(prec), 3),
    "recall_attrited": round(float(rec), 3),
    "f1_attrited": round(float(f1), 3),
    "confusion_matrix": {
        "labels": ["Existing(0)", "Attrited(1)"],
        "tn": tn, "fp": fp, "fn": fn, "tp": tp,
    },
    "note": (
        "Internal random holdout only; not evidence of live future performance "
        "in a Nigerian bank. Rule-based bands are a historical retention-priority "
        "score evaluated retrospectively on the same snapshot."
    ),
    "coefficients": {f: round(float(c), 3) for f, c in coefs.items()},
    "odds_ratios": {f: round(float(o), 2) for f, o in odds.items()},
    "train_n": int(len(y_tr)),
    "test_n": int(len(y_te)),
    "random_state": 42,
}

with open(KPIS_JSON, "w") as fh:
    json.dump(
        {
            "kpis": kpis,
            "attrited_vs_existing": split.reset_index().to_dict("records"),
            "welch_tests": tests,
            "chi_square": chi2_results,
            "risk_bands": risk.to_dict("records"),
            "risk_score_weights": {
                "utilisation_zero_or_le010": [2, 1],
                "months_inactive_ge4_or_2to3": [2, 1],
                "trans_ct_lt45_or_le67": [2, 1],
                "ct_chg_q4q1_lt075_or_le10": [2, 1],
                "products_le2_or_eq3": [2, 1],
                "contacts_ge4_or_2to3": [2, 1],
                "max_score": 12,
                "bands": "Low<=4, Medium 5-7, High>=8",
            },
            "model": model_out,
        },
        fh,
        indent=2,
    )

df.to_csv(SCORED_CSV, index=False)
print(f"\n[SAVE] scored dataset -> {SCORED_CSV}")
print(f"[SAVE] KPI pack -> {KPIS_JSON} ; segment tables -> {SEGMENTS}/")
