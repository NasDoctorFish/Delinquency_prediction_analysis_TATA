""" Explore Geldium delingquency patterns without inventing missing values.

Install: pip install pandas numpy openpyxl
Run: python delinquency_patterns.py customers.csv --output results


"""


import argparse
from pathlib import Path
import numpy as np
import pandas as pd


TARGET = "Delinquent_Account" # The main column for spearman corelation analysis
NUMERIC = ["Age", "Income", "Credit_Store", "Credit_Utilization", "Missed_Payments", "Loan_Balance", "Debt_to_Income_Ratio",
           "Account_Tenure"]
MONTHS = [f"Month_{i}" for i in range(1,7)]


def analyze(path, output, month_order=None, min_group=20):
    path, output = Path(path), Path(output)
    output.mkdir(parents=True, exist_ok=True)
    df = pd.read_excel(path) if path.suffix.lower() in {".xlsx", ".xls"} else pd.read_csv(path)
    df.columns = df.columns.str.strip()
    if df.columns.duplicated().any():
        raise ValueError("Duplicated column names after removing whitespace.")
    for c in df.select_dtypes(include=["object", "string"]):
        df[c] = df[c].astype("string").str.strip().replace("",pd.NA)
    raw = df.copy()
    issues = []

    def flag(column, problem, mask):
        count = int(mask.sum())
        if count:
            issues.append({"column": column, "issue": problem, "rows": count})

    # Invalid numeric values become missing; original data is never overwritten.
    for c in NUMERIC + MONTHS + [TARGET]:
        if c not in df:
            continue
        parsed = pd.to_numeric(df[c], errors="coerce")
        flag(c, "Non-numeric value", df[c].notna() & parsed.isna())
        flag(c, "Infinite value", parsed.isin([np.inf, -np.inf]))
        df[c] = parsed.replace([np.inf, -np.inf], np.nan)
    limits = {"Age": (0, 120), "Income": (0, np.inf),
              "Credit_Score": (300, 850), "Credit_Utilization": (0, 100),
              "Missed_Payments": (0, np.inf), "Loan_Balance": (0, np.inf),
              "Debt_to_Income_Ratio": (0, 100), "Account_Tenure": (0, np.inf)}
    for c, (lo, hi) in limits.items():
        if c in df:
            bad = df[c].notna() & ~df[c].between(lo, hi)
            flag(c, "Outside guide range; verify units/definition", bad)
            df.loc[bad, c] = np.nan
    for c in MONTHS + [TARGET]:
        if c in df:
            bad = df[c].notna() & ~df[c].isin([0, 1] if c == TARGET else [0, 1, 2])
            flag(c, "Invalid status code", bad)
            df.loc[bad, c] = np.nan
    if "Missed_Payments" in df:
        bad = df.Missed_Payments.notna() & (df.Missed_Payments % 1 != 0)
        flag("Missed_Payments", "Non-integer count", bad)
        df.loc[bad, "Missed_Payments"] = np.nan
    flag("all", "Duplicate full row (review before removing)", raw.duplicated())
    if "Customer_ID" in df:
        flag("Customer_ID", "Repeated customer ID (all occurrences)",
             df.Customer_ID.notna() & df.Customer_ID.duplicated(keep=False))
    quality = pd.DataFrame({"raw_missing": raw.isna().sum(),
                            "missing_after_validation": df.isna().sum()})
    quality["missing_pct"] = quality.missing_after_validation / max(len(df), 1) * 100
    quality.to_csv(output / "data_quality.csv", index_label="column")
    pd.DataFrame(issues, columns=["column", "issue", "rows"]).to_csv(output / "validation_issues.csv", index=False)
    
    months = [c for c in MONTHS if c in df]
    if months:
        history = df[months]
        df["Observed_Payment_Months"] = history.notna().sum(axis=1)
        # Require all six months: missing history must not imply no missed payments.
        if len(months) == 6:
            complete = history.notna().all(axis=1)
            df["Late_Months"] = history.eq(1).sum(axis=1).where(complete)
            df["Missed_Months"] = history.eq(2).sum(axis=1).where(complete)
            if month_order:
                ordered = MONTHS if month_order == "oldest-first" else MONTHS[::-1]
                df["Latest_Payment_Status"] = df[ordered[-1]]
                change = df[ordered[-1]] - df[ordered[0]]
                df["Payment_Change"] = change.map(lambda x: "Worsened" if x > 0 else "Improved" if x < 0 else "Unchanged" if pd.notna(x) else pd.NA)

    labeled = df[df[TARGET].notna()].copy()
    if labeled.empty:
        raise ValueError("No valid 0/1 delinquency labels; see quality reports.")
    baseline = labeled[TARGET].mean()
    tables = []

    def rates(feature, groups):
        temp = pd.DataFrame({"group": groups.astype("string").fillna("Missing/Invalid"),
                             "target": labeled[TARGET]})
        table = temp.groupby("group", dropna=False).target.agg(customers="size", delinquent="sum")
        p = table.delinquent / table.customers
        # Wilson 95% intervals make uncertainty in small groups visible.
        n, z = table.customers, 1.96
        center = (p + z*z/(2*n)) / (1 + z*z/n)
        margin = z*np.sqrt(p*(1-p)/n + z*z/(4*n*n)) / (1 + z*z/n)
        table["delinquency_rate_pct"] = p * 100
        table["confidence_interval_95_low_pct"], table["confidence_interval_95_high_pct"] = (center-margin)*100, (center+margin)*100
        table["lift_vs_overall"] = p / baseline if baseline > 0 else np.nan
        table["small_group"] = n < min_group
        table = table.reset_index()
        table.insert(0, "feature", feature)
        tables.append(table)

    # Quantile groups avoid hard-coded thresholds and tolerate tied/constant values.
    for c in NUMERIC:
        if c in labeled and labeled[c].notna().any():
            if labeled[c].nunique() > 1:
                groups = pd.qcut(labeled[c], q=4, duplicates="drop")
                if groups.notna().sum() == 0:
                    groups = labeled[c]
            else:
                groups = labeled[c]
            rates(c, groups)
    for c in ["Employment_Status", "Credit_Card_Type", "Location"] + months + ["Observed_Payment_Months", "Late_Months", "Missed_Months", "Latest_Payment_Status", "Payment_Change"]:
        if c in labeled:
            rates(c, labeled[c])
    result = pd.concat(tables, ignore_index=True) if tables else pd.DataFrame()
    result.to_csv(output / "delinquency_by_group.csv", index=False)
    numeric = [c for c in NUMERIC + ["Late_Months", "Missed_Months"] if c in labeled]
    labeled.groupby(TARGET)[numeric].agg(["count", "mean", "median"]).to_csv(output / "numeric_comparison.csv")
    correlations = labeled[numeric + [TARGET]].corr(method="spearman")[TARGET].drop(TARGET)
    correlations.rename("spearman_correlation").to_csv(output / "numeric_associations.csv", index_label="feature")
    print(f"Rows: {len(df):,}; valid labeled rows: {len(labeled):,}")
    print(f"Overall delinquency rate: {baseline:.2%}")
    if not result.empty:
        reliable = result[(~result.small_group) & (result.group != "Missing/Invalid")]
        print("\nHighest observed rates (adequate-size groups):")
        print(reliable.sort_values("delinquency_rate_pct", ascending=False).head(12).to_string(index=False))
    print(f"\nReports saved to: {output.resolve()}")
    print("Associations only. Confirm payment dates precede the target outcome before modeling.")
    if month_order is None:
        print("Payment trends skipped: set --month-order after confirming chronology.")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("dataset", help="CSV or Excel file")
    parser.add_argument("--output", default="delinquency_results")
    parser.add_argument("--month-order", choices=["oldest-first", "newest-first"])
    parser.add_argument("--min-group", type=int, default=20)
    args = parser.parse_args()
    if args.min_group < 1:
        parser.error("--min-group must be positive")
    analyze(args.dataset, args.output, args.month_order, args.min_group)
