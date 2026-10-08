"""View distinct values per column in the delinquency dataset.

Run: python view_distinct_values.py [path_to_xlsx] [--max N]
"""

import argparse
from pathlib import Path
import pandas as pd


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("dataset", nargs="?", default="Delinquency_prediction_dataset.xlsx")
    parser.add_argument("--max", type=int, default=20, help="Max distinct values to print per column")
    args = parser.parse_args()

    path = Path(args.dataset)
    df = pd.read_excel(path)
    df.columns = df.columns.str.strip()

    for col in df.columns:
        values = df[col].dropna().unique()
        print(f"\n{col}  ({len(values)} distinct, {df[col].isna().sum()} missing)")
        shown = sorted(values.tolist(), key=lambda x: str(x))[: args.max]
        for v in shown:
            print(f"  {v!r}")
        if len(values) > args.max:
            print(f"  ... {len(values) - args.max} more")


if __name__ == "__main__":
    main()
