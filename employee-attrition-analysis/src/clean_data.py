from pathlib import Path
import json
import re

import pandas as pd


# Project file paths
PROJECT_ROOT = Path(__file__).resolve().parents[1]

RAW_FILE = (
    PROJECT_ROOT
    / "data"
    / "raw"
    / "HR-Employee-Attrition.csv"
)

CLEAN_FILE = (
    PROJECT_ROOT
    / "data"
    / "processed"
    / "hr_employee_attrition_clean.csv"
)

REPORT_FILE = (
    PROJECT_ROOT
    / "reports"
    / "data_quality_report.json"
)


def to_snake_case(column_name: str) -> str:
    """Convert column names such as MonthlyIncome to monthly_income."""

    first_step = re.sub(
        r"(.)([A-Z][a-z]+)",
        r"\1_\2",
        column_name
    )

    return re.sub(
        r"([a-z0-9])([A-Z])",
        r"\1_\2",
        first_step
    ).lower()


def load_data() -> pd.DataFrame:
    """Load the raw employee attrition dataset."""

    if not RAW_FILE.exists():
        raise FileNotFoundError(
            f"Raw CSV file was not found: {RAW_FILE}"
        )

    return pd.read_csv(RAW_FILE)


def clean_data(df: pd.DataFrame):
    """Clean and prepare the employee attrition data."""

    original_rows = len(df)
    original_columns = len(df.columns)
    missing_before = int(df.isna().sum().sum())
    duplicates_found = int(df.duplicated().sum())

    # Standardize column names
    df.columns = [
        to_snake_case(column)
        for column in df.columns
    ]

    # Remove extra spaces from text columns
    text_columns = df.select_dtypes(
        include=["object", "string"]
    ).columns

    for column in text_columns:
        df[column] = (
            df[column]
            .astype("string")
            .str.strip()
        )

    # Remove completely empty rows
    df = df.dropna(how="all")

    # Remove duplicate rows
    df = df.drop_duplicates().copy()

    # Find columns with only one unique value
    constant_columns = [
        column
        for column in df.columns
        if df[column].nunique(dropna=False) <= 1
    ]

    # Remove constant columns
    df = df.drop(
        columns=constant_columns
    )

    # Convert Yes/No target values into numerical flags
    yes_no_mapping = {
        "Yes": 1,
        "No": 0
    }

    if "attrition" in df.columns:
        df["attrition_flag"] = (
            df["attrition"]
            .map(yes_no_mapping)
            .astype("Int64")
        )

    if "over_time" in df.columns:
        df["overtime_flag"] = (
            df["over_time"]
            .map(yes_no_mapping)
            .astype("Int64")
        )

    attrition_distribution = {}

    if "attrition" in df.columns:
        attrition_distribution = {
            str(category): int(count)
            for category, count
            in df["attrition"].value_counts().items()
        }

    missing_by_column = {
        column: int(count)
        for column, count
        in df.isna().sum().items()
        if count > 0
    }

    quality_report = {
        "original_rows": original_rows,
        "original_columns": original_columns,
        "cleaned_rows": int(df.shape[0]),
        "cleaned_columns": int(df.shape[1]),
        "missing_values_before": missing_before,
        "missing_values_after": int(
            df.isna().sum().sum()
        ),
        "duplicate_rows_found": duplicates_found,
        "duplicate_rows_removed": (
            original_rows - len(df)
        ),
        "constant_columns_removed": constant_columns,
        "missing_values_by_column": missing_by_column,
        "attrition_distribution": attrition_distribution
    }

    return df, quality_report


def save_results(
    cleaned_df: pd.DataFrame,
    quality_report: dict
) -> None:
    """Save the cleaned CSV and data-quality report."""

    CLEAN_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    REPORT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    cleaned_df.to_csv(
        CLEAN_FILE,
        index=False
    )

    with REPORT_FILE.open(
        "w",
        encoding="utf-8"
    ) as report:
        json.dump(
            quality_report,
            report,
            indent=4
        )


def main() -> None:
    print("Loading raw data...")

    raw_df = load_data()

    print(f"Original shape: {raw_df.shape}")

    cleaned_df, quality_report = clean_data(
        raw_df
    )

    save_results(
        cleaned_df,
        quality_report
    )

    print("\nData cleaning completed successfully.")
    print(f"Cleaned shape: {cleaned_df.shape}")

    print(
        "Constant columns removed:",
        quality_report["constant_columns_removed"]
    )

    print(f"Cleaned file: {CLEAN_FILE}")
    print(f"Quality report: {REPORT_FILE}")


if __name__ == "__main__":
    main()
