from pathlib import Path
import json

import matplotlib.pyplot as plt
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]

CLEAN_FILE = (
    PROJECT_ROOT
    / "data"
    / "processed"
    / "hr_employee_attrition_clean.csv"
)

REPORTS_DIR = PROJECT_ROOT / "reports"
FIGURES_DIR = REPORTS_DIR / "figures"


def load_data() -> pd.DataFrame:
    """Load the cleaned employee attrition data."""

    if not CLEAN_FILE.exists():
        raise FileNotFoundError(
            "Cleaned file was not found. Run clean_data.py first."
        )

    return pd.read_csv(CLEAN_FILE)


def save_summary_reports(df: pd.DataFrame) -> None:
    """Create statistical and attrition summary reports."""

    REPORTS_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    # Numeric variables summary
    numeric_summary = (
        df.select_dtypes(include="number")
        .describe()
        .transpose()
    )

    numeric_summary.to_csv(
        REPORTS_DIR / "numeric_summary.csv"
    )

    # Overall attrition counts
    attrition_counts = (
        df["attrition"]
        .value_counts()
        .rename_axis("attrition")
        .reset_index(name="employee_count")
    )

    attrition_counts.to_csv(
        REPORTS_DIR / "attrition_counts.csv",
        index=False
    )

    # Attrition rate by department
    department_summary = (
        df.groupby(
            "department",
            as_index=False
        )
        .agg(
            employees=("employee_number", "count"),
            employees_left=("attrition_flag", "sum"),
            attrition_rate=("attrition_flag", "mean")
        )
    )

    department_summary["attrition_rate"] = (
        department_summary["attrition_rate"] * 100
    ).round(2)

    department_summary.to_csv(
        REPORTS_DIR / "attrition_by_department.csv",
        index=False
    )

    # Attrition rate by job role
    job_role_summary = (
        df.groupby(
            "job_role",
            as_index=False
        )
        .agg(
            employees=("employee_number", "count"),
            employees_left=("attrition_flag", "sum"),
            attrition_rate=("attrition_flag", "mean")
        )
        .sort_values(
            "attrition_rate",
            ascending=False
        )
    )

    job_role_summary["attrition_rate"] = (
        job_role_summary["attrition_rate"] * 100
    ).round(2)

    job_role_summary.to_csv(
        REPORTS_DIR / "attrition_by_job_role.csv",
        index=False
    )

    overall_summary = {
        "total_employees": int(len(df)),
        "employees_left": int(df["attrition_flag"].sum()),
        "employees_stayed": int(
            len(df) - df["attrition_flag"].sum()
        ),
        "overall_attrition_rate_percent": round(
            float(df["attrition_flag"].mean() * 100),
            2
        ),
        "average_age": round(
            float(df["age"].mean()),
            2
        ),
        "average_monthly_income": round(
            float(df["monthly_income"].mean()),
            2
        )
    }

    with (
        REPORTS_DIR / "eda_summary.json"
    ).open("w", encoding="utf-8") as file:
        json.dump(
            overall_summary,
            file,
            indent=4
        )


def create_attrition_distribution(df: pd.DataFrame) -> None:
    """Create attrition employee-count chart."""

    counts = df["attrition"].value_counts()

    plt.figure(figsize=(7, 5))
    counts.plot(kind="bar")

    plt.title("Employee Attrition Distribution")
    plt.xlabel("Attrition")
    plt.ylabel("Number of Employees")
    plt.xticks(rotation=0)
    plt.tight_layout()

    plt.savefig(
        FIGURES_DIR / "attrition_distribution.png",
        dpi=300
    )

    plt.close()


def create_department_chart(df: pd.DataFrame) -> None:
    """Create attrition rate by department chart."""

    department_rates = (
        df.groupby("department")["attrition_flag"]
        .mean()
        .mul(100)
        .sort_values(ascending=False)
    )

    plt.figure(figsize=(9, 5))
    department_rates.plot(kind="bar")

    plt.title("Attrition Rate by Department")
    plt.xlabel("Department")
    plt.ylabel("Attrition Rate (%)")
    plt.xticks(rotation=20, ha="right")
    plt.tight_layout()

    plt.savefig(
        FIGURES_DIR / "attrition_by_department.png",
        dpi=300
    )

    plt.close()


def create_job_role_chart(df: pd.DataFrame) -> None:
    """Create attrition rate by job role chart."""

    job_role_rates = (
        df.groupby("job_role")["attrition_flag"]
        .mean()
        .mul(100)
        .sort_values()
    )

    plt.figure(figsize=(10, 7))
    job_role_rates.plot(kind="barh")

    plt.title("Attrition Rate by Job Role")
    plt.xlabel("Attrition Rate (%)")
    plt.ylabel("Job Role")
    plt.tight_layout()

    plt.savefig(
        FIGURES_DIR / "attrition_by_job_role.png",
        dpi=300
    )

    plt.close()


def create_overtime_chart(df: pd.DataFrame) -> None:
    """Create attrition rate by overtime chart."""

    overtime_rates = (
        df.groupby("over_time")["attrition_flag"]
        .mean()
        .mul(100)
        .sort_values(ascending=False)
    )

    plt.figure(figsize=(7, 5))
    overtime_rates.plot(kind="bar")

    plt.title("Attrition Rate by Overtime")
    plt.xlabel("Works Overtime")
    plt.ylabel("Attrition Rate (%)")
    plt.xticks(rotation=0)
    plt.tight_layout()

    plt.savefig(
        FIGURES_DIR / "attrition_by_overtime.png",
        dpi=300
    )

    plt.close()


def create_income_boxplot(df: pd.DataFrame) -> None:
    """Compare monthly income by attrition status."""

    fig, ax = plt.subplots(figsize=(8, 5))

    df.boxplot(
        column="monthly_income",
        by="attrition",
        ax=ax,
        grid=False
    )

    ax.set_title("Monthly Income by Attrition Status")
    ax.set_xlabel("Attrition")
    ax.set_ylabel("Monthly Income")

    fig.suptitle("")
    fig.tight_layout()

    fig.savefig(
        FIGURES_DIR / "income_by_attrition.png",
        dpi=300
    )

    plt.close(fig)


def main() -> None:
    print("Loading cleaned data...")

    df = load_data()

    FIGURES_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    print(f"Dataset shape: {df.shape}")

    save_summary_reports(df)
    create_attrition_distribution(df)
    create_department_chart(df)
    create_job_role_chart(df)
    create_overtime_chart(df)
    create_income_boxplot(df)

    print("\nExploratory data analysis completed.")
    print(f"Reports saved in: {REPORTS_DIR}")
    print(f"Charts saved in: {FIGURES_DIR}")


if __name__ == "__main__":
    main()
