"""
prepare_data.py
Loads the raw IBM HR Attrition CSV and returns a cleaned, analysis-ready table.

Used by both the Jupyter notebook and the Excel dashboard builder,
so the cleaning steps live in one place.
"""

from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RAW_FILE = PROJECT_ROOT / "data" / "WA_Fn-UseC_-HR-Employee-Attrition.csv"
CLEAN_FILE = PROJECT_ROOT / "data" / "hr_attrition_clean.csv"

# The survey-style columns are stored as numbers 1-4 (or 1-5).
# These maps turn them into readable labels, following the dataset's documentation.
LEVEL_LABELS = {1: "Low", 2: "Medium", 3: "High", 4: "Very High"}
WORK_LIFE_LABELS = {1: "Bad", 2: "Good", 3: "Better", 4: "Best"}
EDUCATION_LABELS = {1: "Below College", 2: "College", 3: "Bachelor", 4: "Master", 5: "Doctor"}
TRAVEL_LABELS = {
    "Non-Travel": "Non-Travel",
    "Travel_Rarely": "Travel Rarely",
    "Travel_Frequently": "Travel Frequently",
}

# Columns that hold the same value for every employee carry no information.
CONSTANT_COLUMNS = ["EmployeeCount", "Over18", "StandardHours"]


def load_raw(path: Path = RAW_FILE) -> pd.DataFrame:
    """Read the raw CSV. utf-8-sig removes the invisible BOM character at the start of the file."""
    return pd.read_csv(path, encoding="utf-8-sig")


def clean(raw: pd.DataFrame) -> pd.DataFrame:
    """Return a cleaned copy of the raw data with readable labels and grouping bands."""
    df = raw.copy()

    # 1. Drop columns that never change
    df = df.drop(columns=CONSTANT_COLUMNS)

    # 2. A 0/1 flag makes attrition rates easy to calculate (mean of the flag = rate)
    df["AttritionFlag"] = (df["Attrition"] == "Yes").astype(int)

    # 3. Replace numeric codes with readable labels
    df["JobSatisfactionLabel"] = df["JobSatisfaction"].map(LEVEL_LABELS)
    df["EnvironmentSatisfactionLabel"] = df["EnvironmentSatisfaction"].map(LEVEL_LABELS)
    df["WorkLifeBalanceLabel"] = df["WorkLifeBalance"].map(WORK_LIFE_LABELS)
    df["EducationLabel"] = df["Education"].map(EDUCATION_LABELS)
    df["BusinessTravel"] = df["BusinessTravel"].map(TRAVEL_LABELS)

    # 4. Group continuous values into bands that are easy to compare
    df["AgeBand"] = pd.cut(
        df["Age"], bins=[17, 25, 35, 45, 60], labels=["18-25", "26-35", "36-45", "46-60"]
    )
    df["TenureBand"] = pd.cut(
        df["YearsAtCompany"],
        bins=[-1, 2, 5, 10, 40],
        labels=["0-2 yrs", "3-5 yrs", "6-10 yrs", "10+ yrs"],
    )
    df["IncomeBand"] = pd.qcut(
        df["MonthlyIncome"], q=4, labels=["Q1 (Lowest)", "Q2", "Q3", "Q4 (Highest)"]
    )

    # 5. Store bands as plain text so they export cleanly to CSV and Excel
    for col in ["AgeBand", "TenureBand", "IncomeBand"]:
        df[col] = df[col].astype(str)

    return df


def attrition_by(df: pd.DataFrame, column: str) -> pd.DataFrame:
    """Employees, leavers and attrition rate for each value of `column`, highest rate first."""
    summary = (
        df.groupby(column, observed=True)["AttritionFlag"]
        .agg(Employees="count", Left="sum", AttritionRate="mean")
        .sort_values("AttritionRate", ascending=False)
    )
    summary["AttritionRate"] = summary["AttritionRate"].round(3)
    return summary


if __name__ == "__main__":
    cleaned = clean(load_raw())
    cleaned.to_csv(CLEAN_FILE, index=False)
    print(f"Saved {len(cleaned):,} rows to {CLEAN_FILE.relative_to(PROJECT_ROOT)}")
