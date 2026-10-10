# Data

## Source
**IBM HR Analytics Employee Attrition & Performance**, available on Kaggle:
https://www.kaggle.com/datasets/pavansubhasht/ibm-hr-analytics-attrition-dataset

It is a fictional dataset created by IBM data scientists for learning and practice.
The file is also published in IBM's own GitHub repository:
https://github.com/IBM/employee-attrition-aif360

## Files
| File | Description |
|---|---|
| `WA_Fn-UseC_-HR-Employee-Attrition.csv` | Raw data: 1,470 employees, 35 columns |
| `hr_attrition_clean.csv` | Cleaned data created by the notebook (`src/prepare_data.py`) |

## Key columns
| Column | Meaning |
|---|---|
| `Attrition` | **Target.** Yes = the employee left, No = stayed |
| `OverTime` | Whether the employee works overtime (Yes / No) |
| `JobRole`, `Department`, `JobLevel` | Position in the company (JobLevel 1 = entry level, 5 = most senior) |
| `MonthlyIncome` | Monthly salary |
| `YearsAtCompany` | Tenure in years |
| `JobSatisfaction`, `EnvironmentSatisfaction` | Survey score: 1 Low, 2 Medium, 3 High, 4 Very High |
| `WorkLifeBalance` | Survey score: 1 Bad, 2 Good, 3 Better, 4 Best |
| `Education` | 1 Below College, 2 College, 3 Bachelor, 4 Master, 5 Doctor |

## Columns added during cleaning
`AttritionFlag` (1 = left, 0 = stayed), readable `...Label` versions of the survey scores, and `AgeBand`, `TenureBand`, `IncomeBand` (income quartiles).
