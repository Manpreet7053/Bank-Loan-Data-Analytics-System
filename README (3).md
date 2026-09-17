# 🏦 Bank Loan Data Analytics System

> **An end-to-end financial data analytics pipeline** — from raw Excel data through SQL, Python EDA, Excel Dashboards, and Power BI — processing **38,576 loan records** to deliver actionable KPIs and interactive dashboards.

---

## 📌 Project Overview

This project is the **Final Year B.Tech Project (Computer Science & Engineering)** that builds a complete **Bank Loan Data Analytics System**. It demonstrates how a modern analytics stack can be used to monitor loan portfolio performance, classify loan quality, and provide multi-dimensional business insights using industry-standard tools.

| Metric | Value |
|---|---|
| 📋 Total Loan Records | **38,576** |
| 📊 Total Attributes | **25 columns** |
| 💰 Total Funded Amount | **$435.76M** |
| 💵 Total Amount Received | **$473.07M** |
| 📈 Average Interest Rate | **12.05%** |
| ✅ Good Loan Rate | **86.18%** |
| ❌ Bad Loan Rate | **13.82%** |

---

## 🖼️ Dashboard Screenshots

### Power BI — Summary Dashboard
![Summary Dashboard](screenshots/powerbi_summary.png)

### Power BI — Overview Dashboard
![Overview Dashboard](screenshots/powerbi_overview.png)

### Power BI — Details Dashboard
![Details Dashboard](screenshots/powerbi_details.png)

### Excel — Summary Dashboard
![Excel Summary](screenshots/excel_summary.png)

### Excel — Overview Dashboard
![Excel Overview](screenshots/excel_overview.png)

---

## 🗂️ Repository Structure

```
bank-loan-analytics/
│
├── 📁 data/
│   └── financial_loan_data_excel.xlsx     # Raw dataset (38,576 records, 25 columns)
│
├── 📁 sql/
│   └── queries.sql                        # All 25+ SQL queries (KPIs, Overview, Grid View)
│
├── 📁 python/
│   └── Bank_Loan.ipynb                    # Jupyter Notebook — EDA + 8 visualizations
│
├── 📁 powerbi/
│   └── DASHBOARD.pbix                     # Power BI file — 3 interactive dashboards
│
├── 📁 screenshots/
│   ├── powerbi_summary.png
│   ├── powerbi_overview.png
│   ├── powerbi_details.png
│   ├── excel_summary.png
│   └── excel_overview.png
│
├── 📁 report/
│   └── Bank_Loan_Final_Year_Project_Report.pdf   # Complete project report (55 pages)
│
└── README.md
```

---

## 🛠️ Tech Stack

| Layer | Tool | Purpose |
|---|---|---|
| 🗄️ Database | **PostgreSQL 15** | Relational data storage, 25+ SQL KPI queries |
| 🔧 DB Management | **pgAdmin 4** | GUI for query execution and table management |
| 🐍 Language | **Python 3.11** | EDA, data cleaning, KPI computation, charts |
| 📓 IDE | **Jupyter Notebook** | Interactive Python environment |
| 🐼 Data | **pandas** | DataFrame operations, groupby aggregations |
| 📊 Charts | **matplotlib + seaborn** | Static visualizations (line, bar, pie, area) |
| 🗺️ Interactive | **plotly.express** | Home ownership treemap |
| 📊 BI Dashboard | **Microsoft Power BI Desktop** | 3-page interactive dashboard + DAX measures |
| 📋 Spreadsheet | **Microsoft Excel 2021/365** | Source data + 2 Excel dashboards (Pivot Tables) |
| 🔤 Query Lang | **SQL (PostgreSQL)** | All KPI computations and aggregations |
| 📐 Formula Lang | **DAX** | MTD, MoM, Good/Bad Loan measures in Power BI |

---

## ✅ Data Cleaning Steps

Before any analysis, the raw dataset was cleaned using the following pipeline:

```python
import pandas as pd

df = pd.read_excel('data/financial_loan_data_excel.xlsx', sheet_name='financial_loan')

# Step 1 — Check duplicates
print("Duplicates:", df.duplicated().sum())          # Output: 0

# Step 2 — Handle null values (only emp_title has 1,438 nulls = 3.73%)
df['emp_title'].fillna('Not Specified', inplace=True)

# Step 3 — Trim whitespace in term column
df['term'] = df['term'].str.strip()                  # " 36 months" → "36 months"

# Step 4 — Verify no negative values
print("Negative loan_amount:", (df['loan_amount'] < 0).sum())   # Output: 0
print("Negative int_rate:",    (df['int_rate']    < 0).sum())   # Output: 0
print("Negative dti:",         (df['dti']         < 0).sum())   # Output: 0

# Step 5 — Add derived column (also done in Excel with IF formula)
df['Good vs Bad Loan'] = df['loan_status'].apply(
    lambda x: 'Bad Loan' if x == 'Charged Off' else 'Good Loan'
)

print("Clean dataset shape:", df.shape)   # Output: (38576, 25)
```

| Check | Before | After |
|---|---|---|
| Duplicate rows | 0 | 0 ✅ |
| Null values | 1,438 (emp_title) | 0 ✅ |
| Whitespace in term | Present | Trimmed ✅ |
| Negative numeric values | 0 | 0 ✅ |
| Good vs Bad Loan column | Missing | Added ✅ |

---

## 🗃️ Database Schema

```sql
CREATE TABLE financial_loan (
    id                    BIGINT PRIMARY KEY,
    address_state         VARCHAR(2),
    application_type      VARCHAR(50),
    emp_length            VARCHAR(20),
    emp_title             VARCHAR(255),
    grade                 VARCHAR(2),
    home_ownership        VARCHAR(20),
    issue_date            DATE,
    last_credit_pull_date DATE,
    last_payment_date     DATE,
    loan_status           VARCHAR(50),     -- 'Fully Paid' | 'Current' | 'Charged Off'
    next_payment_date     DATE,
    member_id             BIGINT,
    purpose               VARCHAR(100),
    sub_grade             VARCHAR(5),
    term                  VARCHAR(20),     -- '36 months' | '60 months'
    verification_status   VARCHAR(50),
    annual_income         NUMERIC(15, 2),
    dti                   NUMERIC(10, 4),
    installment           NUMERIC(10, 2),
    int_rate              NUMERIC(10, 4),
    loan_amount           INT,
    total_acc             INT,
    total_payment         NUMERIC(15, 2)
);

SET datestyle = 'ISO, DMY';
```

---

## 📊 SQL KPI Queries

### Summary KPIs (Overall / MTD / PMTD)

```sql
-- Total Loan Applications
SELECT COUNT(id) AS Total_Applications FROM financial_loan;
-- MTD (December)
SELECT COUNT(id) AS MTD_Total_Applications
FROM financial_loan WHERE EXTRACT(MONTH FROM issue_date) = 12;
-- PMTD (November) → MoM Change = +6.9%
SELECT COUNT(id) AS PMTD_Total_Applications
FROM financial_loan WHERE EXTRACT(MONTH FROM issue_date) = 11;

-- Total Funded Amount
SELECT SUM(loan_amount) AS Total_Funded_Amount FROM financial_loan;
-- Result: $435,757,075  |  MTD: $53,981,425  |  MoM: +13.0%

-- Total Amount Received
SELECT SUM(total_payment) AS Total_Amount_Collected FROM financial_loan;
-- Result: $473,070,933  |  MTD: $58,074,380  |  MoM: +15.8%

-- Average Interest Rate
SELECT AVG(int_rate)*100 AS Avg_Int_Rate FROM financial_loan;
-- Result: 12.05%  |  MTD: 12.36%  |  MoM: +3.5%

-- Average DTI
SELECT AVG(dti)*100 AS Avg_DTI FROM financial_loan;
-- Result: 13.33%  |  MTD: 13.67%  |  MoM: +2.7%
```

### Good Loan vs Bad Loan

```sql
-- Good Loan Percentage (Fully Paid + Current)
SELECT
    (COUNT(CASE WHEN loan_status IN ('Fully Paid','Current') THEN id END) * 100.0)
    / COUNT(id) AS Good_Loan_Percentage
FROM financial_loan;
-- Result: 86.18%  |  Applications: 33,243  |  Funded: $370.2M  |  Received: $435.8M

-- Bad Loan Percentage (Charged Off)
SELECT
    (COUNT(CASE WHEN loan_status = 'Charged Off' THEN id END) * 100.0)
    / COUNT(id) AS Bad_Loan_Percentage
FROM financial_loan;
-- Result: 13.82%  |  Applications: 5,333  |  Funded: $65.5M  |  Received: $37.3M  |  Recovery: 56.9%
```

### Loan Status Grid View

```sql
SELECT
    loan_status,
    COUNT(id)           AS Total_Loan_Applications,
    SUM(total_payment)  AS Total_Amount_Received,
    SUM(loan_amount)    AS Total_Funded_Amount,
    AVG(int_rate * 100) AS Avg_Interest_Rate,
    AVG(dti * 100)      AS Avg_DTI
FROM financial_loan
GROUP BY loan_status;
```

### Overview Queries

```sql
-- Monthly Trend
SELECT EXTRACT(MONTH FROM issue_date) AS Month_Number,
       TO_CHAR(issue_date,'Month') AS Month_Name,
       COUNT(id), SUM(loan_amount), SUM(total_payment)
FROM financial_loan
GROUP BY Month_Number, Month_Name ORDER BY Month_Number;

-- State Analysis
SELECT address_state, COUNT(id), SUM(loan_amount), SUM(total_payment)
FROM financial_loan GROUP BY address_state ORDER BY address_state;

-- Term Analysis
SELECT term, COUNT(id), SUM(loan_amount), SUM(total_payment)
FROM financial_loan GROUP BY term ORDER BY term;

-- Employee Length Analysis
SELECT emp_length, COUNT(id), SUM(loan_amount), SUM(total_payment)
FROM financial_loan GROUP BY emp_length ORDER BY emp_length;

-- Loan Purpose Analysis
SELECT purpose, COUNT(id), SUM(loan_amount), SUM(total_payment)
FROM financial_loan GROUP BY purpose ORDER BY purpose;

-- Home Ownership Analysis
SELECT home_ownership, COUNT(id), SUM(loan_amount), SUM(total_payment)
FROM financial_loan GROUP BY home_ownership ORDER BY home_ownership;
```

---

## 🐍 Python EDA — Key Code Snippets

### Setup

```python
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

df = pd.read_excel('data/financial_loan_data_excel.xlsx', sheet_name='financial_loan')
print(df.shape)   # (38576, 25)
```

### Monthly Trend Chart

```python
monthly_funded = (
    df.sort_values('issue_date')
      .assign(month_name=lambda x: x['issue_date'].dt.strftime('%b %Y'))
      .groupby('month_name', sort=False)['loan_amount']
      .sum().div(1_000_000).reset_index(name='loan_amount_millions')
)

plt.figure(figsize=(10, 5))
plt.fill_between(monthly_funded['month_name'],
                 monthly_funded['loan_amount_millions'], color='skyblue', alpha=0.5)
plt.plot(monthly_funded['month_name'], monthly_funded['loan_amount_millions'],
         color='blue', linewidth=2)
plt.title('Total Funded Amount by Month', fontsize=14)
plt.tight_layout()
plt.show()
```

### Good vs Bad Loan Metrics

```python
good_loans = df[df['loan_status'].isin(['Fully Paid', 'Current'])]
bad_loans  = df[df['loan_status'] == 'Charged Off']

print(f"Good Loan %:  {len(good_loans)/len(df)*100:.2f}%")   # 86.18%
print(f"Bad Loan %:   {len(bad_loans)/len(df)*100:.2f}%")    # 13.82%
print(f"Recovery Rate (Bad): {bad_loans['total_payment'].sum()/bad_loans['loan_amount'].sum()*100:.1f}%")  # 56.9%
```

---

## 📐 Power BI — DAX Measures

```dax
-- Core Measures
Total Applications    = COUNT(financial_loan[id])
Total Funded Amount   = SUM(financial_loan[loan_amount])
Total Amount Received = SUM(financial_loan[total_payment])
Avg Interest Rate     = AVERAGE(financial_loan[int_rate]) * 100
Avg DTI               = AVERAGE(financial_loan[dti]) * 100

-- MTD (Month-to-Date)
MTD Applications =
    CALCULATE([Total Applications], DATESMTD(financial_loan[issue_date]))

-- PMTD (Previous Month-to-Date)
PMTD Applications =
    CALCULATE([Total Applications],
              DATEADD(DATESMTD(financial_loan[issue_date]), -1, MONTH))

-- MoM Change %
MoM Applications % =
    ([MTD Applications] - [PMTD Applications]) / [PMTD Applications]

-- Good vs Bad Loan
Good Loan % =
    DIVIDE(
        CALCULATE([Total Applications],
            financial_loan[loan_status] IN {"Fully Paid", "Current"}),
        [Total Applications]) * 100

Bad Loan % =
    DIVIDE(
        CALCULATE([Total Applications],
            financial_loan[loan_status] = "Charged Off"),
        [Total Applications]) * 100
```

---

## 📈 Key Results

### KPI Summary

| KPI | Overall | MTD (Dec) | MoM Change |
|---|---|---|---|
| Total Loan Applications | **38,576** | 4,314 | +6.9% |
| Total Funded Amount | **$435.76M** | $53.98M | +13.0% |
| Total Amount Received | **$473.07M** | $58.07M | +15.8% |
| Average Interest Rate | **12.05%** | 12.36% | +3.5% |
| Average DTI | **13.33%** | 13.67% | +2.7% |

### Loan Quality

| Metric | Good Loans | Bad Loans |
|---|---|---|
| Count | 33,243 (86.18%) | 5,333 (13.82%) |
| Funded Amount | $370.22M | $65.53M |
| Amount Received | $435.79M | $37.28M |
| Recovery Rate | **117.7%** | **56.9%** ⚠️ |
| Credit Loss | — | **~$28.24M** |

### Key Insights

- 📈 **87% growth** in loan applications from January (2,301) to December (4,314)
- 🏙️ **California** is the highest-volume state (6,894 applications)
- 🔄 **Debt consolidation** is the dominant loan purpose — 18,200 applications (47.2%)
- ⏱️ **73.2%** of borrowers prefer 36-month terms over 60-month terms
- 👔 Borrowers with **10+ years of employment** receive the highest total funding
- 🏠 **RENT (18,439)** and **MORTGAGE (17,198)** account for 91%+ of all applications

---

## 🚀 How to Run This Project

### Prerequisites

```bash
# Python packages
pip install pandas numpy matplotlib seaborn plotly openpyxl jupyter

# PostgreSQL (download from https://www.postgresql.org/download/)
# Power BI Desktop (download from https://powerbi.microsoft.com/desktop/)
```

### Step 1 — Set up PostgreSQL Database

```bash
# Open pgAdmin or psql and run:
psql -U postgres -f sql/queries.sql
```

### Step 2 — Load Data into PostgreSQL

```sql
-- In pgAdmin, after running the CREATE TABLE:
-- Use the Import/Export tool to import:
-- File: data/financial_loan_data_excel.xlsx (convert to CSV first)
-- Or use Python:
```

```python
import pandas as pd
from sqlalchemy import create_engine

df = pd.read_excel('data/financial_loan_data_excel.xlsx', sheet_name='financial_loan')

engine = create_engine('postgresql://postgres:password@localhost:5432/your_db')
df.to_sql('financial_loan', engine, if_exists='replace', index=False)
print("Data loaded successfully!")
```

### Step 3 — Run Python EDA

```bash
cd python/
jupyter notebook Bank_Loan.ipynb
```

### Step 4 — Open Excel Dashboards

```
Open: data/financial_loan_data_excel.xlsx
Navigate to: SUMMARY DASHBOARD or OVERVIEW DASHBOARD tabs
Use slicers on the left to filter by Grade and Purpose
```

### Step 5 — Open Power BI Dashboard

```
1. Install Power BI Desktop (free)
2. Open: powerbi/DASHBOARD.pbix
3. Navigate between Summary, Overview, and Details pages
4. Use slicers: State, Grade, Purpose, Good vs Bad Loan
```

---

## 📁 Files Description

| File | Description |
|---|---|
| `data/financial_loan_data_excel.xlsx` | Raw dataset + Excel dashboards (4 sheets) |
| `sql/queries.sql` | Complete SQL script — table creation + all 25+ KPI queries |
| `python/Bank_Loan.ipynb` | Jupyter Notebook — full EDA with 8 charts |
| `powerbi/DASHBOARD.pbix` | Power BI Desktop file — 3 interactive dashboards |
| `report/Bank_Loan_Final_Year_Project_Report.pdf` | Full 55-page project report |

---

## 📚 Project Report

The complete **55-page final year project report** is available in the `report/` folder. It includes:

- Chapter 1: Excel Data Source & Data Cleaning
- Chapter 2: Introduction & Problem Statement
- Chapter 3: Literature Review
- Chapter 4: System Requirements & Technology Stack
- Chapter 5: Database Design & Complete SQL Implementation
- Chapter 6: Python EDA with all 8 charts embedded
- Chapter 7: Microsoft Excel Dashboard
- Chapter 8: Power BI Dashboard Design
- Chapter 9: Results & Analysis
- Chapter 10: Conclusion & Future Work
- Appendices: Full SQL script, Data Dictionary, Project Timeline

---

## 🔮 Future Enhancements

- [ ] **Machine Learning**: Loan default prediction using logistic regression / XGBoost
- [ ] **Real-time Pipeline**: Apache Kafka for live loan data streaming
- [ ] **Cloud Deployment**: AWS RDS (PostgreSQL) + Power BI Service
- [ ] **Advanced DAX**: YoY comparisons, rolling 12-month averages, cohort analysis
- [ ] **Regulatory Reports**: Automated HMDA and CRA compliance reports
- [ ] **NLP Interface**: Power BI Q&A with domain-specific synonyms

---

## 👨‍💻 Author

**[Your Full Name]**
B.Tech Computer Science & Engineering — Final Year
[Your University Name] | 2024–25

📧 [your.email@example.com]
🔗 [LinkedIn Profile URL]
🐙 [GitHub Profile URL]

---

## 🙏 Acknowledgements

- Supervisor: **[Supervisor Name]**, [University Name]
- Tools: PostgreSQL, Python, pandas, matplotlib, seaborn, plotly, Microsoft Power BI, Microsoft Excel
- Dataset: Financial loan dataset used for academic/analytical purposes

---

## 📄 License

This project is submitted as a final year academic project at [University Name].
All rights reserved © 2024–25 [Your Name].

---

<div align="center">
  <b>⭐ If you found this project helpful, please star the repository!</b>
</div>
