# Parastatal X Dynamic Budget Monitoring, Forecasting and Management Dashboard

## A Simulated Kenyan Public-Sector Financial Analytics Solution

> **⚠️ DISCLAIMER:** All data in this project is entirely synthetic and generated programmatically using Python (random seed 42). No real persons, institutions, suppliers, account numbers, government documents, or confidential data from any Kenyan public institution are represented. This is a portfolio and demonstration project only.

---

## 📋 Project Overview

This project demonstrates how a Kenyan State Corporation can use **data analytics, Power BI, Power Query, Excel, DAX, and financial forecasting** to strengthen budget control, expenditure monitoring, accountability, management reporting, and evidence-based decision-making.

The fictional organisation — **Parastatal X** — is a Kenyan State Corporation responsible for coordinating national public-service infrastructure, regional operations, research, stakeholder engagement, and service-delivery programmes.

---

## 🎯 Business Problem

Kenyan public sector organisations face persistent challenges in budget monitoring and expenditure control:

- **Fragmented data** across multiple systems with no integrated view
- **Delayed reporting** preventing timely management intervention
- **Approval bottlenecks** causing AIE processing delays
- **Commitment exposure** not visible until it crystallises into expenditure
- **Revenue underperformance** not detected early enough for corrective action
- **Audit findings** repeated year after year due to weak controls
- **Forecast blindness** — no systematic year-end expenditure forecasting

This project builds a **near-real-time financial monitoring dashboard** that addresses all of these challenges.

---

## 🗂️ Project Structure

```
Parastatal_X_Project/
├── 01_Data/                    # All datasets (CSV + Excel workbook)
│   ├── Approved_Annual_Budget.csv
│   ├── Monthly_Budget_Allocation.csv
│   ├── Supplementary_Budget.csv
│   ├── Budget_Reallocation.csv
│   ├── Expenditure_Transactions.csv       # 12,100 rows
│   ├── Commitments.csv
│   ├── Budget_Review_Requests.csv
│   ├── Procurement_Plan.csv
│   ├── Pending_Bills.csv
│   ├── Revenue_AIA.csv
│   ├── Performance_Indicators.csv
│   ├── Audit_Control_Exceptions.csv
│   ├── Data_Quality_Log.csv
│   ├── Dim_*.csv                          # 14 dimension tables
│   └── Parastatal_X_Financial_Monitoring_Data.xlsx   # 26-sheet workbook
├── 02_Code/                    # Python scripts
│   ├── generate_synthetic_data.py         # Main data generation
│   ├── validate_data.py                   # 36 automated validation checks
│   ├── financial_forecasting.py           # 4-method forecast model
│   ├── reconciliation_checks.py           # 18 reconciliation checks
│   ├── build_excel_fast.py                # Excel workbook builder
│   ├── build_docs.py                      # DOCX file builder
│   ├── build_analysis_docs.py             # Analysis documents builder
│   └── requirements.txt
├── 03_PowerBI_Build_Pack/      # Power BI implementation files
│   ├── DAX_Measures.txt                   # 70+ DAX measures
│   ├── Power_Query_M_Code.txt             # 500 lines of M code
│   ├── Theme_JSON.json                    # Custom Power BI theme
│   ├── Data_Model_Specification.docx      # Star schema documentation
│   └── PowerBI_Implementation_Guide.docx  # Step-by-step build guide
├── 04_Analysis/                # Management analysis outputs
│   ├── Senior_Finance_Manager_Analysis.docx
│   ├── Exploratory_Data_Analysis_Report.docx
│   ├── Management_Action_Plan.xlsx
│   ├── Financial_Risk_Register.xlsx
│   └── Reconciliation_Report.xlsx
├── 05_Portfolio/               # Portfolio and career materials
│   ├── GitHub_README.md
│   ├── LinkedIn_Project_Description.txt
│   ├── LinkedIn_Post.txt
│   ├── Project_Case_Study.docx
│   └── Interview_Talking_Points.docx
└── 06_Presentation/            # Executive presentation
    └── Parastatal_X_Budget_Dashboard_Presentation.pptx
```

---

## 📊 Dataset Summary

| Dataset | Rows | Key Metrics |
|---|---|---|
| Approved Annual Budget | 96 | KSh 4.65B across 3 FYs |
| Monthly Budget Allocation | 1,152 | 100% reconciliation to annual budgets |
| Supplementary Budget | 209 | 65% increases, 35% reductions |
| Budget Reallocation | 224 | 75% approved, 25% pending/rejected |
| Expenditure Transactions | 12,100 | KSh 1.478B YTD (FY2025/26) |
| Commitments | 1,350 | KSh 700M open commitments |
| Budget Review Requests | 1,000 | KSh 2.48B backlog |
| Procurement Plan | 390 | 12 contracts delayed |
| Pending Bills | 520 | KSh 1.12B outstanding |
| Revenue/AIA | 1,140 | 88.2% achievement |
| Performance Indicators | 165 | 15 KPIs across 3 FYs |
| Audit Exceptions | 300 | 20% repeat findings |
| Data Quality Log | 250 | 96.2% quality score |

---

## 🏗️ Data Architecture

### Star Schema Design
- **12 Fact Tables** connected to shared dimension tables
- **14 Dimension Tables** including Dim_Date, Dim_Cost_Centre, Dim_Vote, Dim_Region
- Single date spine through `DateKey` (YYYYMMDD integer)
- Financial Year months numbered July=1 through June=12
- Bidirectional filtering disabled to prevent ambiguous results

### Data Flow
```
Python (data generation) → CSV files → Power Query (cleaning/transformation) 
→ Star Schema → DAX Measures → Power BI Dashboard
```

---

## 🛠️ Tools and Technologies

| Category | Tools |
|---|---|
| **Data Generation** | Python, Pandas, NumPy, Random (seed=42) |
| **Data Storage** | CSV, Excel (.xlsx) with 26 sheets |
| **Data Transformation** | Power Query (M language) |
| **Data Modelling** | Star Schema, Power BI relationships |
| **Analytics** | DAX (70+ measures), Python forecasting |
| **Visualisation** | Power BI Desktop (build pack provided) |
| **Documentation** | Python-DOCX, XlsxWriter |
| **Presentation** | Python-PPTX, DOCX |

---

## 💡 Key Financial Findings

### FY2025/2026 (as at 31 May 2026)
1. **Budget Utilisation:** 87.0% (KSh 1.478B / KSh 1.700B) vs 91.7% expected
2. **Forecast Year-End:** KSh 1.763B (103.7%) — slight over-expenditure risk of KSh 63.5M
3. **Open Commitments:** KSh 700M representing major cash exposure
4. **Revenue Achievement:** 88.2% (KSh 130.9M shortfall)
5. **Pending Bills:** KSh 1.12B in uncleared invoices
6. **Data Quality:** 96.2% (above the 95% amber threshold)

---

## 📈 Forecasting Methodology

Four forecasting methods are applied and blended:

| Method | Formula | Year-End Forecast | Weight |
|---|---|---|---|
| Straight-Line Run Rate | YTD / Months_Elapsed × 12 | KSh 1,612,762,691 | 30% |
| 3-Month Moving Average | YTD + (MA3 × remaining months) | KSh 1,653,229,667 | 25% |
| Seasonal Adjustment | Based on historical July-June pattern | KSh 1,722,723,783 | 25% |
| Commitment-Adjusted | YTD + Open Commitments | KSh 2,178,365,800 | 20% |
| **Blended Selected** | **Weighted average** | **KSh 1,763,490,330** | **100%** |

---

## 🔴🟡🟢 Traffic Light Rules

| Metric | 🟢 Green | 🟡 Amber | 🔴 Red |
|---|---|---|---|
| Budget Absorption | Within ±10% of time | 10–20% gap | >20% gap |
| Forecast Over-exp | ≤100% of budget | 100–105% | >105% |
| Request Aging | 0–14 days | 15–30 days | >30 days |
| Pending Bills | 0–30 days | 31–90 days | >90 days |
| Data Quality | ≥98% | 95–98% | <95% |
| Revenue Achievement | ≥90% | 70–90% | <70% |

---

## 📋 Dashboard Pages (12 Pages)

1. **Executive Financial Overview** — KPI cards, waterfall, monthly trend
2. **Budget Performance** — Absorption analysis, variance, drill-down
3. **Budget Review Workflow** — AIE processing, aging, turnaround
4. **Commitments and Cash Exposure** — Open commitments, risk
5. **Procurement Monitoring** — Plan vs contract vs paid
6. **Pending Bills** — Aging analysis, reasons, priority
7. **Revenue and AIA** — Collection performance, banking delays
8. **Performance and Value for Money** — KPIs, 4-quadrant, cost per output
9. **Forecasting and Financial Risk** — Scenarios, year-end projections
10. **Audit, Controls and Data Quality** — Exceptions, DQ score
11. **Cost Centre Drill-Through** — Full CC detail view
12. **Management Action Register** — Live action tracking

---

## 🚀 How to Use

### 1. Reproduce the Data
```bash
pip install -r 02_Code/requirements.txt
python 02_Code/generate_synthetic_data.py
python 02_Code/validate_data.py
python 02_Code/financial_forecasting.py
python 02_Code/reconciliation_checks.py
```

### 2. Open the Excel Workbook
Open `01_Data/Parastatal_X_Financial_Monitoring_Data.xlsx` — no additional software required beyond Microsoft Excel.

### 3. Build in Power BI
1. Open Power BI Desktop
2. Import CSV files using queries from `03_PowerBI_Build_Pack/Power_Query_M_Code.txt`
3. Apply theme from `03_PowerBI_Build_Pack/Theme_JSON.json`
4. Create relationships per `03_PowerBI_Build_Pack/Data_Model_Specification.docx`
5. Add DAX measures from `03_PowerBI_Build_Pack/DAX_Measures.txt`
6. Build 12 dashboard pages per `03_PowerBI_Build_Pack/PowerBI_Implementation_Guide.docx`

---

## ⚖️ Ethical Considerations and Limitations

- **All data is synthetic.** No real government data, personnel, or institutional information has been used.
- **Not an official government system.** This is a simulated portfolio project and should not be represented as an official Kenyan government tool.
- **Data accuracy.** Intentional data quality issues (2.6%) were introduced for demonstration purposes. The cleaned dataset corrects these.
- **Forecast limitations.** Forecasts are modelling exercises based on synthetic data. They should not be used for real financial decisions.
- **Scale assumptions.** Budget amounts and transaction volumes are illustrative of a mid-sized Kenyan parastatal.

---

## 👤 Author

**Portfolio Project — Financial Analytics and Business Intelligence**  
Tools: Python • Power BI • DAX • Power Query • Excel • Python-DOCX • XlsxWriter  
Domain: Public Finance • Budget Management • Kenyan Public Sector Analytics  
Date: June 2026

---

*This project is released for educational and portfolio demonstration purposes. All data is synthetic.*
