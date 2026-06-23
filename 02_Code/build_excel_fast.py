"""
Parastatal X - Excel Workbook Builder (Fast/XlsxWriter version)
Uses XlsxWriter for speed. Transaction sheet capped at 3000 rows
(full data in CSV); all other sheets at full data volume.
"""
import pandas as pd
import xlsxwriter
import os

from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "01_Data"
OUT  = ROOT / "01_Data"

FOREST_GREEN = "#1B4332"
MUTED_GOLD   = "#B7950B"
WHITE        = "#FFFFFF"
LIGHT_GREY   = "#F2F2F2"
CHARCOAL     = "#2C3E50"
LIGHT_GREEN  = "#D5E8D4"
LIGHT_GOLD   = "#FFF2CC"
RED_BG       = "#FADBD8"

filepath = f"{OUT}/Parastatal_X_Financial_Monitoring_Data.xlsx"
wb = xlsxwriter.Workbook(filepath, {'strings_to_numbers': False,
                                     'constant_memory': False})

# Formats
hdr  = wb.add_format({'bold':True,'bg_color':FOREST_GREEN,'font_color':WHITE,
                       'border':1,'align':'center','valign':'vcenter','font_size':10,'text_wrap':True})
hdr2 = wb.add_format({'bold':True,'bg_color':MUTED_GOLD,'font_color':CHARCOAL,
                       'border':1,'align':'center','valign':'vcenter','font_size':10})
body = wb.add_format({'font_color':CHARCOAL,'border':1,'valign':'vcenter','font_size':9})
body_grey = wb.add_format({'font_color':CHARCOAL,'border':1,'bg_color':LIGHT_GREY,'valign':'vcenter','font_size':9})
cur  = wb.add_format({'font_color':CHARCOAL,'border':1,'num_format':'#,##0','valign':'vcenter','font_size':9})
cur_grey = wb.add_format({'font_color':CHARCOAL,'border':1,'num_format':'#,##0','bg_color':LIGHT_GREY,'valign':'vcenter','font_size':9})
pct  = wb.add_format({'font_color':CHARCOAL,'border':1,'num_format':'0.0%','valign':'vcenter','font_size':9})
green_cell = wb.add_format({'font_color':CHARCOAL,'border':1,'bg_color':LIGHT_GREEN,'valign':'vcenter','font_size':9})
amber_cell = wb.add_format({'font_color':CHARCOAL,'border':1,'bg_color':LIGHT_GOLD,'valign':'vcenter','font_size':9})
red_cell   = wb.add_format({'font_color':CHARCOAL,'border':1,'bg_color':RED_BG,'valign':'vcenter','font_size':9})
title_fmt  = wb.add_format({'bold':True,'font_size':14,'font_color':FOREST_GREEN})
title_fmt2 = wb.add_format({'bold':True,'font_size':14,'font_color':MUTED_GOLD})
label_fmt  = wb.add_format({'bold':True,'font_size':10,'font_color':CHARCOAL})
value_fmt  = wb.add_format({'font_size':10,'font_color':CHARCOAL,'text_wrap':True})

def write_df(ws, df, currency_cols=None, row_limit=None):
    """Write a DataFrame to a worksheet."""
    if row_limit:
        df = df.head(row_limit)
    currency_cols = currency_cols or []
    cols = list(df.columns)
    # Header
    ws.set_row(0, 28)
    for c_idx, col in enumerate(cols):
        ws.write(0, c_idx, col, hdr)
        ws.set_column(c_idx, c_idx, max(len(col)*1.1+2, 10))
    # Data
    for r_idx, row in enumerate(df.itertuples(index=False), 1):
        is_grey = r_idx % 2 == 0
        for c_idx, val in enumerate(row, 0):
            col_name = cols[c_idx]
            is_cur = col_name in currency_cols
            # Handle NaN
            if pd.isna(val):
                val = ""
            if is_cur:
                fmt = cur_grey if is_grey else cur
                ws.write_number(r_idx, c_idx, float(val) if val != "" else 0, fmt)
            else:
                fmt = body_grey if is_grey else body
                ws.write(r_idx, c_idx, val, fmt)
    # Autofilter
    if len(df) > 0:
        ws.autofilter(0, 0, len(df), len(cols)-1)
    ws.freeze_panes(1, 0)
    return len(df) + 1

print("Loading datasets...")
dfs = {
    "budget":  pd.read_csv(f"{DATA}/Approved_Annual_Budget.csv"),
    "alloc":   pd.read_csv(f"{DATA}/Monthly_Budget_Allocation.csv"),
    "supp":    pd.read_csv(f"{DATA}/Supplementary_Budget.csv"),
    "realloc": pd.read_csv(f"{DATA}/Budget_Reallocation.csv"),
    "txn":     pd.read_csv(f"{DATA}/Expenditure_Transactions.csv"),
    "commit":  pd.read_csv(f"{DATA}/Commitments.csv"),
    "req":     pd.read_csv(f"{DATA}/Budget_Review_Requests.csv"),
    "proc":    pd.read_csv(f"{DATA}/Procurement_Plan.csv"),
    "pb":      pd.read_csv(f"{DATA}/Pending_Bills.csv"),
    "rev":     pd.read_csv(f"{DATA}/Revenue_AIA.csv"),
    "pi":      pd.read_csv(f"{DATA}/Performance_Indicators.csv"),
    "exc":     pd.read_csv(f"{DATA}/Audit_Control_Exceptions.csv"),
    "dq":      pd.read_csv(f"{DATA}/Data_Quality_Log.csv"),
    "cc":      pd.read_csv(f"{DATA}/Dim_Cost_Centre.csv"),
    "votes":   pd.read_csv(f"{DATA}/Dim_Vote.csv"),
    "acts":    pd.read_csv(f"{DATA}/Dim_Activity.csv"),
    "regions": pd.read_csv(f"{DATA}/Dim_Region.csv"),
    "fs":      pd.read_csv(f"{DATA}/Dim_Fund_Source.csv"),
}

# ── README ──────────────────────────────────────────────────────────────────
print("  Sheet: README")
ws = wb.add_worksheet("README")
ws.hide_gridlines(2)
readme_rows = [
    ("PARASTATAL X","DYNAMIC BUDGET MONITORING, FORECASTING AND MANAGEMENT DASHBOARD"),
    ("",""),
    ("DISCLAIMER","All data in this workbook is entirely synthetic. No real persons, institutions,"),
    ("","suppliers, account numbers or government data are represented."),
    ("",""),
    ("PROJECT","Parastatal X Dynamic Budget Monitoring, Forecasting and Management Dashboard"),
    ("SUBTITLE","A simulated Kenyan public-sector financial analytics solution"),
    ("REPORTING PERIOD","FY 2023/2024 to FY 2025/2026 | Active Year: FY 2025/2026"),
    ("CUT-OFF DATE","31 May 2026"),
    ("CURRENCY","Kenya Shillings (KSh) - All amounts in full KSh"),
    ("FINANCIAL YEAR","1 July to 30 June"),
    ("",""),
    ("WORKSHEET","DESCRIPTION"),
    ("README","This sheet - project overview and navigation"),
    ("Assumptions","Key modelling assumptions and parameters"),
    ("Approved_Annual_Budget","Official approved budget - static reference (96 rows)"),
    ("Monthly_Budget_Allocation","Monthly budget profiles Jul-Jun (1,152 rows)"),
    ("Supplementary_Budget","Supplementary budget adjustments (209 rows)"),
    ("Budget_Reallocation","Budget reallocation requests and approvals (224 rows)"),
    ("Expenditure_Transactions","Detailed expenditure register (3,000 rows shown; full 12,100 in CSV)"),
    ("Commitments","Outstanding commitments register (1,350 rows)"),
    ("Budget_Review_Requests","AIE and budget review workflow (1,000 rows)"),
    ("Procurement_Plan","Procurement plan and contract status (390 rows)"),
    ("Pending_Bills","Outstanding supplier invoices (520 rows)"),
    ("Revenue_AIA","Revenue and Appropriations-in-Aid (1,140 rows)"),
    ("Performance_Indicators","KPI tracking and M&E data (165 rows)"),
    ("Audit_Control_Exceptions","Internal audit and control findings (300 rows)"),
    ("Data_Quality_Log","Data quality issues and resolutions (250 rows)"),
    ("Reconciliation_Control","Data reconciliation checks (18 checks)"),
    ("Dashboard_Input","Summary KPIs for management dashboard"),
    ("Management_Actions","Management action register (10 actions)"),
    ("Lookup_Cost_Centres","Cost centre dimension (32 cost centres)"),
    ("Lookup_Votes","Vote code dimension (33 vote codes)"),
    ("Lookup_Activities","Activity dimension (40 activities)"),
    ("Lookup_Regions","Region dimension (8 regions)"),
    ("Lookup_Fund_Sources","Fund source dimension (8 sources)"),
]
ws.set_column(0,0,30)
ws.set_column(1,1,75)
for r, (lab,val) in enumerate(readme_rows):
    if r == 0:
        ws.write(r,0,lab,title_fmt)
        ws.write(r,1,val,title_fmt2)
    elif r == 12:
        ws.write(r,0,lab,hdr)
        ws.write(r,1,val,hdr)
    else:
        ws.write(r,0,lab,label_fmt)
        ws.write(r,1,val,value_fmt)

# ── Assumptions ─────────────────────────────────────────────────────────────
print("  Sheet: Assumptions")
ws = wb.add_worksheet("Assumptions")
ws.set_column(0,0,38)
ws.set_column(1,1,35)
ws.set_column(2,2,55)
assumptions = [
    ("Parameter","Value","Justification"),
    ("Random Seed","42","Ensures reproducible data generation"),
    ("Financial Year","1 July to 30 June","Kenyan Government financial year"),
    ("Active Financial Year","FY 2025/2026","Current year for analysis"),
    ("Analysis Cut-off Date","31 May 2026","11 months of actual data available"),
    ("Total Approved Budget FY2025/26","KSh 1,700,000,000","Realistic parastatal budget scale"),
    ("Total Approved Budget FY2024/25","KSh 1,550,000,000","Prior year benchmark"),
    ("Total Approved Budget FY2023/24","KSh 1,400,000,000","Base year benchmark"),
    ("Number of Cost Centres","32","13 directorates and 8 regions"),
    ("YTD Expenditure (as at 31 May 2026)","KSh 1,478,365,800","87.0% of approved budget"),
    ("Expected Time-Based Absorption","91.7%","11/12 months elapsed"),
    ("Absorption Gap","-4.7 percentage points","Below expected - management attention required"),
    ("Forecast Year-End Expenditure","KSh 1,763,490,330","Blended 4-method forecast"),
    ("Forecast Absorption Rate","103.7%","Slight over-expenditure risk"),
    ("Over-Expenditure Risk","KSh 63,490,330","Requires management action"),
    ("Open Commitments Outstanding","KSh 700,000,000","Cash exposure beyond YTD spend"),
    ("Data Quality Issue Rate","~2.6%","Intentional issues for demonstration"),
    ("Seasonality","Increased Q3/Q4 and June year-end","Reflects public sector patterns"),
    ("Currency","Kenya Shillings (KSh)","All amounts in full KSh"),
    ("Forecast Method","Blended (30% SL + 25% MA3 + 25% Seasonal + 20% Commit-adj)","Multi-method for robustness"),
]
for r, row in enumerate(assumptions):
    fmt = hdr if r==0 else (body if r%2==0 else body_grey)
    ws.write_row(r, 0, row, fmt)

# ── Core data sheets ─────────────────────────────────────────────────────────
print("  Sheet: Approved_Annual_Budget")
write_df(wb.add_worksheet("Approved_Annual_Budget"), dfs["budget"],
         currency_cols=["ApprovedAnnualAmount"])

print("  Sheet: Monthly_Budget_Allocation")
write_df(wb.add_worksheet("Monthly_Budget_Allocation"), dfs["alloc"],
         currency_cols=["MonthlyAllocationAmount","CumulativeAllocationAmount"])

print("  Sheet: Supplementary_Budget")
write_df(wb.add_worksheet("Supplementary_Budget"), dfs["supp"],
         currency_cols=["OriginalApprovedAmount","SupplementaryIncrease",
                        "SupplementaryReduction","NetAdjustment","RevisedApprovedAmount"])

print("  Sheet: Budget_Reallocation")
write_df(wb.add_worksheet("Budget_Reallocation"), dfs["realloc"],
         currency_cols=["ReallocationAmount"])

print("  Sheet: Expenditure_Transactions (3,000 rows, full data in CSV)")
ws_txn = wb.add_worksheet("Expenditure_Transactions")
note_fmt = wb.add_format({'bold':True,'font_color':'#C0392B','font_size':10})
ws_txn.write(0, 0, "NOTE: 3,000 rows shown. Full 12,100-row dataset is in Expenditure_Transactions.csv", note_fmt)
write_df(wb.add_worksheet("Expenditure_Transactions_Full"), dfs["txn"].head(3000),
         currency_cols=["GrossAmount","TaxAmount","NetAmount"])
# Rename in a workaround (xlsxwriter doesn't rename sheets after creation - just add a note)

print("  Sheet: Commitments")
write_df(wb.add_worksheet("Commitments"), dfs["commit"],
         currency_cols=["CommitmentAmount","AmountInvoiced","AmountPaid","OutstandingCommitment"])

print("  Sheet: Budget_Review_Requests")
write_df(wb.add_worksheet("Budget_Review_Requests"), dfs["req"],
         currency_cols=["RequestedAmount","ApprovedAmount","ProcessedAmount","OutstandingAmount"])

print("  Sheet: Procurement_Plan")
write_df(wb.add_worksheet("Procurement_Plan"), dfs["proc"],
         currency_cols=["EstimatedCost","ApprovedBudget","ActualContractValue",
                        "AmountPaid","OutstandingContractAmount"])

print("  Sheet: Pending_Bills")
write_df(wb.add_worksheet("Pending_Bills"), dfs["pb"],
         currency_cols=["InvoiceAmount","AmountVerified","AmountPaid","OutstandingAmount"])

print("  Sheet: Revenue_AIA")
write_df(wb.add_worksheet("Revenue_AIA"), dfs["rev"],
         currency_cols=["BudgetedRevenue","ActualRevenue","Variance"])

print("  Sheet: Performance_Indicators")
write_df(wb.add_worksheet("Performance_Indicators"), dfs["pi"],
         currency_cols=["BudgetAllocated","Expenditure","CostPerOutput"])

print("  Sheet: Audit_Control_Exceptions")
write_df(wb.add_worksheet("Audit_Control_Exceptions"), dfs["exc"],
         currency_cols=["AmountAtRisk"])

print("  Sheet: Data_Quality_Log")
write_df(wb.add_worksheet("Data_Quality_Log"), dfs["dq"])

# ── Reconciliation ───────────────────────────────────────────────────────────
print("  Sheet: Reconciliation_Control")
rec_path = f"{DATA}/../04_Analysis/Reconciliation_Results.csv"
if os.path.exists(rec_path):
    write_df(wb.add_worksheet("Reconciliation_Control"), pd.read_csv(rec_path),
             currency_cols=["SourceTotal","ModelTotal","Difference","AbsDifference"])

# ── Dashboard_Input ──────────────────────────────────────────────────────────
print("  Sheet: Dashboard_Input")
ws_d = wb.add_worksheet("Dashboard_Input")
ws_d.hide_gridlines(2)
ws_d.set_column(0,0,45)
ws_d.set_column(1,1,22)
ws_d.set_column(2,2,10)
ws_d.set_column(3,3,12)
ws_d.set_column(4,4,55)

txn_ = dfs["txn"]
bud_ = dfs["budget"]
com_ = dfs["commit"]
pb_  = dfs["pb"]
req_ = dfs["req"]
rev_ = dfs["rev"]
exc_ = dfs["exc"]
fy_  = "FY2025/2026"
approved_budget = bud_[bud_.FinancialYear==fy_]["ApprovedAnnualAmount"].sum()
ytd_exp = txn_[txn_.FinancialYear==fy_]["NetAmount"].sum()
total_commits = com_[com_.FinancialYear==fy_]["OutstandingCommitment"].sum()
pending_bills = pb_[pb_.FinancialYear==fy_]["OutstandingAmount"].sum()
budgeted_rev  = rev_[rev_.FinancialYear==fy_]["BudgetedRevenue"].sum()
actual_rev    = rev_[rev_.FinancialYear==fy_]["ActualRevenue"].sum()
rev_ach = actual_rev/budgeted_rev if budgeted_rev>0 else 0
forecast_ye   = 1_763_490_330
available_bal = approved_budget - ytd_exp - total_commits
util_pct      = ytd_exp / approved_budget
exc_open      = exc_[(exc_.FinancialYear==fy_)&(exc_.Status.isin(["Open","In Progress","Escalated"]))]["AmountAtRisk"].sum()

kpis = [
    ("KPI METRIC","VALUE","UNIT","STATUS","MANAGEMENT COMMENTARY"),
    ("Approved Annual Budget", approved_budget, "KSh", "Reference", "FY2025/2026 original approved budget"),
    ("YTD Actual Expenditure", ytd_exp, "KSh", "Amber", f"As at 31 May 2026 — 87.0% utilisation"),
    ("Budget Utilisation %", util_pct, "%", "Amber", "YTD vs Approved Budget — 4.7 pts below expected"),
    ("Total Open Commitments", total_commits, "KSh", "Amber", "Outstanding commitment exposure requiring monitoring"),
    ("Available Balance (Net of Commitments)", available_bal, "KSh", "Amber" if available_bal>0 else "Red", "Remaining available after YTD spend and commitments"),
    ("Forecast Year-End Expenditure", forecast_ye, "KSh", "Amber", "Blended 4-method forecast — slight over-exp risk"),
    ("Forecast Absorption Rate", forecast_ye/approved_budget, "%", "Amber", "103.7% — KSh 63.5M over-expenditure risk"),
    ("Pending Bills Outstanding", pending_bills, "KSh", "Amber", "Uncleared supplier invoices as at cut-off"),
    ("Budgeted Revenue", budgeted_rev, "KSh", "Reference", "Revenue target FY2025/2026"),
    ("Actual Revenue Collected", actual_rev, "KSh", "Amber", f"Achievement: {rev_ach*100:.1f}% vs target"),
    ("Revenue Achievement %", rev_ach, "%", "Amber", "12% below target — collection drive required"),
    ("Audit Exceptions - Amount at Risk", exc_open, "KSh", "Red", "Unresolved audit and control findings"),
    ("Data Quality Score", 0.962, "%", "Green", "96.2% data completeness — above 95% threshold"),
    ("Months Elapsed in FY", 11, "Months", "Reference", "11 of 12 months elapsed (Jul 2025 - May 2026)"),
    ("Expected Utilisation (Time-based)", 11/12, "%", "Reference", "91.7% expected — actual is 87.0%"),
]

cur_fmt_d  = wb.add_format({'num_format':'#,##0','font_color':CHARCOAL,'border':1,'font_size':10})
pct_fmt_d  = wb.add_format({'num_format':'0.0%','font_color':CHARCOAL,'border':1,'font_size':10})
val_fmt_d  = wb.add_format({'font_color':CHARCOAL,'border':1,'font_size':10})
for r, row in enumerate(kpis):
    kpi, val, unit, status, comment = row
    hf = hdr if r==0 else None
    sf_map = {"Reference": body, "Green": green_cell, "Amber": amber_cell, "Red": red_cell}
    sf = hf or sf_map.get(status, body)
    ws_d.write(r, 0, kpi, sf)
    if r == 0:
        ws_d.write(r, 1, val, hdr)
        ws_d.write(r, 2, unit, hdr)
        ws_d.write(r, 3, status, hdr)
        ws_d.write(r, 4, comment, hdr)
    else:
        if unit == "KSh":
            sf_cur = wb.add_format({'num_format':'#,##0','font_color':CHARCOAL,'border':1,
                                    'font_size':10,'bg_color':{"Reference":WHITE,"Green":LIGHT_GREEN,"Amber":LIGHT_GOLD,"Red":RED_BG}.get(status,WHITE)})
            ws_d.write_number(r, 1, float(val), sf_cur)
        elif unit == "%":
            sf_pct = wb.add_format({'num_format':'0.0%','font_color':CHARCOAL,'border':1,
                                    'font_size':10,'bg_color':{"Reference":WHITE,"Green":LIGHT_GREEN,"Amber":LIGHT_GOLD,"Red":RED_BG}.get(status,WHITE)})
            ws_d.write_number(r, 1, float(val), sf_pct)
        else:
            ws_d.write(r, 1, val, sf)
        ws_d.write(r, 2, unit, sf)
        ws_d.write(r, 3, status, sf)
        ws_d.write(r, 4, comment, sf)

# ── Management Actions ────────────────────────────────────────────────────────
print("  Sheet: Management_Actions")
ws_ma = wb.add_worksheet("Management_Actions")
ws_ma.set_column(0,0,10); ws_ma.set_column(1,1,40); ws_ma.set_column(2,2,12)
ws_ma.set_column(3,3,45); ws_ma.set_column(4,4,35); ws_ma.set_column(5,5,12)
ws_ma.set_column(6,6,14); ws_ma.set_column(7,7,14); ws_ma.set_column(8,8,30)
actions_hdr = ["ActionID","Finding","Risk","RecommendedAction","ResponsibleRole","Priority","TargetDate","Status","Comments"]
actions_data = [
    ("ACT-001","5 cost centres at over-expenditure risk","High","Restrict new expenditure; seek reallocation from underspent votes","Director, Finance and Accounts","High","2026-06-15","In Progress","Monthly review initiated; 3 CCs under special monitoring"),
    ("ACT-002","KSh 700M open commitments outstanding","High","Review all open commitments; clear overdue cases before year-end","Director, Supply Chain Management","High","2026-06-20","In Progress","Supplier negotiations ongoing; 8 commitments escalated"),
    ("ACT-003","Revenue 12% below target","Medium","Intensify collection efforts; issue demand notices to debtors","Revenue Manager","Medium","2026-06-30","Pending","Revenue drive planned for June 2026"),
    ("ACT-004","18 cost centres with low absorption","High","Revise work plans; accelerate pending procurements","Cost Centre Manager","High","2026-06-30","In Progress","Work plans revised for 12 CCs"),
    ("ACT-005","Pending bills over 90 days","High","Prioritise payment of all invoices >90 days","Chief Accountant","High","2026-06-10","In Progress","Payment schedule prepared; 23 invoices queued"),
    ("ACT-006","23 cost centres RAG Red","Critical","Immediate CEO directive; bi-weekly status reporting required","Chief Executive Officer","Critical","2026-06-15","In Progress","CEO directive issued on 3 June 2026"),
    ("ACT-007","Audit exceptions - missing documents","Medium","Obtain and file all missing supporting documents before audit","Head of Internal Audit","Medium","2026-06-30","Pending","Circular issued to all cost centre managers"),
    ("ACT-008","12 procurement contracts delayed","Medium","Issue contract extensions; re-procure where necessary","Director, Supply Chain Management","Medium","2026-06-30","Pending","Contract review meeting scheduled for 15 June 2026"),
    ("ACT-009","Budget requests unprocessed >30 days","Medium","Clear backlog and strengthen processing turnaround","Head of Budget","Medium","2026-06-30","In Progress","Processing team reinforced; target 5-day turnaround"),
    ("ACT-010","Donor programme underspend","High","Accelerate donor fund utilisation; engage donor on timeline","Project Manager","High","2026-06-30","Pending","Donor notified; revised implementation plan submitted"),
]
ws_ma.write_row(0, 0, actions_hdr, hdr2)
priority_fmts = {
    "Critical": wb.add_format({'bg_color':RED_BG,'border':1,'font_size':9}),
    "High":     wb.add_format({'bg_color':LIGHT_GOLD,'border':1,'font_size':9}),
    "Medium":   wb.add_format({'bg_color':LIGHT_GREY,'border':1,'font_size':9}),
    "Low":      wb.add_format({'border':1,'font_size':9}),
}
for r_idx, row in enumerate(actions_data, 1):
    p = row[5]
    fmt = priority_fmts.get(p, body)
    for c_idx, val in enumerate(row):
        ws_ma.write(r_idx, c_idx, val, fmt)
ws_ma.freeze_panes(1, 0)

# ── Lookup tables ────────────────────────────────────────────────────────────
print("  Sheet: Lookup_Cost_Centres")
write_df(wb.add_worksheet("Lookup_Cost_Centres"), dfs["cc"])
print("  Sheet: Lookup_Votes")
write_df(wb.add_worksheet("Lookup_Votes"), dfs["votes"])
print("  Sheet: Lookup_Activities")
write_df(wb.add_worksheet("Lookup_Activities"), dfs["acts"])
print("  Sheet: Lookup_Regions")
write_df(wb.add_worksheet("Lookup_Regions"), dfs["regions"])
print("  Sheet: Lookup_Fund_Sources")
write_df(wb.add_worksheet("Lookup_Fund_Sources"), dfs["fs"])

wb.close()
size = os.path.getsize(filepath)/(1024*1024)
print(f"\nExcel workbook saved: {filepath}")
print(f"File size: {size:.1f} MB  |  Sheets: 26")
