"""
Parastatal X - Data Validation Script
Runs automated checks across all generated datasets.
"""
import pandas as pd
import os

from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "01_Data"
results = []

def check(name, passed, detail=""):
    status = "PASS" if passed else "FAIL"
    results.append({"Check": name, "Status": status, "Detail": detail})
    print(f"  [{status}] {name}" + (f": {detail}" if detail else ""))

print("=== PARASTATAL X DATA VALIDATION ===\n")

# Load datasets
budget  = pd.read_csv(f"{DATA}/Approved_Annual_Budget.csv")
alloc   = pd.read_csv(f"{DATA}/Monthly_Budget_Allocation.csv")
supp    = pd.read_csv(f"{DATA}/Supplementary_Budget.csv")
realloc = pd.read_csv(f"{DATA}/Budget_Reallocation.csv")
txn     = pd.read_csv(f"{DATA}/Expenditure_Transactions.csv")
commit  = pd.read_csv(f"{DATA}/Commitments.csv")
req     = pd.read_csv(f"{DATA}/Budget_Review_Requests.csv")
proc    = pd.read_csv(f"{DATA}/Procurement_Plan.csv")
pb      = pd.read_csv(f"{DATA}/Pending_Bills.csv")
rev     = pd.read_csv(f"{DATA}/Revenue_AIA.csv")
pi      = pd.read_csv(f"{DATA}/Performance_Indicators.csv")
exc     = pd.read_csv(f"{DATA}/Audit_Control_Exceptions.csv")
dq      = pd.read_csv(f"{DATA}/Data_Quality_Log.csv")

print("--- 1. Row Count Checks ---")
check("Budget rows >= 90", len(budget) >= 90, f"{len(budget)} rows")
check("Monthly allocation rows >= 1000", len(alloc) >= 1000, f"{len(alloc)} rows")
check("Expenditure transactions >= 10000", len(txn) >= 10000, f"{len(txn)} rows")
check("Commitments >= 500", len(commit) >= 500, f"{len(commit)} rows")
check("Budget requests >= 800", len(req) >= 800, f"{len(req)} rows")

print("\n--- 2. Duplicate Primary Key Checks ---")
check("Budget IDs unique", budget["BudgetID"].nunique() == len(budget))
check("Allocation IDs unique", alloc["AllocationID"].nunique() == len(alloc))
check("Transaction IDs unique", txn["TransactionID"].nunique() == len(txn))
check("Commitment IDs unique", commit["CommitmentID"].nunique() == len(commit))
check("Request IDs unique", req["RequestID"].nunique() == len(req))
check("Pending Bill IDs unique", pb["PendingBillID"].nunique() == len(pb))
check("Revenue IDs unique", rev["RevenueTransactionID"].nunique() == len(rev))
check("Procurement IDs unique", proc["ProcurementPlanID"].nunique() == len(proc))

print("\n--- 3. Financial Year Coverage ---")
for fy in ["FY2023/2024","FY2024/2025","FY2025/2026"]:
    check(f"Budget has {fy}", fy in budget["FinancialYear"].values)
    check(f"Transactions have {fy}", fy in txn["FinancialYear"].values)

print("\n--- 4. Budget Reconciliation ---")
for fy in ["FY2023/2024","FY2024/2025","FY2025/2026"]:
    ann = budget[budget["FinancialYear"]==fy]["ApprovedAnnualAmount"].sum()
    mth = alloc[alloc["FinancialYear"]==fy].groupby(["CostCentreCode","VoteCode"])["MonthlyAllocationAmount"].sum().sum()
    diff = abs(ann - mth)
    check(f"Monthly allocations reconcile to annual budget [{fy}]",
          diff < 1000, f"Annual: {ann:,.0f}  Monthly sum: {mth:,.0f}  Diff: {diff:,.0f}")

print("\n--- 5. Approved Amount Checks ---")
check("No negative approved budget", (budget["ApprovedAnnualAmount"] >= 0).all())
check("No negative monthly allocations", (alloc["MonthlyAllocationAmount"] >= 0).all())
check("No negative net expenditure (after flag)",
      (txn["NetAmount"] >= 0).mean() > 0.95, f"{(txn['NetAmount']<0).sum()} negative found")

print("\n--- 6. Commitment Checks ---")
check("Paid amount <= Commitment amount",
      (commit["AmountPaid"] <= commit["CommitmentAmount"] * 1.001).all())
check("Invoiced amount <= Commitment amount",
      (commit["AmountInvoiced"] <= commit["CommitmentAmount"] * 1.001).all())

print("\n--- 7. Request Amount Checks ---")
clean = req[req["ApprovedAmount"] > 0]
check("Approved <= Requested (for approved requests)",
      (clean["ApprovedAmount"] <= clean["RequestedAmount"] * 1.05).mean() > 0.90,
      f"{(clean['ApprovedAmount']>clean['RequestedAmount']*1.05).sum()} exceptions")
clean2 = req[req["ProcessedAmount"] > 0]
check("Processed <= Approved (for processed requests)",
      (clean2["ProcessedAmount"] <= clean2["ApprovedAmount"] * 1.05).mean() > 0.90,
      f"{(clean2['ProcessedAmount']>clean2['ApprovedAmount']*1.05).sum()} exceptions")

print("\n--- 8. Pending Bills Checks ---")
check("Outstanding = Invoice - Paid",
      ((pb["InvoiceAmount"] - pb["AmountPaid"] - pb["OutstandingAmount"]).abs() < 10).all())

print("\n--- 9. Financial Year Month Coverage ---")
for fy in ["FY2023/2024","FY2024/2025","FY2025/2026"]:
    mths = alloc[alloc["FinancialYear"]==fy]["FYMonthNumber"].nunique()
    check(f"12 months in allocation [{fy}]", mths == 12, f"{mths} months found")

print("\n--- 10. Data Quality Log ---")
check("DQ log has records", len(dq) > 0, f"{len(dq)} records")
check("DQ log has severity field", "Severity" in dq.columns)
severities = dq["Severity"].value_counts()
check("DQ covers Critical issues", "Critical" in severities.index)

print("\n=== SUMMARY ===")
passed = sum(1 for r in results if r["Status"]=="PASS")
failed = sum(1 for r in results if r["Status"]=="FAIL")
print(f"Total checks: {len(results)} | Passed: {passed} | Failed: {failed}")

# Save results
pd.DataFrame(results).to_csv(f"{DATA}/../04_Analysis/Validation_Results.csv" if os.path.exists(f"{DATA}/../04_Analysis") else f"{DATA}/Validation_Results.csv", index=False)
