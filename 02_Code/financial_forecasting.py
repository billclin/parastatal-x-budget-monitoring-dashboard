"""
Parastatal X - Financial Forecasting Script
Generates year-end expenditure forecasts using multiple methods.
Cut-off date: 31 May 2026 (Month 11 of FY2025/2026)
"""
import pandas as pd
import numpy as np

from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "01_Data"
OUT  = ROOT / "04_Analysis"

txn    = pd.read_csv(f"{DATA}/Expenditure_Transactions.csv")
budget = pd.read_csv(f"{DATA}/Approved_Annual_Budget.csv")
commit = pd.read_csv(f"{DATA}/Commitments.csv")

# ── FY2025/2026 YTD actuals ─────────────────────────────────────────────────
active = txn[txn["FinancialYear"]=="FY2025/2026"].copy()
active["FYMonthNum"] = pd.to_datetime(active["TransactionDate"]).apply(
    lambda d: ((d.month - 7) % 12) + 1)

MONTHS_ELAPSED  = 11   # July 2025 to May 2026
MONTHS_REMAINING = 1   # June 2026

approved_budget = budget[budget["FinancialYear"]=="FY2025/2026"]["ApprovedAnnualAmount"].sum()
ytd_expenditure = active["NetAmount"].sum()
avg_monthly     = ytd_expenditure / MONTHS_ELAPSED

# Monthly breakdown for seasonality
monthly_exp = active.groupby("FYMonthNum")["NetAmount"].sum()
weights_hist = [1.0,0.8,1.2,1.0,1.1,0.9,1.0,1.1,1.2,1.3,1.5,2.0]

# ── 3-Month Moving Average ──────────────────────────────────────────────────
last3 = [monthly_exp.get(i, avg_monthly) for i in [9,10,11]]
ma3   = np.mean(last3)

# ── Method 1: Straight-line run rate ───────────────────────────────────────
forecast_sl = ytd_expenditure + (avg_monthly * MONTHS_REMAINING)

# ── Method 2: Weighted 3-month moving average ──────────────────────────────
forecast_ma = ytd_expenditure + (ma3 * MONTHS_REMAINING)

# ── Method 3: Seasonal forecast ────────────────────────────────────────────
seasonal_weight_remaining = weights_hist[11] / sum(weights_hist[:11]) * ytd_expenditure
forecast_seasonal = ytd_expenditure + seasonal_weight_remaining

# ── Method 4: Commitment-adjusted forecast ─────────────────────────────────
open_commits = commit[(commit["FinancialYear"]=="FY2025/2026") &
                      (commit["CommitmentStatus"].isin(["Open","Partially Invoiced","Partially Paid","Overdue"]))]["OutstandingCommitment"].sum()
forecast_commit_adj = ytd_expenditure + open_commits

# ── Selected (blended) forecast ────────────────────────────────────────────
forecast_selected = round((forecast_sl * 0.30 + forecast_ma * 0.25 +
                           forecast_seasonal * 0.25 + forecast_commit_adj * 0.20))

balance = approved_budget - forecast_selected
over_exp_risk = max(0, forecast_selected - approved_budget)
utilisation_pct = ytd_expenditure / approved_budget * 100
absorption_pct  = forecast_selected / approved_budget * 100
expected_util   = MONTHS_ELAPSED / 12 * 100

print("=" * 65)
print("PARASTATAL X - FINANCIAL FORECAST SUMMARY")
print("Cut-off: 31 May 2026 | Active FY: FY2025/2026")
print("=" * 65)
print(f"\nApproved Annual Budget          : KSh {approved_budget:>20,.0f}")
print(f"YTD Actual Expenditure          : KSh {ytd_expenditure:>20,.0f}")
print(f"Months Elapsed                  : {MONTHS_ELAPSED} of 12")
print(f"Average Monthly Expenditure     : KSh {avg_monthly:>20,.0f}")
print(f"\nYTD Utilisation                 : {utilisation_pct:.1f}%")
print(f"Expected Utilisation (time-based): {expected_util:.1f}%")
print(f"Absorption Gap                  : {utilisation_pct - expected_util:+.1f}%")
print(f"\n--- FORECAST METHODS ---")
print(f"Method 1 - Straight-line Run Rate    : KSh {forecast_sl:>15,.0f}")
print(f"Method 2 - 3-Month Moving Average    : KSh {forecast_ma:>15,.0f}")
print(f"Method 3 - Seasonal Forecast         : KSh {forecast_seasonal:>15,.0f}")
print(f"Method 4 - Commitment-Adjusted       : KSh {forecast_commit_adj:>15,.0f}")
print(f"\nSELECTED BLENDED FORECAST            : KSh {forecast_selected:>15,.0f}")
print(f"Forecast Absorption Rate             : {absorption_pct:.1f}%")
print(f"Forecast Closing Balance             : KSh {balance:>15,.0f}")
print(f"Over-Expenditure Risk                : KSh {over_exp_risk:>15,.0f}")
print(f"Open Commitments Outstanding         : KSh {open_commits:>15,.0f}")

# ── Cost-centre level forecast ─────────────────────────────────────────────
cc_ytd  = active.groupby("CostCentreCode")["NetAmount"].sum().reset_index()
cc_ytd.columns = ["CostCentreCode","YTDExpenditure"]
cc_budget_tbl = budget[budget["FinancialYear"]=="FY2025/2026"].groupby("CostCentreCode")["ApprovedAnnualAmount"].sum().reset_index()
cc_fc = cc_ytd.merge(cc_budget_tbl, on="CostCentreCode", how="left")
cc_fc["ForecastYE"] = cc_fc["YTDExpenditure"] * (12/11)
cc_fc["ForecastBalance"] = cc_fc["ApprovedAnnualAmount"] - cc_fc["ForecastYE"]
cc_fc["ForecastAbsorptionPct"] = cc_fc["ForecastYE"] / cc_fc["ApprovedAnnualAmount"] * 100
cc_fc["OverExpRisk"] = cc_fc["ForecastBalance"].apply(lambda x: "Yes" if x < 0 else "No")
cc_fc["RAGStatus"] = cc_fc["ForecastAbsorptionPct"].apply(
    lambda p: "Green" if abs(p-100) <= 10 else ("Amber" if abs(p-100) <= 20 else "Red"))

cc_fc.to_csv(f"{OUT}/CC_Level_Forecast.csv", index=False)

print(f"\n--- COST CENTRE RISK SUMMARY ---")
print(f"Over-expenditure risk CCs  : {(cc_fc['OverExpRisk']=='Yes').sum()}")
print(f"Green status CCs           : {(cc_fc['RAGStatus']=='Green').sum()}")
print(f"Amber status CCs           : {(cc_fc['RAGStatus']=='Amber').sum()}")
print(f"Red status CCs             : {(cc_fc['RAGStatus']=='Red').sum()}")

# ── Scenario Analysis ──────────────────────────────────────────────────────
print(f"\n--- SCENARIO ANALYSIS ---")
scenarios = {
    "Base Case":     forecast_selected,
    "Optimistic":    forecast_selected * 0.97,
    "Adverse":       forecast_selected * 1.08,
}
for s, f in scenarios.items():
    b = approved_budget - f
    print(f"{s:<15}: Forecast KSh {f:>15,.0f}  |  Balance KSh {b:>15,.0f}  |  {'OVER-EXP RISK' if b<0 else 'Within Budget'}")

# Save forecast summary
summary = {
    "ApprovedBudget": approved_budget,
    "YTDExpenditure": ytd_expenditure,
    "MonthsElapsed": MONTHS_ELAPSED,
    "AvgMonthlyExp": avg_monthly,
    "YTDUtilisationPct": round(utilisation_pct,1),
    "ExpectedUtilPct": round(expected_util,1),
    "Forecast_StraightLine": round(forecast_sl),
    "Forecast_MovingAvg": round(forecast_ma),
    "Forecast_Seasonal": round(forecast_seasonal),
    "Forecast_CommitAdj": round(forecast_commit_adj),
    "Forecast_Blended": forecast_selected,
    "ForecastAbsorptionPct": round(absorption_pct,1),
    "ForecastClosingBalance": round(balance),
    "OpenCommitmentsOutstanding": round(open_commits),
    "OverExpRisk": round(over_exp_risk),
    "BaseCase": round(scenarios["Base Case"]),
    "Optimistic": round(scenarios["Optimistic"]),
    "Adverse": round(scenarios["Adverse"]),
}
pd.DataFrame([summary]).to_csv(f"{OUT}/Forecast_Summary.csv", index=False)
print("\nForecast outputs saved.")
