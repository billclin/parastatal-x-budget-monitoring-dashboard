"""
Parastatal X Dynamic Budget Monitoring, Forecasting and Management Dashboard
Synthetic Data Generation Script
Author: Financial Analytics Portfolio Project
Date: 2026-06-20
Random Seed: 42
Disclaimer: All data is entirely synthetic. No real persons, institutions,
suppliers, account numbers or government data are represented.
"""

import random
import numpy as np
import pandas as pd
from datetime import date, timedelta, datetime
import os

# ── Seed ──────────────────────────────────────────────────────────────────────
random.seed(42)
np.random.seed(42)

from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "01_Data"
os.makedirs(OUT, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════════════
# DIMENSION DEFINITIONS
# ══════════════════════════════════════════════════════════════════════════════

FINANCIAL_YEARS = ["FY2023/2024", "FY2024/2025", "FY2025/2026"]
FY_START = {"FY2023/2024": date(2023, 7, 1),
             "FY2024/2025": date(2024, 7, 1),
             "FY2025/2026": date(2025, 7, 1)}
FY_END   = {"FY2023/2024": date(2024, 6, 30),
             "FY2024/2025": date(2025, 6, 30),
             "FY2025/2026": date(2026, 6, 30)}
CUTOFF   = date(2026, 5, 31)

REGIONS = [
    ("RG-01","Nairobi Region"),
    ("RG-02","Central Region"),
    ("RG-03","Coast Region"),
    ("RG-04","Eastern Region"),
    ("RG-05","North Eastern Region"),
    ("RG-06","Nyanza Region"),
    ("RG-07","Rift Valley Region"),
    ("RG-08","Western Region"),
]

DIRECTORATES = [
    ("DIR-01","Office of the Chief Executive Officer"),
    ("DIR-02","Finance and Accounts Directorate"),
    ("DIR-03","Strategy, Planning and Performance Directorate"),
    ("DIR-04","Human Resource and Administration Directorate"),
    ("DIR-05","Procurement and Supply Chain Directorate"),
    ("DIR-06","Information and Communication Technology Directorate"),
    ("DIR-07","Legal and Corporation Secretary Directorate"),
    ("DIR-08","Internal Audit and Risk Directorate"),
    ("DIR-09","Corporate Communications and Customer Experience Directorate"),
    ("DIR-10","Research, Monitoring and Evaluation Directorate"),
    ("DIR-11","Technical Operations Directorate"),
    ("DIR-12","Projects and Infrastructure Directorate"),
    ("DIR-13","Regional Operations Directorate"),
]

PROGRAMMES = [
    ("PRG-01","General Administration and Support Services"),
    ("PRG-02","Finance and Resource Management"),
    ("PRG-03","Strategic Planning and Performance Management"),
    ("PRG-04","Human Capital Development"),
    ("PRG-05","Procurement and Asset Management"),
    ("PRG-06","ICT Infrastructure and Digital Services"),
    ("PRG-07","Legal Affairs and Compliance"),
    ("PRG-08","Internal Audit and Risk Management"),
    ("PRG-09","Communication and Stakeholder Engagement"),
    ("PRG-10","Research and Knowledge Management"),
    ("PRG-11","Technical Services Delivery"),
    ("PRG-12","Infrastructure Development and Projects"),
    ("PRG-13","Regional Service Delivery"),
    ("PRG-14","Donor-Funded Programmes"),
    ("PRG-15","Special Initiatives and Pilot Projects"),
]

FUND_SOURCES = [
    ("FS-01","Government of Kenya - Recurrent"),
    ("FS-02","Government of Kenya - Development"),
    ("FS-03","Appropriations-in-Aid"),
    ("FS-04","World Bank Grant"),
    ("FS-05","USAID Programme Fund"),
    ("FS-06","EU Sector Support"),
    ("FS-07","AfDB Development Loan"),
    ("FS-08","UNDP Trust Fund"),
]

DONORS = [
    ("DON-00","No Donor / GOK"),
    ("DON-01","World Bank"),
    ("DON-02","USAID"),
    ("DON-03","European Union"),
    ("DON-04","African Development Bank"),
    ("DON-05","UNDP"),
    ("DON-06","DFID / FCDO"),
    ("DON-07","GIZ"),
]

# Cost centres: 50 cost centres spread across directorates and regions
def make_cost_centres():
    rows = []
    cc_id = 1
    dir_map = {d[0]: d[1] for d in DIRECTORATES}
    reg_map = {r[0]: r[1] for r in REGIONS}
    # Headquarters cost centres (one per directorate)
    for d_code, d_name in DIRECTORATES:
        cc_code = f"CC-{cc_id:03d}"
        rows.append({"CostCentreCode": cc_code,
                     "CostCentreName": f"{d_name} - HQ",
                     "DirectorateCode": d_code,
                     "DirectorateName": d_name,
                     "RegionCode": "RG-01",
                     "RegionName": "Nairobi Region",
                     "Type": "Headquarters"})
        cc_id += 1
    # Regional cost centres (one per region for Regional Operations, plus 3 extra)
    for r_code, r_name in REGIONS:
        cc_code = f"CC-{cc_id:03d}"
        rows.append({"CostCentreCode": cc_code,
                     "CostCentreName": f"Regional Office - {r_name}",
                     "DirectorateCode": "DIR-13",
                     "DirectorateName": "Regional Operations Directorate",
                     "RegionCode": r_code,
                     "RegionName": r_name,
                     "Type": "Regional"})
        cc_id += 1
    # Project cost centres
    for i in range(1, 8):
        cc_code = f"CC-{cc_id:03d}"
        rows.append({"CostCentreCode": cc_code,
                     "CostCentreName": f"Infrastructure Project Unit {i:02d}",
                     "DirectorateCode": "DIR-12",
                     "DirectorateName": "Projects and Infrastructure Directorate",
                     "RegionCode": random.choice(REGIONS)[0],
                     "RegionName": "",
                     "Type": "Project"})
        cc_id += 1
    # Fill region names for project CC
    for r in rows:
        if r["RegionName"] == "":
            r["RegionName"] = dict(REGIONS)[r["RegionCode"]]
    # Technical cost centres
    for i in range(1, 5):
        cc_code = f"CC-{cc_id:03d}"
        rows.append({"CostCentreCode": cc_code,
                     "CostCentreName": f"Technical Service Unit {i:02d}",
                     "DirectorateCode": "DIR-11",
                     "DirectorateName": "Technical Operations Directorate",
                     "RegionCode": random.choice(REGIONS)[0],
                     "RegionName": "",
                     "Type": "Technical"})
        cc_id += 1
    for r in rows:
        if r["RegionName"] == "":
            r["RegionName"] = dict(REGIONS)[r["RegionCode"]]
    return rows

COST_CENTRES = make_cost_centres()
CC_CODES = [c["CostCentreCode"] for c in COST_CENTRES]
CC_MAP = {c["CostCentreCode"]: c for c in COST_CENTRES}

# Activities and sub-activities
ACTIVITIES = [
    ("ACT-001","General Management and Administration","PRG-01"),
    ("ACT-002","Board and Governance Support","PRG-01"),
    ("ACT-003","Financial Accounting and Reporting","PRG-02"),
    ("ACT-004","Budget Preparation and Management","PRG-02"),
    ("ACT-005","Internal Controls and Compliance","PRG-02"),
    ("ACT-006","Strategic Planning","PRG-03"),
    ("ACT-007","Performance Monitoring and Reporting","PRG-03"),
    ("ACT-008","Staff Recruitment and Selection","PRG-04"),
    ("ACT-009","Training and Capacity Building","PRG-04"),
    ("ACT-010","Staff Welfare and Benefits","PRG-04"),
    ("ACT-011","Procurement Planning","PRG-05"),
    ("ACT-012","Tender Management","PRG-05"),
    ("ACT-013","Contract Management","PRG-05"),
    ("ACT-014","ICT Systems Development","PRG-06"),
    ("ACT-015","ICT Infrastructure Maintenance","PRG-06"),
    ("ACT-016","Cybersecurity and Data Protection","PRG-06"),
    ("ACT-017","Legal Advisory Services","PRG-07"),
    ("ACT-018","Contract Review and Litigation","PRG-07"),
    ("ACT-019","Internal Audit Programmes","PRG-08"),
    ("ACT-020","Risk Assessment and Management","PRG-08"),
    ("ACT-021","Public Communications","PRG-09"),
    ("ACT-022","Stakeholder Engagement","PRG-09"),
    ("ACT-023","Research and Studies","PRG-10"),
    ("ACT-024","Monitoring and Evaluation","PRG-10"),
    ("ACT-025","Technical Service Delivery","PRG-11"),
    ("ACT-026","Equipment Maintenance and Repair","PRG-11"),
    ("ACT-027","Infrastructure Construction","PRG-12"),
    ("ACT-028","Project Management","PRG-12"),
    ("ACT-029","Regional Service Coordination","PRG-13"),
    ("ACT-030","Community Outreach","PRG-13"),
    ("ACT-031","Donor Programme Implementation","PRG-14"),
    ("ACT-032","Donor Reporting and Compliance","PRG-14"),
    ("ACT-033","Pilot Programme Operations","PRG-15"),
    ("ACT-034","Innovation and Technology Pilots","PRG-15"),
    ("ACT-035","Revenue Collection and Management","PRG-02"),
    ("ACT-036","Asset Management","PRG-05"),
    ("ACT-037","Fleet Management","PRG-04"),
    ("ACT-038","Facility Management","PRG-01"),
    ("ACT-039","Security Services","PRG-01"),
    ("ACT-040","Environmental and Social Safeguards","PRG-12"),
]

SUB_ACTIVITIES = []
for a_code, a_name, prg in ACTIVITIES:
    for i in range(1, random.randint(3, 5)):
        sub_code = f"{a_code}-S{i:02d}"
        sub_name = f"{a_name} - Sub-Activity {i}"
        SUB_ACTIVITIES.append((sub_code, sub_name, a_code))

# Vote codes
VOTES = [
    ("V-2110","Basic Salaries - Permanent","Recurrent","Personnel Emoluments"),
    ("V-2120","Basic Salaries - Temporary","Recurrent","Personnel Emoluments"),
    ("V-2210","Travel and Accommodation","Recurrent","Operations and Maintenance"),
    ("V-2211","Domestic Travel","Recurrent","Operations and Maintenance"),
    ("V-2212","International Travel","Recurrent","Operations and Maintenance"),
    ("V-2220","Communication, Supplies and Services","Recurrent","Operations and Maintenance"),
    ("V-2230","Domestic Utilities","Recurrent","Operations and Maintenance"),
    ("V-2240","Rentals of Produced Assets","Recurrent","Operations and Maintenance"),
    ("V-2250","Repairs and Maintenance","Recurrent","Operations and Maintenance"),
    ("V-2310","Training Expenses","Recurrent","Operations and Maintenance"),
    ("V-2330","Advertising and Public Relations","Recurrent","Operations and Maintenance"),
    ("V-2340","Research, Consultancy and Surveys","Recurrent","Operations and Maintenance"),
    ("V-2360","Fuel and Lubricants","Recurrent","Operations and Maintenance"),
    ("V-2380","Insurance","Recurrent","Operations and Maintenance"),
    ("V-2390","Other Operating Expenses","Recurrent","Operations and Maintenance"),
    ("V-2510","Purchase of Office Furniture","Development","Acquisition of Assets"),
    ("V-2520","Purchase of ICT Equipment","Development","Acquisition of Assets"),
    ("V-2530","Purchase of Vehicles","Development","Acquisition of Assets"),
    ("V-2540","Purchase of Specialised Equipment","Development","Acquisition of Assets"),
    ("V-3110","Construction of Buildings","Development","Capital Expenditure"),
    ("V-3120","Rehabilitation of Buildings","Development","Capital Expenditure"),
    ("V-3130","Road and Civil Works","Development","Capital Expenditure"),
    ("V-3140","Borehole and Water Works","Development","Capital Expenditure"),
    ("V-2410","Allowances","Recurrent","Personnel Emoluments"),
    ("V-2420","Leave Passage","Recurrent","Personnel Emoluments"),
    ("V-2430","Medical Insurance","Recurrent","Personnel Emoluments"),
    ("V-2440","NSSF Employer Contributions","Recurrent","Personnel Emoluments"),
    ("V-2450","Gratuity and Leave Commutation","Recurrent","Personnel Emoluments"),
    ("V-2460","Pension Contributions","Recurrent","Personnel Emoluments"),
    ("V-4110","Donor Technical Assistance","Donor","Donor Expenditure"),
    ("V-4120","Donor Capital Grants","Donor","Donor Expenditure"),
    ("V-4130","Donor Operational Grants","Donor","Donor Expenditure"),
    ("V-5110","AIA - Service Fee Revenue","AIA","Appropriations-in-Aid"),
    ("V-5120","AIA - Licensing Revenue","AIA","Appropriations-in-Aid"),
    ("V-5130","AIA - Training Revenue","AIA","Appropriations-in-Aid"),
]

VOTE_CODES = [v[0] for v in VOTES]
VOTE_MAP = {v[0]: v for v in VOTES}

PROCUREMENT_CATEGORIES = [
    "Professional Services","ICT Services","Fuel and Lubricants",
    "Maintenance","Training","Travel and Accommodation","Utilities",
    "Security Services","Printing and Stationery","Construction",
    "Equipment","Consultancy","Medical Supplies","Office Furniture",
    "Cleaning Services"
]

SUPPLIER_CATEGORIES = [
    "Professional Services","ICT Services","Fuel and Lubricants",
    "Maintenance","Training","Travel and Accommodation","Utilities",
    "Security Services","Printing and Stationery","Construction",
    "Equipment","Consultancy"
]

OFFICER_ROLES = [
    "Chief Executive Officer","Director, Finance and Accounts",
    "Head of Budget","Chief Accountant","Accounts Officer",
    "Director, Supply Chain Management","Procurement Officer",
    "Director, Strategy and Planning","Director, Human Resource and Administration",
    "Head of Internal Audit","Internal Auditor","Regional Manager",
    "Cost Centre Manager","Project Manager","Revenue Manager",
    "ICT Manager","Legal Counsel","M&E Officer","Data Officer",
    "Finance Officer","Senior Finance Officer","Budget Analyst",
    "Director, ICT","Director, Legal","Director, Communications",
    "Director, Research","Director, Technical Operations",
    "Director, Projects and Infrastructure","Director, Regional Operations",
]

RISK_RATINGS = ["Critical","High","Medium","Low"]

def rand_date(start, end):
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, max(0, delta)))

def fy_months(fy):
    """Return list of (month_num, month_name, calendar_year, date) tuples for a FY."""
    months = []
    start = FY_START[fy]
    for i in range(12):
        m = (start.month - 1 + i) % 12 + 1
        y = start.year + (start.month - 1 + i) // 12
        months.append((i+1, date(y, m, 1).strftime("%B"), y, date(y, m, 1)))
    return months

def seasonality_weights():
    """Jul=1, Aug=0.8, Sep=1.2, Oct=1.0, Nov=1.1, Dec=0.9,
       Jan=1.0, Feb=1.1, Mar=1.2, Apr=1.3, May=1.5, Jun=2.0"""
    return [1.0, 0.8, 1.2, 1.0, 1.1, 0.9, 1.0, 1.1, 1.2, 1.3, 1.5, 2.0]

def weighted_monthly_split(annual, fy):
    """Split annual budget into months with seasonal weights."""
    w = seasonality_weights()
    total_w = sum(w)
    raw = [annual * wi / total_w for wi in w]
    # Round to integers
    splits = [round(r) for r in raw]
    # Fix rounding diff
    diff = annual - sum(splits)
    splits[-1] += diff
    return splits

def supplier_id():
    return f"SUP-{random.randint(1000,9999)}"

def doc_ref(prefix):
    return f"{prefix}-{random.randint(100000,999999)}"

def get_prg_for_dir(dir_code):
    mapping = {
        "DIR-01": "PRG-01","DIR-02": "PRG-02","DIR-03": "PRG-03",
        "DIR-04": "PRG-04","DIR-05": "PRG-05","DIR-06": "PRG-06",
        "DIR-07": "PRG-07","DIR-08": "PRG-08","DIR-09": "PRG-09",
        "DIR-10": "PRG-10","DIR-11": "PRG-11","DIR-12": "PRG-12",
        "DIR-13": "PRG-13",
    }
    return mapping.get(dir_code, "PRG-01")

def get_act_for_prg(prg_code):
    acts = [a[0] for a in ACTIVITIES if a[2] == prg_code]
    return random.choice(acts) if acts else random.choice(ACTIVITIES)[0]

def get_sub_for_act(act_code):
    subs = [s[0] for s in SUB_ACTIVITIES if s[2] == act_code]
    return random.choice(subs) if subs else random.choice(SUB_ACTIVITIES)[0]

def get_vote_for_cat(econ_class=None):
    if econ_class == "Development":
        v = [v[0] for v in VOTES if v[2] == "Development"]
    elif econ_class == "Donor":
        v = [v[0] for v in VOTES if v[2] == "Donor"]
    elif econ_class == "AIA":
        v = [v[0] for v in VOTES if v[2] == "AIA"]
    else:
        v = [v[0] for v in VOTES if v[2] == "Recurrent"]
    return random.choice(v)

PRG_MAP = {p[0]: p[1] for p in PROGRAMMES}
DIR_MAP = {d[0]: d[1] for d in DIRECTORATES}
REG_MAP = {r[0]: r[1] for r in REGIONS}
ACT_MAP = {a[0]: a[1] for a in ACTIVITIES}
FS_MAP  = {f[0]: f[1] for f in FUND_SOURCES}
DON_MAP = {d[0]: d[1] for d in DONORS}

# Budget scale: Total approved budget ~KSh 4.8 billion across all FYs
# FY2023/2024: 1.4B, FY2024/2025: 1.55B, FY2025/2026: 1.7B

def cc_budget(fy):
    """Assign budget amounts to each cost centre for a FY."""
    budgets = {}
    fy_totals = {"FY2023/2024": 1_400_000_000,
                 "FY2024/2025": 1_550_000_000,
                 "FY2025/2026": 1_700_000_000}
    total = fy_totals[fy]
    # Weight each cost centre
    weights = []
    for cc in COST_CENTRES:
        if cc["Type"] == "Headquarters":
            if cc["DirectorateCode"] in ["DIR-12","DIR-11"]:
                weights.append(8.0)
            elif cc["DirectorateCode"] in ["DIR-02","DIR-04"]:
                weights.append(5.0)
            else:
                weights.append(2.5)
        elif cc["Type"] == "Regional":
            weights.append(3.5)
        elif cc["Type"] == "Project":
            weights.append(7.0)
        else:
            weights.append(4.0)
    total_w = sum(weights)
    for i, cc in enumerate(COST_CENTRES):
        budgets[cc["CostCentreCode"]] = round(total * weights[i] / total_w)
    # Fix rounding
    diff = total - sum(budgets.values())
    budgets[COST_CENTRES[-1]["CostCentreCode"]] += diff
    return budgets

print("Building dimension tables...")

# ══════════════════════════════════════════════════════════════════════════════
# 6.1 APPROVED ANNUAL BUDGET
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Approved_Annual_Budget...")
budget_rows = []
budget_id = 1
for fy in FINANCIAL_YEARS:
    budgets = cc_budget(fy)
    approval_date = FY_START[fy] - timedelta(days=random.randint(15, 45))
    for cc in COST_CENTRES:
        cc_code = cc["CostCentreCode"]
        dir_code = cc["DirectorateCode"]
        prg_code = get_prg_for_dir(dir_code)
        act_code = get_act_for_prg(prg_code)
        sub_code = get_sub_for_act(act_code)
        # 70% recurrent, 25% dev, 5% AIA
        r = random.random()
        if r < 0.70:
            fs_code = "FS-01"
            don_code = "DON-00"
            vote_code = get_vote_for_cat("Recurrent")
            econ = "Recurrent"
        elif r < 0.90:
            fs_code = "FS-02"
            don_code = "DON-00"
            vote_code = get_vote_for_cat("Development")
            econ = "Development"
        elif r < 0.95:
            fs_code = random.choice(["FS-04","FS-05","FS-06","FS-07","FS-08"])
            don_code = random.choice(["DON-01","DON-02","DON-03","DON-04","DON-05"])
            vote_code = get_vote_for_cat("Donor")
            econ = "Donor"
        else:
            fs_code = "FS-03"
            don_code = "DON-00"
            vote_code = get_vote_for_cat("AIA")
            econ = "AIA"

        amt = budgets[cc_code]
        budget_rows.append({
            "BudgetID": f"ABB-{budget_id:05d}",
            "FinancialYear": fy,
            "BudgetVersion": "Original",
            "ApprovalDate": approval_date.strftime("%Y-%m-%d"),
            "ApprovalAuthority": "Board of Directors",
            "ProgrammeCode": prg_code,
            "ProgrammeName": PRG_MAP[prg_code],
            "DirectorateCode": dir_code,
            "DirectorateName": DIR_MAP[dir_code],
            "CostCentreCode": cc_code,
            "CostCentreName": cc["CostCentreName"],
            "RegionCode": cc["RegionCode"],
            "RegionName": cc["RegionName"],
            "FundSourceCode": fs_code,
            "FundSourceName": FS_MAP[fs_code],
            "DonorCode": don_code,
            "DonorName": DON_MAP[don_code],
            "ActivityCode": act_code,
            "ActivityName": ACT_MAP[act_code],
            "SubActivityCode": sub_code,
            "SubActivityName": [s[1] for s in SUB_ACTIVITIES if s[0]==sub_code][0] if [s[1] for s in SUB_ACTIVITIES if s[0]==sub_code] else sub_code,
            "VoteCode": vote_code,
            "VoteDescription": VOTE_MAP[vote_code][1],
            "EconomicClassification": econ,
            "ApprovedAnnualAmount": amt,
            "SourceDocument": doc_ref("BD"),
            "Status": "Approved",
            "Notes": "",
        })
        budget_id += 1

df_budget = pd.DataFrame(budget_rows)
df_budget.to_csv(f"{OUT}/Approved_Annual_Budget.csv", index=False)
print(f"    Approved_Annual_Budget: {len(df_budget)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.2 MONTHLY BUDGET ALLOCATION
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Monthly_Budget_Allocation...")
alloc_rows = []
alloc_id = 1
for fy in FINANCIAL_YEARS:
    fy_months_list = fy_months(fy)
    for _, row in df_budget[df_budget["FinancialYear"]==fy].iterrows():
        annual = row["ApprovedAnnualAmount"]
        splits = weighted_monthly_split(annual, fy)
        cumulative = 0
        for i, (m_num, m_name, cal_yr, m_date) in enumerate(fy_months_list):
            cumulative += splits[i]
            alloc_rows.append({
                "AllocationID": f"MA-{alloc_id:06d}",
                "FinancialYear": fy,
                "Month": m_name,
                "FYMonthNumber": m_num,
                "CalendarYear": cal_yr,
                "BudgetVersion": row["BudgetVersion"],
                "ProgrammeCode": row["ProgrammeCode"],
                "DirectorateCode": row["DirectorateCode"],
                "CostCentreCode": row["CostCentreCode"],
                "RegionCode": row["RegionCode"],
                "FundSourceCode": row["FundSourceCode"],
                "DonorCode": row["DonorCode"],
                "ActivityCode": row["ActivityCode"],
                "SubActivityCode": row["SubActivityCode"],
                "VoteCode": row["VoteCode"],
                "MonthlyAllocationAmount": splits[i],
                "CumulativeAllocationAmount": cumulative,
                "AllocationDate": m_date.strftime("%Y-%m-%d"),
                "AllocationReference": doc_ref("MA"),
                "SourceDocument": row["SourceDocument"],
                "CreatedByRole": "Head of Budget",
                "CreatedDate": (m_date - timedelta(days=random.randint(1,5))).strftime("%Y-%m-%d"),
                "LastUpdatedDate": m_date.strftime("%Y-%m-%d"),
            })
            alloc_id += 1

df_alloc = pd.DataFrame(alloc_rows)
df_alloc.to_csv(f"{OUT}/Monthly_Budget_Allocation.csv", index=False)
print(f"    Monthly_Budget_Allocation: {len(df_alloc)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.3 SUPPLEMENTARY BUDGET
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Supplementary_Budget...")
supp_rows = []
supp_reasons = [
    "Increased utility costs","New government directive","Donor fund reallocation",
    "Emergency procurement","COVID-19 related costs","Staff costs revision",
    "Legal settlement provision","Currency fluctuation impact","Cost overrun approval",
    "Revised work plan","Emergency infrastructure repair","Donor top-up",
    "Savings redirected from underspent vote","Budget rationalisation",
]
supp_id = 1
for fy in FINANCIAL_YEARS:
    n = random.randint(60, 90)
    for _ in range(n):
        bb = df_budget[df_budget["FinancialYear"]==fy].sample(1).iloc[0]
        orig = bb["ApprovedAnnualAmount"]
        r = random.random()
        if r < 0.65:
            increase = round(random.uniform(0.02, 0.20) * orig / 1000) * 1000
            reduction = 0
        else:
            increase = 0
            reduction = round(random.uniform(0.02, 0.15) * orig / 1000) * 1000
        net = increase - reduction
        revised = orig + net
        eff_date = rand_date(FY_START[fy] + timedelta(days=60), FY_END[fy] - timedelta(days=30))
        supp_rows.append({
            "SupplementaryReference": f"SUP-{fy[:9]}-{supp_id:04d}",
            "FinancialYear": fy,
            "SupplementaryNumber": random.choice([1,2,3]),
            "ApprovalDate": eff_date.strftime("%Y-%m-%d"),
            "ProgrammeCode": bb["ProgrammeCode"],
            "DirectorateCode": bb["DirectorateCode"],
            "CostCentreCode": bb["CostCentreCode"],
            "RegionCode": bb["RegionCode"],
            "ActivityCode": bb["ActivityCode"],
            "SubActivityCode": bb["SubActivityCode"],
            "VoteCode": bb["VoteCode"],
            "FundSourceCode": bb["FundSourceCode"],
            "OriginalApprovedAmount": orig,
            "SupplementaryIncrease": increase,
            "SupplementaryReduction": reduction,
            "NetAdjustment": net,
            "RevisedApprovedAmount": max(0, revised),
            "ReasonForAdjustment": random.choice(supp_reasons),
            "ApprovalAuthority": random.choice(["Board of Directors","Cabinet Secretary","Principal Secretary"]),
            "SourceDocument": doc_ref("SUPP"),
            "EffectiveDate": eff_date.strftime("%Y-%m-%d"),
        })
        supp_id += 1

df_supp = pd.DataFrame(supp_rows)
df_supp.to_csv(f"{OUT}/Supplementary_Budget.csv", index=False)
print(f"    Supplementary_Budget: {len(df_supp)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.4 BUDGET REALLOCATIONS
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Budget_Reallocation...")
realloc_rows = []
realloc_reasons = [
    "Priority activity underfunded","Savings from completed activity",
    "Emergency requirement in destination","Vote exhaustion risk",
    "Donor compliance requirement","Work plan revision",
    "Delayed procurement - savings available","Performance contract revision",
]
realloc_id = 1
for fy in FINANCIAL_YEARS:
    n = random.randint(60, 100)
    cc_pairs = [(c1["CostCentreCode"], c2["CostCentreCode"])
                for c1 in COST_CENTRES for c2 in COST_CENTRES
                if c1["CostCentreCode"] != c2["CostCentreCode"]]
    for _ in range(n):
        src_cc, dst_cc = random.choice(cc_pairs)
        src_dir = CC_MAP[src_cc]["DirectorateCode"]
        dst_dir = CC_MAP[dst_cc]["DirectorateCode"]
        src_act = get_act_for_prg(get_prg_for_dir(src_dir))
        dst_act = get_act_for_prg(get_prg_for_dir(dst_dir))
        src_vote = random.choice(VOTE_CODES)
        dst_vote = random.choice(VOTE_CODES)
        amt = round(random.uniform(200_000, 5_000_000) / 1000) * 1000
        req_date = rand_date(FY_START[fy], FY_END[fy] - timedelta(days=60))
        days_approval = random.randint(3, 45)
        days_process = random.randint(1, 15)
        approval_date = req_date + timedelta(days=days_approval)
        process_date = approval_date + timedelta(days=days_process)
        status = random.choices(["Approved","Pending","Rejected","Returned for Clarification"],
                                weights=[0.75, 0.12, 0.08, 0.05])[0]
        realloc_rows.append({
            "ReallocationReference": f"RLC-{fy[:9]}-{realloc_id:04d}",
            "RequestDate": req_date.strftime("%Y-%m-%d"),
            "ApprovalDate": approval_date.strftime("%Y-%m-%d") if status=="Approved" else "",
            "FinancialYear": fy,
            "RequestingCostCentre": dst_cc,
            "RequestingDirectorate": dst_dir,
            "SourceCostCentre": src_cc,
            "SourceActivity": src_act,
            "SourceVoteCode": src_vote,
            "DestinationCostCentre": dst_cc,
            "DestinationActivity": dst_act,
            "DestinationVoteCode": dst_vote,
            "ReallocationAmount": amt,
            "Reason": random.choice(realloc_reasons),
            "RequestedByRole": "Cost Centre Manager",
            "ReviewedByRole": "Head of Budget",
            "ApprovedByRole": "Director, Finance and Accounts" if status=="Approved" else "",
            "Status": status,
            "ProcessingDate": process_date.strftime("%Y-%m-%d") if status=="Approved" else "",
            "SupportingDocument": doc_ref("RLC"),
            "DaysToApproval": days_approval,
            "DaysToProcess": days_process if status=="Approved" else None,
        })
        realloc_id += 1

df_realloc = pd.DataFrame(realloc_rows)
df_realloc.to_csv(f"{OUT}/Budget_Reallocation.csv", index=False)
print(f"    Budget_Reallocation: {len(df_realloc)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.5 EXPENDITURE TRANSACTIONS
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Expenditure_Transactions...")
txn_rows = []
txn_id = 1
pay_statuses = ["Paid","Pending Payment","Cancelled"]

for fy in FINANCIAL_YEARS:
    fy_months_list = fy_months(fy)
    if fy == "FY2025/2026":
        active_months = [m for m in fy_months_list if m[3] <= CUTOFF]
    else:
        active_months = fy_months_list

    n_txn = 4500 if fy == "FY2025/2026" else 3800
    for _ in range(n_txn):
        m_num, m_name, cal_yr, m_date = random.choice(active_months)
        # More spend in later months (seasonality)
        month_weight = seasonality_weights()[m_num-1]
        cc = random.choice(COST_CENTRES)
        cc_code = cc["CostCentreCode"]
        dir_code = cc["DirectorateCode"]
        prg_code = get_prg_for_dir(dir_code)
        act_code = get_act_for_prg(prg_code)
        sub_code = get_sub_for_act(act_code)
        vote_code = random.choice(VOTE_CODES)
        fs_code = random.choices(
            ["FS-01","FS-02","FS-03","FS-04","FS-05"],
            weights=[0.55, 0.25, 0.08, 0.07, 0.05])[0]
        don_code = "DON-00" if fs_code in ["FS-01","FS-02","FS-03"] else random.choice(["DON-01","DON-02","DON-03","DON-04"])
        proc_cat = random.choice(PROCUREMENT_CATEGORIES)
        gross = round(random.uniform(30_000, 600_000) * month_weight / 1000) * 1000
        tax = round(gross * random.choice([0, 0, 0.16]) / 100) * 100
        net = gross - tax
        txn_date = rand_date(m_date, min(m_date + timedelta(days=30), FY_END[fy]))
        post_date = txn_date + timedelta(days=random.randint(0, 5))
        pay_status = random.choices(pay_statuses, weights=[0.80, 0.15, 0.05])[0]
        pay_date_val = (post_date + timedelta(days=random.randint(5, 45))).strftime("%Y-%m-%d") if pay_status == "Paid" else ""
        dq_flag = random.choices(["Clean","Warning","Error"], weights=[0.96, 0.03, 0.01])[0]
        txn_rows.append({
            "TransactionID": f"TXN-{txn_id:07d}",
            "FinancialYear": fy,
            "TransactionDate": txn_date.strftime("%Y-%m-%d"),
            "PostingDate": post_date.strftime("%Y-%m-%d"),
            "DocumentNumber": doc_ref("DOC"),
            "InvoiceNumber": doc_ref("INV"),
            "SupplierID": supplier_id(),
            "SupplierCategory": random.choice(SUPPLIER_CATEGORIES),
            "ProgrammeCode": prg_code,
            "DirectorateCode": dir_code,
            "CostCentreCode": cc_code,
            "RegionCode": cc["RegionCode"],
            "FundSourceCode": fs_code,
            "DonorCode": don_code,
            "ActivityCode": act_code,
            "SubActivityCode": sub_code,
            "VoteCode": vote_code,
            "ProcurementCategory": proc_cat,
            "Description": f"{proc_cat} - {ACT_MAP[act_code]}",
            "GrossAmount": gross,
            "TaxAmount": tax,
            "NetAmount": net,
            "CommitmentReference": doc_ref("CMT") if random.random() < 0.6 else "",
            "PurchaseOrderReference": doc_ref("PO") if random.random() < 0.7 else "",
            "PaymentVoucherReference": doc_ref("PV"),
            "PaymentStatus": pay_status,
            "PaymentDate": pay_date_val,
            "AccountingPeriod": f"{fy}-M{m_num:02d}",
            "EnteredByRole": "Accounts Officer",
            "ApprovedByRole": "Chief Accountant",
            "SourceSystem": "IFMIS",
            "DataQualityFlag": dq_flag,
        })
        txn_id += 1

df_txn = pd.DataFrame(txn_rows)
df_txn.to_csv(f"{OUT}/Expenditure_Transactions.csv", index=False)
print(f"    Expenditure_Transactions: {len(df_txn)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.6 COMMITMENTS
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Commitments...")
commit_rows = []
commit_id = 1
commit_statuses = ["Open","Partially Invoiced","Fully Invoiced","Partially Paid","Fully Paid","Cancelled","Overdue"]
risk_rates = ["Critical","High","Medium","Low"]
for fy in FINANCIAL_YEARS:
    n = 550 if fy == "FY2025/2026" else 400
    for _ in range(n):
        cc = random.choice(COST_CENTRES)
        cc_code = cc["CostCentreCode"]
        dir_code = cc["DirectorateCode"]
        act_code = get_act_for_prg(get_prg_for_dir(dir_code))
        vote_code = random.choice(VOTE_CODES)
        fs_code = random.choices(["FS-01","FS-02","FS-03","FS-04"],weights=[0.55,0.25,0.10,0.10])[0]
        commit_amt = round(random.uniform(80_000, 3_500_000) / 1000) * 1000
        commit_date = rand_date(FY_START[fy], FY_END[fy] - timedelta(days=30))
        if fy == "FY2025/2026":
            commit_date = rand_date(FY_START[fy], CUTOFF - timedelta(days=10))
        invoiced_pct = random.choices([0, random.uniform(0.3,0.9), 1.0],
                                      weights=[0.25, 0.45, 0.30])[0]
        invoiced = round(commit_amt * invoiced_pct / 1000) * 1000
        paid_pct = random.uniform(0, invoiced_pct) if invoiced_pct > 0 else 0
        paid = round(commit_amt * paid_pct / 1000) * 1000
        outstanding = commit_amt - paid
        expected_del = commit_date + timedelta(days=random.randint(30, 180))
        days_outstanding = (CUTOFF - commit_date).days
        if invoiced_pct == 0:
            status = "Open" if days_outstanding < 90 else "Overdue"
        elif paid >= commit_amt:
            status = "Fully Paid"
        elif paid > 0:
            status = "Partially Paid"
        elif invoiced >= commit_amt:
            status = "Fully Invoiced"
        else:
            status = "Partially Invoiced"
        if random.random() < 0.03:
            status = "Cancelled"
            outstanding = 0
        risk = "Critical" if days_outstanding > 180 else ("High" if days_outstanding > 90 else ("Medium" if days_outstanding > 45 else "Low"))
        don_code = "DON-00" if fs_code in ["FS-01","FS-02","FS-03"] else random.choice(["DON-01","DON-02"])
        commit_rows.append({
            "CommitmentID": f"CMT-{commit_id:06d}",
            "CommitmentDate": commit_date.strftime("%Y-%m-%d"),
            "FinancialYear": fy,
            "PurchaseRequisitionReference": doc_ref("PR"),
            "PurchaseOrderReference": doc_ref("PO"),
            "ProgrammeCode": get_prg_for_dir(dir_code),
            "DirectorateCode": dir_code,
            "CostCentreCode": cc_code,
            "RegionCode": cc["RegionCode"],
            "ActivityCode": act_code,
            "VoteCode": vote_code,
            "FundSourceCode": fs_code,
            "SupplierCategory": random.choice(SUPPLIER_CATEGORIES),
            "CommitmentAmount": commit_amt,
            "AmountInvoiced": invoiced,
            "AmountPaid": paid,
            "OutstandingCommitment": outstanding,
            "ExpectedDeliveryDate": expected_del.strftime("%Y-%m-%d"),
            "CommitmentStatus": status,
            "DaysOutstanding": days_outstanding,
            "ResponsibleRole": random.choice(["Procurement Officer","Cost Centre Manager","Project Manager"]),
            "RiskRating": risk,
        })
        commit_id += 1

df_commit = pd.DataFrame(commit_rows)
df_commit.to_csv(f"{OUT}/Commitments.csv", index=False)
print(f"    Commitments: {len(df_commit)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.7 BUDGET REVIEW REQUESTS
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Budget_Review_Requests...")
req_rows = []
req_id = 1
req_purposes = [
    "Authority to Incur Expenditure","Vote reallocation request",
    "Supplementary budget request","Emergency procurement approval",
    "Training programme funding","Infrastructure repair authorisation",
    "Travel request - international","Consultancy engagement approval",
    "Equipment purchase request","Staff medical claim reimbursement",
]
req_statuses = [
    "Received","Awaiting Review","Awaiting Approval","Approved",
    "Partially Approved","Rejected","Returned for Clarification",
    "Approved but Unprocessed","Partially Processed","Fully Processed",
    "Overdue","Status Requires Review"
]
aging_bands = ["0-14 Days","15-30 Days","31-60 Days","61-90 Days","Over 90 Days"]
for fy in FINANCIAL_YEARS:
    n = 400 if fy == "FY2025/2026" else 300
    for _ in range(n):
        cc = random.choice(COST_CENTRES)
        req_date = rand_date(FY_START[fy], FY_END[fy] - timedelta(days=10))
        if fy == "FY2025/2026":
            req_date = rand_date(FY_START[fy], CUTOFF - timedelta(days=5))
        recv_date = req_date + timedelta(days=random.randint(0, 3))
        req_amt = round(random.uniform(80_000, 4_000_000) / 1000) * 1000
        status = random.choices(req_statuses, weights=[
            0.04,0.06,0.08,0.25,0.07,0.05,0.06,0.10,0.12,0.10,0.04,0.03])[0]
        approved_pct = random.uniform(0.7, 1.0) if "Approved" in status or "Processed" in status else 0
        approved_amt = round(req_amt * approved_pct / 1000) * 1000 if approved_pct > 0 else 0
        approval_date = recv_date + timedelta(days=random.randint(3,40)) if approved_amt > 0 else None
        processed_pct = random.uniform(0.5, 1.0) if "Processed" in status else 0
        processed_amt = round(approved_amt * processed_pct / 1000) * 1000
        outstanding_amt = approved_amt - processed_amt
        aging_days = (CUTOFF - recv_date).days if fy == "FY2025/2026" else random.randint(0, 180)
        if aging_days <= 14:
            aband = "0-14 Days"
        elif aging_days <= 30:
            aband = "15-30 Days"
        elif aging_days <= 60:
            aband = "31-60 Days"
        elif aging_days <= 90:
            aband = "61-90 Days"
        else:
            aband = "Over 90 Days"
        req_rows.append({
            "RequestID": f"RQT-{req_id:06d}",
            "DocumentReference": doc_ref("RQT"),
            "DocumentDate": req_date.strftime("%Y-%m-%d"),
            "DateReceived": recv_date.strftime("%Y-%m-%d"),
            "RequestingDirectorate": cc["DirectorateCode"],
            "RequestingCostCentre": cc["CostCentreCode"],
            "RegionCode": cc["RegionCode"],
            "Purpose": random.choice(req_purposes),
            "RequestedAmount": req_amt,
            "ApprovalDate": approval_date.strftime("%Y-%m-%d") if approval_date else "",
            "ApprovedAmount": approved_amt,
            "ApprovedByRole": random.choice(["Director, Finance and Accounts","Chief Executive Officer","Head of Budget"]) if approved_amt > 0 else "",
            "ApprovalComment": "Approved with conditions" if approved_amt > 0 else "",
            "ProcessReference": doc_ref("PRF"),
            "ProcessingDate": (approval_date + timedelta(days=random.randint(2,20))).strftime("%Y-%m-%d") if approval_date and processed_amt > 0 else "",
            "ProcessedAmount": processed_amt,
            "OutstandingAmount": outstanding_amt,
            "ActivityCode": get_act_for_prg(get_prg_for_dir(cc["DirectorateCode"])),
            "VoteCode": random.choice(VOTE_CODES),
            "FundSourceCode": random.choice(["FS-01","FS-02","FS-03"]),
            "DonorCode": "DON-00",
            "Status": status,
            "AgingDays": aging_days,
            "AgingBand": aband,
            "Priority": random.choice(["High","Medium","Low"]),
            "SupportingDocumentStatus": random.choices(["Complete","Incomplete","Missing"],weights=[0.75,0.15,0.10])[0],
            "DataCompletenessStatus": random.choices(["Complete","Incomplete"],weights=[0.90,0.10])[0],
            "FinancialYear": fy,
        })
        req_id += 1

df_req = pd.DataFrame(req_rows)
df_req.to_csv(f"{OUT}/Budget_Review_Requests.csv", index=False)
print(f"    Budget_Review_Requests: {len(df_req)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.8 PROCUREMENT PLAN
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Procurement_Plan...")
proc_rows = []
proc_id = 1
proc_methods = ["Open Tender","Restricted Tender","Direct Procurement","Request for Quotations","Framework Agreement","Low-Value Procurement"]
proc_statuses = ["Planned","Advertised","Evaluation","Awarded","Contract Signed","Ongoing","Completed","Cancelled","Delayed"]
for fy in FINANCIAL_YEARS:
    n = 150 if fy == "FY2025/2026" else 120
    for _ in range(n):
        cc = random.choice(COST_CENTRES)
        plan_date = rand_date(FY_START[fy], FY_END[fy] - timedelta(days=90))
        est_cost = round(random.uniform(200_000, 8_000_000) / 1000) * 1000
        app_budget = round(est_cost * random.uniform(0.95, 1.10) / 1000) * 1000
        contract_val = round(est_cost * random.uniform(0.85, 1.15) / 1000) * 1000 if random.random() < 0.75 else 0
        paid = round(contract_val * random.uniform(0.2, 1.0) / 1000) * 1000 if contract_val > 0 else 0
        outstanding = contract_val - paid
        pct_complete = round(paid / contract_val * 100) if contract_val > 0 else 0
        status = random.choice(proc_statuses)
        if fy == "FY2025/2026" and plan_date > CUTOFF:
            status = "Planned"
            contract_val = 0
            paid = 0
            outstanding = 0
            pct_complete = 0
        delay_days = random.randint(0, 120) if status in ["Delayed","Ongoing","Completed"] else 0
        cont_start = plan_date + timedelta(days=random.randint(30, 90)) if contract_val > 0 else None
        cont_end = cont_start + timedelta(days=random.randint(90, 365)) if cont_start else None
        risk = "Critical" if delay_days > 90 else ("High" if delay_days > 45 else ("Medium" if delay_days > 14 else "Low"))
        proc_rows.append({
            "ProcurementPlanID": f"PP-{proc_id:05d}",
            "FinancialYear": fy,
            "Directorate": cc["DirectorateCode"],
            "CostCentre": cc["CostCentreCode"],
            "RegionCode": cc["RegionCode"],
            "ProcurementDescription": f"{random.choice(PROCUREMENT_CATEGORIES)} - {CC_MAP[cc['CostCentreCode']]['CostCentreName']}",
            "ProcurementCategory": random.choice(PROCUREMENT_CATEGORIES),
            "ProcurementMethod": random.choice(proc_methods),
            "PlannedProcurementDate": plan_date.strftime("%Y-%m-%d"),
            "EstimatedCost": est_cost,
            "ApprovedBudget": app_budget,
            "ActualContractValue": contract_val,
            "SupplierCategory": random.choice(SUPPLIER_CATEGORIES),
            "TenderReference": doc_ref("TND") if contract_val > 0 else "",
            "ContractStartDate": cont_start.strftime("%Y-%m-%d") if cont_start else "",
            "ContractEndDate": cont_end.strftime("%Y-%m-%d") if cont_end else "",
            "ProcurementStatus": status,
            "PercentageComplete": pct_complete,
            "AmountPaid": paid,
            "OutstandingContractAmount": outstanding,
            "ResponsibleRole": "Procurement Officer",
            "DelayDays": delay_days,
            "RiskRating": risk,
        })
        proc_id += 1

df_proc = pd.DataFrame(proc_rows)
df_proc.to_csv(f"{OUT}/Procurement_Plan.csv", index=False)
print(f"    Procurement_Plan: {len(df_proc)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.9 PENDING BILLS
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Pending_Bills...")
pb_rows = []
pb_id = 1
pb_reasons = [
    "Incomplete Documentation","Insufficient Budget","Pending Inspection",
    "Pending Approval","Contract Dispute","Cash-Flow Constraint",
    "Incorrect Invoice","Procurement Exception",
]
pb_statuses_dispute = ["No Dispute","Under Review","Disputed"]
for fy in FINANCIAL_YEARS:
    n = 220 if fy == "FY2025/2026" else 150
    for _ in range(n):
        cc = random.choice(COST_CENTRES)
        inv_date = rand_date(FY_START[fy], FY_END[fy] - timedelta(days=5))
        if fy == "FY2025/2026":
            inv_date = rand_date(FY_START[fy], CUTOFF - timedelta(days=5))
        recv_date = inv_date + timedelta(days=random.randint(1, 10))
        inv_amt = round(random.uniform(40_000, 2_000_000) / 1000) * 1000
        verified_pct = random.uniform(0.7, 1.0) if random.random() < 0.80 else 0
        verified = round(inv_amt * verified_pct / 1000) * 1000
        paid = round(verified * random.uniform(0, 0.5) / 1000) * 1000
        outstanding = inv_amt - paid
        aging_days = (CUTOFF if fy == "FY2025/2026" else FY_END[fy]) - recv_date
        aging_days = aging_days.days
        if aging_days <= 30: aband = "0-30 Days"
        elif aging_days <= 90: aband = "31-90 Days"
        else: aband = "Over 90 Days"
        risk = "High" if aging_days > 90 else ("Medium" if aging_days > 30 else "Low")
        exp_pay = recv_date + timedelta(days=random.randint(30, 90))
        pb_rows.append({
            "PendingBillID": f"PB-{pb_id:06d}",
            "InvoiceNumber": doc_ref("INV"),
            "InvoiceDate": inv_date.strftime("%Y-%m-%d"),
            "DateReceived": recv_date.strftime("%Y-%m-%d"),
            "SupplierCategory": random.choice(SUPPLIER_CATEGORIES),
            "Directorate": cc["DirectorateCode"],
            "CostCentre": cc["CostCentreCode"],
            "RegionCode": cc["RegionCode"],
            "VoteCode": random.choice(VOTE_CODES),
            "FundSourceCode": random.choice(["FS-01","FS-02","FS-03"]),
            "InvoiceAmount": inv_amt,
            "AmountVerified": verified,
            "AmountPaid": paid,
            "OutstandingAmount": outstanding,
            "ReasonPending": random.choice(pb_reasons),
            "AgingDays": aging_days,
            "AgingBand": aband,
            "PaymentPriority": random.choice(["High","Medium","Low"]),
            "DisputeStatus": random.choices(pb_statuses_dispute, weights=[0.80,0.12,0.08])[0],
            "ExpectedPaymentDate": exp_pay.strftime("%Y-%m-%d"),
            "ResponsibleRole": random.choice(["Chief Accountant","Finance Officer","Accounts Officer"]),
            "RiskRating": risk,
            "FinancialYear": fy,
        })
        pb_id += 1

df_pb = pd.DataFrame(pb_rows)
df_pb.to_csv(f"{OUT}/Pending_Bills.csv", index=False)
print(f"    Pending_Bills: {len(df_pb)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.10 REVENUE AND AIA
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Revenue_AIA...")
rev_rows = []
rev_id = 1
rev_streams = [
    "Service Fees","Licensing Fees","Training Fees","Consultancy Income",
    "Facility Hire","Publication Sales","Inspection Fees","Certification Fees",
]
rev_channels = ["Bank Transfer","MPESA","Cheque","Cash","Online Portal"]
cust_cats = ["Government Agency","Private Company","NGO","Individual","International Organisation"]
for fy in FINANCIAL_YEARS:
    n = 420 if fy == "FY2025/2026" else 360
    fy_mths = fy_months(fy)
    if fy == "FY2025/2026":
        fy_mths = [m for m in fy_mths if m[3] <= CUTOFF]
    for _ in range(n):
        m_num, m_name, cal_yr, m_date = random.choice(fy_mths)
        cc = random.choice(COST_CENTRES)
        stream = random.choice(rev_streams)
        budgeted = round(random.uniform(100_000, 800_000) / 1000) * 1000
        actual_pct = random.gauss(0.88, 0.15)
        actual = round(budgeted * max(0.2, min(1.5, actual_pct)) / 1000) * 1000
        variance = actual - budgeted
        txn_date = rand_date(m_date, min(m_date + timedelta(days=30), FY_END[fy]))
        bank_days = random.randint(0, 14)
        banking_date = txn_date + timedelta(days=bank_days)
        exc_flag = "Yes" if bank_days > 5 or actual_pct < 0.5 else "No"
        rev_rows.append({
            "RevenueTransactionID": f"REV-{rev_id:06d}",
            "FinancialYear": fy,
            "TransactionDate": txn_date.strftime("%Y-%m-%d"),
            "RegionCode": cc["RegionCode"],
            "RevenueCentre": cc["CostCentreCode"],
            "RevenueStream": stream,
            "CustomerCategory": random.choice(cust_cats),
            "ReceiptReference": doc_ref("RCP"),
            "BudgetedRevenue": budgeted,
            "ActualRevenue": actual,
            "Variance": variance,
            "CollectionChannel": random.choice(rev_channels),
            "ReconciliationStatus": random.choices(["Reconciled","Unreconciled","Pending"],weights=[0.82,0.10,0.08])[0],
            "BankingDate": banking_date.strftime("%Y-%m-%d"),
            "DaysToBank": bank_days,
            "ResponsibleRole": "Revenue Manager",
            "ExceptionFlag": exc_flag,
            "Month": m_name,
            "FYMonthNumber": m_num,
        })
        rev_id += 1

df_rev = pd.DataFrame(rev_rows)
df_rev.to_csv(f"{OUT}/Revenue_AIA.csv", index=False)
print(f"    Revenue_AIA: {len(df_rev)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.11 PERFORMANCE INDICATORS
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Performance_Indicators...")
pi_rows = []
pi_id = 1
indicators = [
    ("IND-01","Number of Infrastructure Projects Completed","Projects","DIR-12"),
    ("IND-02","Percentage Budget Absorption Rate","%","DIR-02"),
    ("IND-03","Number of Staff Trained","Persons","DIR-04"),
    ("IND-04","Number of Regional Offices Operational","Offices","DIR-13"),
    ("IND-05","Percentage Procurement Plan Executed","%","DIR-05"),
    ("IND-06","Number of ICT Systems Deployed","Systems","DIR-06"),
    ("IND-07","Number of Audit Recommendations Implemented","Items","DIR-08"),
    ("IND-08","Number of Stakeholder Engagements","Events","DIR-09"),
    ("IND-09","Number of Research Reports Published","Reports","DIR-10"),
    ("IND-10","Revenue Collection Achievement","%","DIR-02"),
    ("IND-11","Number of Technical Inspections Conducted","Inspections","DIR-11"),
    ("IND-12","Number of Legal Cases Resolved","Cases","DIR-07"),
    ("IND-13","Service Delivery Satisfaction Index","%","DIR-09"),
    ("IND-14","Percentage Performance Contract Achievement","%","DIR-03"),
    ("IND-15","Number of M&E Reports Submitted","Reports","DIR-10"),
]
units = {"Projects":"Projects","%" : "%","Persons":"Persons","Offices":"Offices",
         "Systems":"Systems","Items":"Items","Events":"Events","Reports":"Reports",
         "Inspections":"Inspections","Cases":"Cases"}
quarters = ["Q1 (Jul-Sep)","Q2 (Oct-Dec)","Q3 (Jan-Mar)","Q4 (Apr-Jun)"]
for fy in FINANCIAL_YEARS:
    for q_idx, quarter in enumerate(quarters):
        if fy == "FY2025/2026" and q_idx >= 3:
            continue  # Only Q1-Q3 actual for FY2025/2026 at cutoff
        for ind_code, ind_name, unit, dir_code in indicators:
            cc = random.choice([c for c in COST_CENTRES if c["DirectorateCode"]==dir_code] or COST_CENTRES)
            annual_target = random.randint(4, 120) if unit != "%" else random.randint(60, 100)
            q_target = round(annual_target / 4)
            perf_factor = random.gauss(0.88, 0.18)
            actual = round(q_target * max(0.1, min(1.5, perf_factor)))
            if unit == "%":
                actual = min(100, max(0, actual))
            pct = round(actual / q_target * 100, 1) if q_target > 0 else 0
            budget_alloc = round(random.uniform(2_000_000, 30_000_000) / 1000) * 1000
            exp = round(budget_alloc * random.uniform(0.4, 1.1) / 1000) * 1000
            exp = min(exp, budget_alloc * 1.05)
            cost_per_output = round(exp / max(1, actual)) if actual > 0 else 0
            if pct >= 90: p_status = "On Track"
            elif pct >= 70: p_status = "At Risk"
            else: p_status = "Off Track"
            rep_date = FY_START[fy] + timedelta(days=90*(q_idx+1))
            pi_rows.append({
                "PerformanceID": f"PI-{pi_id:05d}",
                "FinancialYear": fy,
                "Quarter": quarter,
                "ProgrammeCode": get_prg_for_dir(dir_code),
                "DirectorateCode": dir_code,
                "CostCentreCode": cc["CostCentreCode"],
                "RegionCode": cc["RegionCode"],
                "IndicatorCode": ind_code,
                "IndicatorName": ind_name,
                "UnitOfMeasure": unit,
                "AnnualTarget": annual_target,
                "QuarterlyTarget": q_target,
                "ActualAchievement": actual,
                "AchievementPercentage": pct,
                "BudgetAllocated": budget_alloc,
                "Expenditure": exp,
                "CostPerOutput": cost_per_output,
                "PerformanceStatus": p_status,
                "DataSource": random.choice(["Management Information System","Field Reports","HR System","Finance System"]),
                "ReportingDate": rep_date.strftime("%Y-%m-%d"),
                "ResponsibleRole": "M&E Officer",
            })
            pi_id += 1

df_pi = pd.DataFrame(pi_rows)
df_pi.to_csv(f"{OUT}/Performance_Indicators.csv", index=False)
print(f"    Performance_Indicators: {len(df_pi)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.12 AUDIT AND CONTROL EXCEPTIONS
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Audit_Control_Exceptions...")
exc_rows = []
exc_id = 1
exc_types = [
    "Missing Supporting Documents","Duplicate Payment Risk","Incorrect Vote Code",
    "Expenditure Above Budget","Delayed Banking","Unsupported Reallocation",
    "Long-Outstanding Commitment","Procurement Non-Compliance","Unreconciled Transaction",
    "Approval-Control Gap",
]
exc_statuses = ["Open","In Progress","Resolved","Escalated","Deferred"]
for fy in FINANCIAL_YEARS:
    n = 120 if fy == "FY2025/2026" else 90
    for _ in range(n):
        cc = random.choice(COST_CENTRES)
        det_date = rand_date(FY_START[fy], FY_END[fy] - timedelta(days=10))
        if fy == "FY2025/2026":
            det_date = rand_date(FY_START[fy], CUTOFF)
        exc_type = random.choice(exc_types)
        amt_risk = round(random.uniform(30_000, 1_500_000) / 1000) * 1000
        risk = random.choices(["Critical","High","Medium","Low"],weights=[0.10,0.25,0.40,0.25])[0]
        tgt_date = det_date + timedelta(days=random.randint(14, 90))
        status = random.choices(exc_statuses, weights=[0.30,0.25,0.30,0.10,0.05])[0]
        res_date = (tgt_date + timedelta(days=random.randint(-10,30))).strftime("%Y-%m-%d") if status == "Resolved" else ""
        days_out = (CUTOFF - det_date).days if fy == "FY2025/2026" else (FY_END[fy] - det_date).days
        repeat = random.choices(["Yes","No"], weights=[0.20,0.80])[0]
        exc_rows.append({
            "ExceptionID": f"EXC-{exc_id:06d}",
            "DetectionDate": det_date.strftime("%Y-%m-%d"),
            "FinancialYear": fy,
            "DirectorateCode": cc["DirectorateCode"],
            "CostCentreCode": cc["CostCentreCode"],
            "RegionCode": cc["RegionCode"],
            "TransactionReference": doc_ref("TXN"),
            "ExceptionType": exc_type,
            "ExceptionDescription": f"{exc_type} identified in {CC_MAP[cc['CostCentreCode']]['CostCentreName']}",
            "AmountAtRisk": amt_risk,
            "RiskRating": risk,
            "ResponsibleRole": random.choice(["Cost Centre Manager","Chief Accountant","Procurement Officer"]),
            "CorrectiveAction": f"Obtain missing documents and repost transaction",
            "TargetResolutionDate": tgt_date.strftime("%Y-%m-%d"),
            "ResolutionDate": res_date,
            "Status": status,
            "DaysOutstanding": days_out,
            "RepeatFindingFlag": repeat,
        })
        exc_id += 1

df_exc = pd.DataFrame(exc_rows)
df_exc.to_csv(f"{OUT}/Audit_Control_Exceptions.csv", index=False)
print(f"    Audit_Control_Exceptions: {len(df_exc)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# 6.13 DATA QUALITY LOG (Intentional Issues)
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating Data_Quality_Log (with intentional issues)...")
dq_rows = []
dq_id = 1
dq_issue_types = [
    ("Missing Cost Centre Code","CostCentreCode","","CC-XXX","Replace with default Unknown CC","High"),
    ("Missing Vote Code","VoteCode","","V-9999","Replace with Unclassified vote","High"),
    ("Duplicate Transaction Reference","TransactionID","TXN-0001234","TXN-0001234-DUP","Deduplicate and retain first occurrence","Critical"),
    ("Inconsistent Region Name","RegionCode","nairobi region","RG-01","Standardise to uppercase code","Medium"),
    ("Leading Trailing Space","DirectorateCode"," DIR-02 ","DIR-02","Trim whitespace","Low"),
    ("Different Capitalisation","FundSourceName","government of kenya - recurrent","Government of Kenya - Recurrent","Apply title case","Low"),
    ("Invalid Date","TransactionDate","2024-13-01","2024-01-13","Correct transposed date","High"),
    ("Blank Approval Field","ApprovedByRole","","Head of Budget","Populate from workflow","Medium"),
    ("Processed Amount Exceeds Approved","ProcessedAmount","3200000 > 3000000","3000000","Cap at approved amount","Critical"),
    ("Monthly Allocation Mismatch","MonthlyAllocationAmount","Sum=1410000","1400000","Rebalance to annual total","High"),
    ("Unmatched Lookup Code","ActivityCode","ACT-999","","Remove or map to valid activity","High"),
    ("Duplicate Budget Request","RequestID","RQT-000123 x2","RQT-000123-A","Rename duplicate","Medium"),
    ("Negative Expenditure Amount","NetAmount","-150000","150000","Correct sign","Critical"),
    ("Missing Supporting Document","SupportingDocumentStatus","Missing","Obtained","Attach document","High"),
    ("Incorrect Financial Year","FinancialYear","FY2024/2025","FY2025/2026","Reassign to correct FY","High"),
    ("Inconsistent Donor Name","DonorName","world bank","World Bank","Standardise to proper noun","Medium"),
    ("Incorrect Data Type","GrossAmount","One million","1000000","Convert text to numeric","Critical"),
    ("Formula Error","ApprovedAnnualAmount","#REF!","2500000","Recalculate from source","High"),
]

# Generate ~250 DQ records spread across datasets
datasets = ["Expenditure_Transactions","Budget_Review_Requests","Monthly_Budget_Allocation",
            "Commitments","Approved_Annual_Budget","Budget_Reallocation","Pending_Bills"]
for i in range(250):
    issue = random.choice(dq_issue_types)
    sev = issue[5]
    dq_rows.append({
        "DQLogID": f"DQL-{dq_id:05d}",
        "RecordIdentifier": doc_ref("REC"),
        "Dataset": random.choice(datasets),
        "Field": issue[1],
        "IssueType": issue[0],
        "OriginalValue": issue[2],
        "CorrectedValue": issue[3],
        "CorrectionRule": issue[4],
        "ResolutionStatus": random.choices(["Resolved","Pending","In Review"],weights=[0.65,0.25,0.10])[0],
        "Severity": sev,
        "DateCorrected": rand_date(date(2025,7,1), CUTOFF).strftime("%Y-%m-%d") if random.random()>0.3 else "",
        "FinancialYear": random.choice(FINANCIAL_YEARS),
        "CorrectedByRole": "Data Officer",
    })
    dq_id += 1

df_dq = pd.DataFrame(dq_rows)
df_dq.to_csv(f"{OUT}/Data_Quality_Log.csv", index=False)
print(f"    Data_Quality_Log: {len(df_dq)} rows")

# ══════════════════════════════════════════════════════════════════════════════
# DIMENSION / LOOKUP TABLES
# ══════════════════════════════════════════════════════════════════════════════
print("  Generating dimension tables...")

# Dim_Date
date_rows = []
start_d = date(2023, 7, 1)
end_d   = date(2026, 6, 30)
d = start_d
while d <= end_d:
    fy = "FY2023/2024" if d < date(2024,7,1) else ("FY2024/2025" if d < date(2025,7,1) else "FY2025/2026")
    fy_m = ((d.month - 7) % 12) + 1
    q_num = (fy_m - 1) // 3 + 1
    date_rows.append({
        "DateKey": int(d.strftime("%Y%m%d")),
        "Date": d.strftime("%Y-%m-%d"),
        "Day": d.day,"Month": d.month,"MonthName": d.strftime("%B"),
        "Quarter": f"Q{q_num}","CalendarYear": d.year,
        "FinancialYear": fy,"FYMonthNumber": fy_m,
        "FYQuarter": f"Q{q_num} ({['Jul','Oct','Jan','Apr'][q_num-1]}-{['Sep','Dec','Mar','Jun'][q_num-1]})",
        "WeekDay": d.strftime("%A"),
        "IsWeekend": "Yes" if d.weekday() >= 5 else "No",
        "IsPublicHoliday": "No",
    })
    d += timedelta(days=1)
pd.DataFrame(date_rows).to_csv(f"{OUT}/Dim_Date.csv", index=False)

pd.DataFrame([{"FinancialYearCode":fy,"FinancialYearName":fy,
               "StartDate":FY_START[fy].strftime("%Y-%m-%d"),
               "EndDate":FY_END[fy].strftime("%Y-%m-%d"),
               "IsActive": "Yes" if fy=="FY2025/2026" else "No"}
              for fy in FINANCIAL_YEARS]).to_csv(f"{OUT}/Dim_Financial_Year.csv",index=False)

pd.DataFrame([{"ProgrammeCode":p[0],"ProgrammeName":p[1]} for p in PROGRAMMES]).to_csv(f"{OUT}/Dim_Programme.csv",index=False)
pd.DataFrame([{"DirectorateCode":d[0],"DirectorateName":d[1]} for d in DIRECTORATES]).to_csv(f"{OUT}/Dim_Directorate.csv",index=False)
pd.DataFrame(COST_CENTRES).to_csv(f"{OUT}/Dim_Cost_Centre.csv",index=False)
pd.DataFrame([{"RegionCode":r[0],"RegionName":r[1]} for r in REGIONS]).to_csv(f"{OUT}/Dim_Region.csv",index=False)
pd.DataFrame([{"ActivityCode":a[0],"ActivityName":a[1],"ProgrammeCode":a[2]} for a in ACTIVITIES]).to_csv(f"{OUT}/Dim_Activity.csv",index=False)
pd.DataFrame([{"SubActivityCode":s[0],"SubActivityName":s[1],"ActivityCode":s[2]} for s in SUB_ACTIVITIES]).to_csv(f"{OUT}/Dim_Sub_Activity.csv",index=False)
pd.DataFrame([{"VoteCode":v[0],"VoteDescription":v[1],"EconomicType":v[2],"EconomicClassification":v[3]} for v in VOTES]).to_csv(f"{OUT}/Dim_Vote.csv",index=False)
pd.DataFrame([{"FundSourceCode":f[0],"FundSourceName":f[1]} for f in FUND_SOURCES]).to_csv(f"{OUT}/Dim_Fund_Source.csv",index=False)
pd.DataFrame([{"DonorCode":d[0],"DonorName":d[1]} for d in DONORS]).to_csv(f"{OUT}/Dim_Donor.csv",index=False)
pd.DataFrame([{"Category":c} for c in PROCUREMENT_CATEGORIES]).to_csv(f"{OUT}/Dim_Procurement_Category.csv",index=False)
pd.DataFrame([{"RiskRating":r,"SortOrder":i+1} for i,r in enumerate(RISK_RATINGS)]).to_csv(f"{OUT}/Dim_Risk_Rating.csv",index=False)
pd.DataFrame([{"OfficerRole":r} for r in OFFICER_ROLES]).to_csv(f"{OUT}/Dim_Officer_Role.csv",index=False)
pd.DataFrame([{"SupplierCategory":s} for s in SUPPLIER_CATEGORIES]).to_csv(f"{OUT}/Dim_Supplier_Category.csv",index=False)

print("  Dimension tables generated.")

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
print("\n=== DATA GENERATION COMPLETE ===")
total_rows = (len(df_budget)+len(df_alloc)+len(df_supp)+len(df_realloc)+
              len(df_txn)+len(df_commit)+len(df_req)+len(df_proc)+
              len(df_pb)+len(df_rev)+len(df_pi)+len(df_exc)+len(df_dq))
print(f"Total rows generated: {total_rows:,}")
print(f"\nApproved Annual Budget  : {len(df_budget):>7,} rows  | Total Budget FY2025/26: KSh {df_budget[df_budget.FinancialYear=='FY2025/2026'].ApprovedAnnualAmount.sum():>15,.0f}")
print(f"Monthly Allocations     : {len(df_alloc):>7,} rows")
print(f"Supplementary Budget    : {len(df_supp):>7,} rows")
print(f"Budget Reallocations    : {len(df_realloc):>7,} rows")
print(f"Expenditure Transactions: {len(df_txn):>7,} rows  | Total Expenditure FY2025/26 YTD: KSh {df_txn[df_txn.FinancialYear=='FY2025/2026'].NetAmount.sum():>15,.0f}")
print(f"Commitments             : {len(df_commit):>7,} rows")
print(f"Budget Review Requests  : {len(df_req):>7,} rows")
print(f"Procurement Plan        : {len(df_proc):>7,} rows")
print(f"Pending Bills           : {len(df_pb):>7,} rows")
print(f"Revenue / AIA           : {len(df_rev):>7,} rows")
print(f"Performance Indicators  : {len(df_pi):>7,} rows")
print(f"Audit Control Exceptions: {len(df_exc):>7,} rows")
print(f"Data Quality Log        : {len(df_dq):>7,} rows")
