# Parastatal X — Code Documentation

## Scripts Overview

| Script | Purpose | Run Order |
|---|---|---|
| `generate_synthetic_data.py` | Generates all 13 CSVs + 14 dimension tables | 1 |
| `validate_data.py` | Runs 36 automated validation checks | 2 |
| `financial_forecasting.py` | Produces 4-method forecast + cost-centre forecasts | 3 |
| `reconciliation_checks.py` | Runs 18 reconciliation balance checks | 4 |
| `build_excel_fast.py` | Builds 26-sheet Excel workbook | 5 |
| `build_docs.py` | Builds Power BI DOCX files | 6 |
| `build_analysis_docs.py` | Builds management analysis documents | 7 |
| `build_portfolio_docs.py` | Builds case study and interview talking points | 8 |
| `build_presentation.py` | Builds 15-slide PPTX presentation | 9 |

## Quick Start

```bash
cd 02_Code
pip install -r requirements.txt
python generate_synthetic_data.py    # ~5 seconds
python validate_data.py              # ~2 seconds — all 36 checks pass
python financial_forecasting.py      # ~2 seconds
python reconciliation_checks.py      # ~2 seconds
python build_excel_fast.py           # ~15 seconds
python build_docs.py                 # ~10 seconds
python build_analysis_docs.py        # ~20 seconds
python build_portfolio_docs.py       # ~10 seconds
python build_presentation.py         # ~5 seconds
```

## Key Parameters

- Random seed: `42` (set at top of `generate_synthetic_data.py`)
- Budget scale: KSh 1.70B (FY2025/2026)
- Active FY: FY2025/2026
- Analysis cut-off: 31 May 2026

**Disclaimer: All generated data is entirely synthetic.**
