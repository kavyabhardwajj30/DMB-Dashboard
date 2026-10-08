import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
import app
import pandas as pd

dmb = app.get_active_dmb_data()
aug_isc = dmb[(dmb['month'] == '2026-08-01') & dmb['function'].isin(['ISC & Procurement', 'ISC', 'Procurement'])]
for _, r in aug_isc.iterrows():
    print(f"KPI: '{r['kpi_name']}' | Target: {r['Target']} | Actual: {r['Actual']} | Units: '{r['units']}' | Status: {r['status']}")

print("\n--- Let's trace how cards are created for ISC in August 2026 ---")
rca_data = app.get_function_rca_data()
review_kpis = rca_data["review_kpis"]
rev_aug = review_kpis[(review_kpis["function_key"] == "isc_procurement") & (review_kpis["month"] == "2026-08-01")]
print("Review KPIs for ISC in Aug 2026:")
for _, r in rev_aug.iterrows():
    print(f"  Sheet: '{r['source_sheet']}' | KPI: '{r['kpi_name']}' | Target: {r['target']} | Actual: {r['actual']} | Status: {r['status']}")

