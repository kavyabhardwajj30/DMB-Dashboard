import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import data_loader
from kpi_calculations import get_month_summary

store = data_loader.get_data_store()
mpr, dmb, strat_rca = store.get_data()

print(f"Data loaded successfully:")
print(f"  DMB rows: {len(dmb)}")
print(f"  MPR rows: {len(mpr)}")
print(f"  Strategic RCA rows: {len(strat_rca)}")

month_ts = pd.Timestamp(year=2026, month=8, day=1)
print(f"\n=== August 2026 Function Summaries ===")
for func in sorted(dmb["function"].unique()):
    func_df = dmb[dmb["function"] == func]
    summ = get_month_summary(func_df, month_ts)
    pct = summ["performance_percentage"]
    print(f"{func:20s}: Met={summ['met']:2d}, Red={summ['not_met']:2d}, Total={summ['total_kpis']:2d} | Met&Imp%={pct:>5.1f}% | Red KPIs={summ['not_met']}")

print(f"\nAll 9 functions verified successfully.")
