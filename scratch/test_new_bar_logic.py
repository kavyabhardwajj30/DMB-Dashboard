import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import data_loader

dmb_data = data_loader.get_data_store().get_dmb_data()
months = ['2026-01-01', '2026-04-01', '2026-06-01', '2026-08-01']

functions = ["Quality", "Customer Service", "Marketing", "ISC & Procurement", "Regulatory", "R&D"]

for m_str in months:
    m = pd.Timestamp(m_str)
    current_data = dmb_data[dmb_data['month'].eq(m)].copy()
    print(f"\n=================== Month: {m.strftime('%B %Y')} ===================")
    for fn in functions:
        if fn == "ISC & Procurement":
            fn_rows = current_data[current_data["function"].isin(["ISC & Procurement", "ISC", "Procurement"])].copy()
            known_kpis = set(dmb_data[dmb_data["function"].isin(["ISC & Procurement", "ISC", "Procurement"])]["kpi_name"].dropna().unique())
        else:
            fn_rows = current_data[current_data["function"].eq(fn)].copy()
            known_kpis = set(dmb_data[dmb_data["function"].eq(fn)]["kpi_name"].dropna().unique())

        month_kpis = set(fn_rows["kpi_name"].dropna().unique()) if not fn_rows.empty else set()
        all_kpis = known_kpis.union(month_kpis)
        total_expected = max(len(all_kpis), len(fn_rows))

        valid = fn_rows[fn_rows["Actual"].notna() & fn_rows["Target"].notna()] if not fn_rows.empty else pd.DataFrame()
        met = int(valid["is_met"].sum()) if not valid.empty else 0
        not_met = int(valid["status"].eq("Not Met").sum()) if not valid.empty else 0
        filled = met + not_met
        not_filled = max(0, total_expected - filled)
        total = filled + not_filled

        met_pct = round((met / total) * 100, 1) if total else 0
        not_met_pct = round((not_met / total) * 100, 1) if total else 0
        not_filled_pct = round((not_filled / total) * 100, 1) if total else 0

        print(f"{fn:22} | Total: {total:2d} | Met: {met:2d} ({met_pct:5.1f}%) | Not Met: {not_met:2d} ({not_met_pct:5.1f}%) | Not Filled: {not_filled:2d} ({not_filled_pct:5.1f}%)")
