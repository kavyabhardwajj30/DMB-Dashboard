import pandas as pd
import data_loader

dmb = data_loader.get_data_store().get_dmb_data()
for fn in ['Quality', 'Customer Service', 'Marketing', 'ISC & Procurement', 'Regulatory', 'R&D', 'Europe', 'Growth', 'NAR', 'Commercial Excellence']:
    if fn == 'ISC & Procurement':
        f_all = dmb[dmb['function'].isin(['ISC & Procurement', 'ISC', 'Procurement'])]
    else:
        f_all = dmb[dmb['function'] == fn]
    kpis = f_all['kpi_name'].dropna().unique().tolist()
    print(f"\nFunction: {fn} (Total unique KPIs across months: {len(kpis)})")
    print(f"  KPIs: {kpis}")
    for m in sorted(f_all['month'].dropna().unique()):
        fm = f_all[f_all['month'] == m]
        valid = fm[fm['Actual'].notna() & fm['Target'].notna()]
        met = (valid['status'] == 'Met').sum()
        not_met = (valid['status'] == 'Not Met').sum()
        not_filled = len(kpis) - len(valid)
        print(f"    {m.strftime('%Y-%m-%d')}: defined_in_month={len(fm)}, valid={len(valid)}, met={met}, not_met={not_met}, unpopulated/not_filled={not_filled}")
