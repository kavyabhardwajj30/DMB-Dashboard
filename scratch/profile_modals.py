import sys
import os
import time

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import app
import data_loader

print("Loading data...")
mpr_data = app.get_active_mpr_data()
dmb_data = app.get_active_dmb_data()
strat_rca = app.get_active_rca_actions()

months, default_month = app.get_dynamic_reporting_months(dmb_data, strat_rca)
month_str = default_month.strftime("%Y-%m-%d")
print(f"Testing for month: {month_str}")

functions = ["Quality", "Regulatory", "ISC & Procurement", "R&D", "Customer Service", "Marketing"]

print("\n--- Measuring continuous red modal generation ---")
for fn in functions:
    t0 = time.perf_counter()
    res = app.get_continuous_red_modal_content(fn, month_str)
    t1 = time.perf_counter()
    print(f"RCA Modal for {fn:20s}: {(t1 - t0)*1000:.2f} ms")

print("\n--- Measuring function gauge modal generation ---")
for fn in functions:
    t0 = time.perf_counter()
    res = app.create_function_kpis_table_detail(fn, default_month)
    t1 = time.perf_counter()
    print(f"Gauge Modal for {fn:20s}: {(t1 - t0)*1000:.2f} ms")
