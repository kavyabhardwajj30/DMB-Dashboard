import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import get_function_rca_data, create_continuous_red_detail

print("=== TESTING CONTINUOUS RED DETAIL FOR ALL FUNCTIONS ===")
for fn in ["ISC & Procurement", "Customer Service", "Quality", "Regulatory", "Marketing", "R&D", "NAR", "Europe", "Growth"]:
    try:
        detail = create_continuous_red_detail(fn, "2026-07-01")
        print(f"Function: {fn:20} -> Rendered successfully!")
    except Exception as e:
        print(f"ERROR on {fn}: {e}")
