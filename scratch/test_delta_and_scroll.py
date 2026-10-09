import sys
import os
sys.path.insert(0, os.path.abspath("."))
from app import create_continuous_red_detail, get_active_dmb_data, compute_kpi_delta_badge, kpi_card

print("=== 1. Testing compute_kpi_delta_badge ===")
test_cases = [
    ("95.0%", "83.0%"),
    ("54.4%", "53.6%"),
    ("5.87 Mn", "5.77 Mn"),
    ("80%", "85%"),
    ("10", "10"),
    (0.544, 0.536),
    (None, "80%"),
    ("—", "50%"),
]
for t, a in test_cases:
    res = compute_kpi_delta_badge(t, a)
    if res:
        print(f"Target: {str(t):10} | Actual: {str(a):10} -> {res['label']} {res['val'].encode('ascii', 'replace').decode()} ({res['class']})")
    else:
        print(f"Target: {str(t):10} | Actual: {str(a):10} -> None")

print("\n=== 2. Testing kpi_card ===")
card = kpi_card("KPIs not met", "mpr-not-met-kpis", "#dc3d56", initial_value=5)
print("Card ID:", getattr(card, "id", None))
print("Card className:", getattr(card, "className", None))
print("Card title:", getattr(card, "title", None))

print("\n=== 3. Testing create_continuous_red_detail ===")
for fn in ["Customer Service", "ISC & Procurement", "Quality"]:
    detail = create_continuous_red_detail(fn, "2026-07-01")
    print(f"Rendered modal content for {fn}: SUCCESS (type: {type(detail)})")
