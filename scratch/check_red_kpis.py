import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
import app
import pandas as pd

def format_val(val, unit):
    if val is None or pd.isna(val):
        return "—"
    try:
        val_f = float(val)
    except (ValueError, TypeError):
        return f"{val} {unit}".strip()
    unit_clean = str(unit).strip() if unit and pd.notna(unit) else ""
    if unit_clean == "%":
        if 0 < val_f <= 1.0:
            val_f *= 100
        return f"{val_f:.1f}%" if val_f % 1 != 0 else f"{int(val_f)}%"
    elif unit_clean == "Mn":
        return f"{val_f:.2f} Mn"
    elif unit_clean.lower() in ["no.", "no", "nos"]:
        return f"{int(val_f)}" if val_f == int(val_f) else f"{val_f:.1f}"
    else:
        suffix = f" {unit_clean}" if unit_clean else ""
        return f"{val_f:.2f}{suffix}" if val_f % 1 != 0 else f"{int(val_f)}{suffix}"

dmb = app.get_active_dmb_data()

print("=== ALL RED KPIS FOR ISC & PROCUREMENT ACROSS 2026 ===")
isc = dmb[dmb["function"].isin(["ISC & Procurement", "ISC", "Procurement"]) & (dmb["status"] == "Not Met")]
for month in sorted(isc["month"].unique()):
    m_df = isc[isc["month"] == month]
    m_str = pd.Timestamp(month).strftime("%B %Y")
    print(f"\n--- Reporting Month: {m_str} ---")
    for _, r in m_df.iterrows():
        t_str = format_val(r["Target"], r["units"])
        a_str = format_val(r["Actual"], r["units"])
        c_red = " (Continuous-Red)" if r.get("is_continuous_red") else ""
        print(f"  • {r['kpi_name']}{c_red}: Target = {t_str} | Actual = {a_str} (Operator: {r.get('operator')})")
