import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from app import get_dynamic_reporting_months, get_active_rca_actions, get_active_mpr_data, get_active_dmb_data

rca = get_active_rca_actions()
mpr = get_active_mpr_data()
dmb = get_active_dmb_data()

m_now, def_now = get_dynamic_reporting_months(mpr, rca)
print("Current date (Sept 2026) months:", [m.strftime("%b %Y") for m in m_now])
print("Current default:", def_now.strftime("%b %Y"))

m_oct, def_oct = get_dynamic_reporting_months(mpr, rca, reference_date="2026-10-01")
print("\nOctober 1, 2026 months:", [m.strftime("%b %Y") for m in m_oct])
print("October 1 default:", def_oct.strftime("%b %Y"))

m_nov, def_nov = get_dynamic_reporting_months(mpr, rca, reference_date="2026-11-01")
print("\nNovember 1, 2026 months:", [m.strftime("%b %Y") for m in m_nov])
print("November 1 default:", def_nov.strftime("%b %Y"))
