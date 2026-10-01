import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from app import get_dynamic_reporting_months, get_active_rca_actions, get_active_mpr_data, get_active_dmb_data

rca = get_active_rca_actions()
mpr = get_active_mpr_data()
dmb = get_active_dmb_data()

m_now, def_now = get_dynamic_reporting_months(mpr, rca)
print("Current date months:", [m.strftime("%b %Y") for m in m_now])
print("Current default:", def_now.strftime("%b %Y"))

m_oct, def_oct = get_dynamic_reporting_months(mpr, rca, reference_date="2026-10-01")
print("\nOctober 1, 2026 months:", [m.strftime("%b %Y") for m in m_oct])
print("October 1 default:", def_oct.strftime("%b %Y"))

m_nov, def_nov = get_dynamic_reporting_months(mpr, rca, reference_date="2026-11-01")
print("\nNovember 1, 2026 months:", [m.strftime("%b %Y") for m in m_nov])
print("November 1 default:", def_nov.strftime("%b %Y"))

m_dec, def_dec = get_dynamic_reporting_months(mpr, rca, reference_date="2026-12-01")
print("\nDecember 1, 2026 months:", [m.strftime("%b %Y") for m in m_dec])
print("December 1 default:", def_dec.strftime("%b %Y"))

m_jan27, def_jan27 = get_dynamic_reporting_months(mpr, rca, reference_date="2027-01-01")
print("\nJanuary 1, 2027 months:", [m.strftime("%b %Y") for m in m_jan27])
print("January 1 default:", def_jan27.strftime("%b %Y"))

m_feb27, def_feb27 = get_dynamic_reporting_months(mpr, rca, reference_date="2027-02-01")
print("\nFebruary 1, 2027 months:", [m.strftime("%b %Y") for m in m_feb27])
print("February 1 default:", def_feb27.strftime("%b %Y"))

m_mar27, def_mar27 = get_dynamic_reporting_months(mpr, rca, reference_date="2027-03-01")
print("\nMarch 1, 2027 months:", [m.strftime("%b %Y") for m in m_mar27])
print("March 1 default:", def_mar27.strftime("%b %Y"))
