import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app

print("Testing app.serve_layout()...")
layout = app.serve_layout()
print("Layout created successfully!")

dmb_data = app.get_active_dmb_data()
months = sorted(dmb_data['month'].dropna().unique())
for m in months:
    m_str = m.strftime('%Y-%m-%d')
    cards = app.get_dmb_function_cards_content(m_str)
    print(f"Month {m.strftime('%B %Y')}: rendered {len(cards)} function cards.")

print("\nALL DMB FUNCTION CARDS AND LAYOUT TESTED WITH ZERO ERRORS!")
