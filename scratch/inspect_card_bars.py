import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app

cards = app.get_dmb_function_cards_content("2026-08-01")
for card in cards:
    # function header
    fn_name = card.children[0].children[0].children
    sub = card.children[1].children
    bar = card.children[2]
    segments = [(seg.children, getattr(seg, 'className', ''), seg.style) for seg in bar.children]
    print(f"\nFunction: {fn_name}")
    print(f"  Bar segments: {segments}")
