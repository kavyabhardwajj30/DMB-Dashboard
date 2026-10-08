import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
import app
import pandas as pd

detail = app.create_continuous_red_detail("ISC & Procurement", "2026-07-01")
print("Successfully generated ISC & Procurement modal content for July 2026!")

# Check if target and actual are rendered in the card header
from dash import html

def inspect_component(c, depth=0):
    if hasattr(c, "children"):
        ch = c.children
        if isinstance(ch, list):
            for child in ch:
                inspect_component(child, depth + 1)
        elif ch:
            inspect_component(ch, depth + 1)
    if hasattr(c, "className"):
        cls = getattr(c, "className", "")
        if "rca-recovery-card-header" in str(cls) or "rca-metric-pill" in str(cls):
            print("  " * depth + f"<{c.__class__.__name__} class='{cls}'> {getattr(c, 'children', '')}")

print("\nInspecting rendered structure:")
inspect_component(detail)
