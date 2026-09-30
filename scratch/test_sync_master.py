import sys
import io
import re
import time
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent.parent / "DMB-Dashboard"
DATA_DIR = BASE_DIR / "data"

MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTHS_LONG = [f"{m}-2026" for m in MONTHS_SHORT]

def clean_str(val):
    if val is None or pd.isna(val):
        return ""
    return str(val).replace("\xa0", " ").strip()

def find_source_files():
    func_files = sorted(
        [p for p in DATA_DIR.glob("*.xlsx") if not p.name.startswith("~$") and ("functional" in p.name.lower() or "review" in p.name.lower()) and "master" not in p.name.lower()],
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    strat_files = sorted(
        [p for p in DATA_DIR.glob("*.xlsx") if not p.name.startswith("~$") and ("strategic" in p.name.lower() or "aop" in p.name.lower()) and "master" not in p.name.lower()],
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    return (func_files[0] if func_files else None, strat_files[0] if strat_files else None)

print("Found files:", find_source_files())
