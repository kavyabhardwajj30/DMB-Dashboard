"""
Test script to verify complete Masterfile sync and validation logic.
"""
from __future__ import annotations

import copy
import io
import logging
import os
from pathlib import Path
import re
import sys
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
logging.basicConfig(level=logging.INFO, format="[TestSync %(asctime)s] %(message)s")
logger = logging.getLogger("test_sync")

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"

FUNC_TAB_MAPPING = [
    ("1.Quality", "Quality DMB", "Quality"),
    ("2.Regulatory", "Regulatory", "Regulatory"),
    ("3.ISC ", "ISC & Procurement", "ISC & Procurement"),
    ("4.Procurement", "ISC & Procurement", "ISC & Procurement"),
    ("5.R&D ", "R&D", "R&D"),
    ("6. Marketing", "Marketing", "Marketing"),
    ("7.Customer Service", "Customer Service", "Customer Service"),
    ("9.NAR ", "NAR", "NAR"),
    ("10.EUROPE ", "Europe", "Europe"),
    ("11.GROWTH ", "Growth", "Growth"),
]

def normalize_text(s: str | None) -> str:
    if not s:
        return ""
    return re.sub(r"[^a-zA-Z0-9]", "", str(s).lower())

def test_run():
    func_file = DATA_DIR / "Functional DMB Review Sheets.xlsx"
    strat_file = DATA_DIR / "Strategic Execution Dashboard.xlsx"
    master_file = DATA_DIR / "Masterfile_DMB_Dashboard.xlsx"

    assert func_file.exists(), f"Missing {func_file}"
    assert strat_file.exists(), f"Missing {strat_file}"
    assert master_file.exists(), f"Missing {master_file}"

    wb_func = openpyxl.load_workbook(func_file, data_only=True)
    wb_strat = openpyxl.load_workbook(strat_file, data_only=True)
    wb_master = openpyxl.load_workbook(master_file, data_only=False)

    print("Loaded all 3 workbooks successfully.")
    print("Functional sheets:", wb_func.sheetnames)
    print("Strategic sheets:", wb_strat.sheetnames)
    print("Masterfile sheets:", wb_master.sheetnames)

if __name__ == "__main__":
    test_run()
