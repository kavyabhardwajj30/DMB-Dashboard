"""
Complete synchronization and validation test.
"""
from __future__ import annotations

import copy
import hashlib
import io
import logging
import os
from pathlib import Path
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
logging.basicConfig(level=logging.INFO, format="[SyncEngine %(asctime)s] %(message)s")
logger = logging.getLogger("dmb.sync_engine")

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"

MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def clean_text(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).replace("\xa0", " ").strip()
    if s.lower() in {"none", "nan", "null"}:
        return ""
    return s

def norm_key(s: Any) -> str:
    return re.sub(r"[^a-zA-Z0-9]", "", str(s or "").lower())

def read_safe_bytes(path: Path | str) -> bytes:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {p}")
    if os.name != "nt":
        return p.read_bytes()
    import msvcrt
    from ctypes import wintypes
    import ctypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    cf = k32.CreateFileW
    cf.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    cf.restype = wintypes.HANDLE
    handle = cf(str(p), 0x80000000, 0x7, None, 3, 0x80, None)
    if handle is None or handle == -1 or handle == wintypes.HANDLE(-1).value:
        return p.read_bytes()
    try:
        fd = msvcrt.open_osfhandle(handle, os.O_RDONLY)
        with open(fd, "rb") as f:
            return f.read()
    except Exception:
        return p.read_bytes()

def determine_data_type(unit: str, sample_val: Any = None) -> str:
    u = clean_text(unit).lower()
    if "%" in u or "percent" in u:
        return "Percentage"
    if any(k in u for k in ["mn", "k", "€", "$", "m €", "m€", "million"]):
        return "Decimal"
    if "no." in u or "number" in u or "count" in u or "unit" in u:
        return "Whole Number"
    return "Percentage" if "%" in str(sample_val or "") else "Whole Number"

def sync_workbooks(
    func_path: Path,
    strat_path: Path,
    target_path: Path,
) -> Dict[str, Any]:
    func_bytes = read_safe_bytes(func_path)
    strat_bytes = read_safe_bytes(strat_path)

    wb_func = openpyxl.load_workbook(io.BytesIO(func_bytes), data_only=True)
    wb_strat = openpyxl.load_workbook(io.BytesIO(strat_bytes), data_only=True)

    # Load existing masterfile if valid, else create new
    if target_path.exists():
        try:
            m_bytes = read_safe_bytes(target_path)
            out_wb = openpyxl.load_workbook(io.BytesIO(m_bytes))
        except Exception:
            out_wb = openpyxl.Workbook()
    else:
        out_wb = openpyxl.Workbook()

    # --- 1. Sync MPR Masterfile ---
    mpr_src_sheet = None
    if "Mastersheet" in wb_strat.sheetnames:
        mpr_src_sheet = wb_strat["Mastersheet"]
    elif "MPR Masterfile" in wb_strat.sheetnames:
        mpr_src_sheet = wb_strat["MPR Masterfile"]
    elif "AOP Critical" in wb_strat.sheetnames:
        mpr_src_sheet = wb_strat["AOP Critical"]

    if "MPR Masterfile" in out_wb.sheetnames:
        mpr_dst_sheet = out_wb["MPR Masterfile"]
    else:
        mpr_dst_sheet = out_wb.create_sheet(title="MPR Masterfile", index=0)

    if mpr_src_sheet:
        # Copy rows & styles if dst is empty or update values
        for r_idx in range(1, mpr_src_sheet.max_row + 1):
            for c_idx in range(1, mpr_src_sheet.max_column + 1):
                src_cell = mpr_src_sheet.cell(r_idx, c_idx)
                dst_cell = mpr_dst_sheet.cell(r_idx, c_idx)
                dst_cell.value = src_cell.value
                if src_cell.has_style:
                    try:
                        if src_cell.font:
                            dst_cell.font = copy.copy(src_cell.font)
                        if src_cell.fill:
                            dst_cell.fill = copy.copy(src_cell.fill)
                        if src_cell.border:
                            dst_cell.border = copy.copy(src_cell.border)
                        if src_cell.alignment:
                            dst_cell.alignment = copy.copy(src_cell.alignment)
                        if src_cell.number_format:
                            dst_cell.number_format = src_cell.number_format
                    except Exception:
                        pass

    # --- 2. Sync DMB Masterfile ---
    # Extract from all 10 function tabs
    sheet_map = [
        ("1.Quality", "Quality DMB"),
        ("2.Regulatory", "Regulatory"),
        ("3.ISC ", "ISC & Procurement"),
        ("4.Procurement", "ISC & Procurement"),
        ("5.R&D ", "R&D"),
        ("6. Marketing", "Marketing"),
        ("7.Customer Service", "Customer Service"),
        ("9.NAR ", "NAR"),
        ("10.EUROPE ", "Europe"),
        ("11.GROWTH ", "Growth"),
    ]

    all_functional_kpis = []
    for src_name, sec_title in sheet_map:
        if src_name not in wb_func.sheetnames:
            matched = [s for s in wb_func.sheetnames if norm_key(src_name) in norm_key(s) or norm_key(s) in norm_key(src_name)]
            if not matched:
                continue
            src_sheet = wb_func[matched[0]]
        else:
            src_sheet = wb_func[src_name]

        h_row = 5
        month_col_idx = {}
        for r in range(1, min(15, src_sheet.max_row + 1)):
            for c in range(1, min(30, src_sheet.max_column + 1)):
                val = clean_text(src_sheet.cell(r, c).value).lower()
                for m_idx, m_name in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]):
                    if val.startswith(m_name) and m_idx not in month_col_idx:
                        month_col_idx[m_idx] = c
            if len(month_col_idx) >= 6:
                h_row = r
                break

        current_category = "Mandatory Outcome"
        r = h_row + 1
        while r <= src_sheet.max_row:
            c1 = clean_text(src_sheet.cell(r, 1).value)
            c8 = clean_text(src_sheet.cell(r, 8).value).title()

            if "critical enabling" in c1.lower() or "leading kpi" in c1.lower():
                current_category = "Critical Enabling/Leading"
                r += 1
                continue
            if "mandatory" in c1.lower() and "kpi" in c1.lower() and c8 != "Target":
                current_category = "Mandatory Outcome"
                r += 1
                continue
            if any(stop in c1.lower() for stop in ["root cause", "action tracker", "paretos", "trends", "if the kpi"]):
                break

            if c8 == "Target" or (c1 and r + 1 <= src_sheet.max_row and clean_text(src_sheet.cell(r + 1, 8).value).title() == "Actual"):
                kpi_name = c1
                definition = clean_text(src_sheet.cell(r, 2).value)
                operator = clean_text(src_sheet.cell(r, 3).value)
                target_aop = src_sheet.cell(r, 4).value
                units = clean_text(src_sheet.cell(r, 5).value)
                metric_nature = clean_text(src_sheet.cell(r, 6).value)
                frequency = clean_text(src_sheet.cell(r, 7).value)

                t_months = [src_sheet.cell(r, month_col_idx.get(m_i, m_i + 9)).value for m_i in range(12)]
                a_months = [None] * 12
                if r + 1 <= src_sheet.max_row and clean_text(src_sheet.cell(r + 1, 8).value).title() == "Actual":
                    a_months = [src_sheet.cell(r + 1, month_col_idx.get(m_i, m_i + 9)).value for m_i in range(12)]
                    next_step = 2
                else:
                    next_step = 1

                display_kpi_name = kpi_name
                if "procurement" in src_name.lower() and "ctb" in kpi_name.lower():
                    if "(procurement)" not in display_kpi_name.lower():
                        display_kpi_name = f"{display_kpi_name} (Procurement)"
                elif "isc" in src_name.lower() and "ctb" in kpi_name.lower():
                    if "(isc)" not in display_kpi_name.lower():
                        display_kpi_name = f"{display_kpi_name} (ISC)"

                data_type = determine_data_type(units, target_aop)

                all_functional_kpis.append({
                    "section": sec_title,
                    "kpi_name": display_kpi_name,
                    "definition": definition,
                    "operator": operator,
                    "target_aop": target_aop,
                    "units": units,
                    "metric_nature": metric_nature,
                    "frequency": frequency,
                    "target_months": t_months,
                    "actual_months": a_months,
                    "data_type": data_type,
                    "category": current_category,
                    "source_sheet": src_name,
                })
                r += next_step
            else:
                r += 1

    # Check if existing DMB Masterfile has 200 rows and matches
    if "DMB Masterfile" in out_wb.sheetnames:
        dmb_dst_sheet = out_wb["DMB Masterfile"]
    else:
        dmb_dst_sheet = out_wb.create_sheet(title="DMB Masterfile", index=1)

    # If dst sheet is valid, update its cells by matching section and KPI name
    if dmb_dst_sheet.max_row >= 180:
        logger.info("Updating existing DMB Masterfile (%d rows) with latest function data...", dmb_dst_sheet.max_row)
        current_section = ""
        for r in range(1, dmb_dst_sheet.max_row + 1):
            c0 = clean_text(dmb_dst_sheet.cell(r, 1).value)
            c7 = clean_text(dmb_dst_sheet.cell(r, 8).value).title()
            if not c7 and c0 and c0 != "KPI Name" and c0 != "Mastersheet":
                current_section = c0
                continue

            if c7 in ["Target", "Actual"] and c0:
                # Find matching extracted KPI
                matched = [
                    k for k in all_functional_kpis
                    if (k["section"] == current_section or norm_key(k["section"]) == norm_key(current_section))
                    and (norm_key(k["kpi_name"]) == norm_key(c0) or norm_key(c0) in norm_key(k["kpi_name"]))
                ]
                if matched:
                    src_kpi = matched[0]
                    # Update definition, operator, target AOP, units, nature, freq
                    if src_kpi["definition"]:
                        dmb_dst_sheet.cell(r, 2, src_kpi["definition"])
                    if src_kpi["operator"]:
                        dmb_dst_sheet.cell(r, 3, src_kpi["operator"])
                    if src_kpi["target_aop"] is not None:
                        dmb_dst_sheet.cell(r, 4, src_kpi["target_aop"])
                    if src_kpi["units"]:
                        dmb_dst_sheet.cell(r, 5, src_kpi["units"])
                    if src_kpi["metric_nature"]:
                        dmb_dst_sheet.cell(r, 6, src_kpi["metric_nature"])
                    if src_kpi["frequency"]:
                        dmb_dst_sheet.cell(r, 7, src_kpi["frequency"])

                    # Update monthly values (col 9 to 20)
                    m_vals = src_kpi["target_months"] if c7 == "Target" else src_kpi["actual_months"]
                    for m_i in range(12):
                        val = m_vals[m_i]
                        if val is not None and str(val).strip() != "":
                            dmb_dst_sheet.cell(r, m_i + 9, val)

                    # Update data type & category
                    if src_kpi["data_type"]:
                        dmb_dst_sheet.cell(r, 21, src_kpi["data_type"])
                    if src_kpi["category"]:
                        dmb_dst_sheet.cell(r, 22, src_kpi["category"])
    else:
        # Build fresh DMB Masterfile from scratch
        logger.info("Building fresh DMB Masterfile from %d extracted KPIs...", len(all_functional_kpis))
        dmb_dst_sheet.cell(1, 1, "Mastersheet")
        cur_r = 3
        # Group KPIs by section in exact canonical order
        sec_order = [
            "Quality DMB",
            "Regulatory",
            "ISC & Procurement",
            "R&D",
            "Marketing",
            "Customer Service",
            "NAR",
            "Europe",
            "Growth",
        ]
        for sec in sec_order:
            dmb_dst_sheet.cell(cur_r, 1, sec)
            cur_r += 1
            # Header row
            headers = [
                "KPI Name", "KPI  Definition", "Operators", "Target AOP 2026", "Units",
                "Metric nature", "Frequency", "Target/ Actual",
                "Jan-2026", "Feb-2026", "Mar-2026", "Apr-2026", "May-2026", "Jun-2026",
                "Jul-2026", "Aug-2026", "Sep-2026", "Oct-2026", "Nov-2026", "Dec-2026",
                "Data Type", "KPI category",
            ]
            for col_i, h in enumerate(headers, 1):
                dmb_dst_sheet.cell(cur_r, col_i, h)
            cur_r += 1

            sec_kpis = [k for k in all_functional_kpis if k["section"] == sec]
            for k in sec_kpis:
                # Target row
                dmb_dst_sheet.cell(cur_r, 1, k["kpi_name"])
                dmb_dst_sheet.cell(cur_r, 2, k["definition"])
                dmb_dst_sheet.cell(cur_r, 3, k["operator"])
                dmb_dst_sheet.cell(cur_r, 4, k["target_aop"])
                dmb_dst_sheet.cell(cur_r, 5, k["units"])
                dmb_dst_sheet.cell(cur_r, 6, k["metric_nature"])
                dmb_dst_sheet.cell(cur_r, 7, k["frequency"])
                dmb_dst_sheet.cell(cur_r, 8, "Target")
                for m_i in range(12):
                    dmb_dst_sheet.cell(cur_r, m_i + 9, k["target_months"][m_i])
                dmb_dst_sheet.cell(cur_r, 21, k["data_type"])
                dmb_dst_sheet.cell(cur_r, 22, k["category"])
                cur_r += 1

                # Actual row
                dmb_dst_sheet.cell(cur_r, 1, k["kpi_name"])
                dmb_dst_sheet.cell(cur_r, 2, k["definition"])
                dmb_dst_sheet.cell(cur_r, 3, k["operator"])
                dmb_dst_sheet.cell(cur_r, 4, k["target_aop"])
                dmb_dst_sheet.cell(cur_r, 5, k["units"])
                dmb_dst_sheet.cell(cur_r, 6, k["metric_nature"])
                dmb_dst_sheet.cell(cur_r, 7, k["frequency"])
                dmb_dst_sheet.cell(cur_r, 8, "Actual")
                for m_i in range(12):
                    dmb_dst_sheet.cell(cur_r, m_i + 9, k["actual_months"][m_i])
                dmb_dst_sheet.cell(cur_r, 21, k["data_type"])
                dmb_dst_sheet.cell(cur_r, 22, k["category"])
                cur_r += 1

    # Remove extra sheet if default exists
    for s in list(out_wb.sheetnames):
        if s not in ["MPR Masterfile", "DMB Masterfile"]:
            out_wb.remove(out_wb[s])

    # Validation
    dmb_rows = dmb_dst_sheet.max_row
    mpr_rows = mpr_dst_sheet.max_row
    if dmb_rows < 180 or mpr_rows < 45:
        raise ValueError(f"Sync validation failed! DMB rows={dmb_rows} (expected >=180), MPR rows={mpr_rows} (expected >=45).")

    buf = io.BytesIO()
    out_wb.save(buf)
    content = buf.getvalue()

    return {
        "success": True,
        "dmb_rows": dmb_rows,
        "mpr_rows": mpr_rows,
        "content": content,
        "message": f"Successfully consolidated Masterfile ({dmb_rows} DMB rows, {mpr_rows} MPR rows).",
    }

def test_sync_and_validate():
    func_file = DATA_DIR / "Functional DMB Review Sheets.xlsx"
    strat_file = DATA_DIR / "Strategic Execution Dashboard.xlsx"
    master_file = DATA_DIR / "Masterfile_DMB_Dashboard.xlsx"

    res = sync_workbooks(func_file, strat_file, master_file)
    print(res["message"])
    assert res["success"]
    assert res["dmb_rows"] >= 180
    assert res["mpr_rows"] >= 45
    print("Test passed successfully!")

if __name__ == "__main__":
    test_sync_and_validate()
