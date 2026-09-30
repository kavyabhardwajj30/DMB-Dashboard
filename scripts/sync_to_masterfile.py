"""
DMB Masterfile Synchronizer & Consolidator (Python)
===================================================
Synchronizes updated data from:
  1. Functional DMB Review Sheets (e.g. Functional DMB Review Sheets-17th_sept.xlsx)
  2. Strategic Execution Dashboard (e.g. Strategic Execution Dashboard-17Th_sept.xlsx)
into the canonical target:
  -> Masterfile_DMB_Dashboard.xlsx (containing 'DMB Masterfile' & 'MPR Masterfile' sheets)

Modes:
  1. One-shot Execution:
       python scripts/sync_to_masterfile.py
  2. Live Watcher Daemon (Auto-sync on save):
       python scripts/sync_to_masterfile.py --watch
"""

from __future__ import annotations

import argparse
import ctypes
import hashlib
import io
import logging
import os
from pathlib import Path
import shutil
import sys
import time
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="[Masterfile Sync %(asctime)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("dmb.sync_masterfile")

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

TARGET_MASTERFILE = DATA_DIR / "Masterfile_DMB_Dashboard.xlsx"


def read_file_safe_bytes(file_path: Path | str) -> bytes:
    """Read file bytes safely on Windows even if open in Excel."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if os.name != "nt":
        return file_path.read_bytes()

    import msvcrt
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE

    handle = create_file(
        str(file_path),
        0x80000000,  # GENERIC_READ
        0x7,         # FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
        None,
        3,           # OPEN_EXISTING
        0x80,        # FILE_ATTRIBUTE_NORMAL
        None,
    )
    if handle is None or handle == wintypes.HANDLE(-1).value or handle == -1:
        return file_path.read_bytes()

    try:
        fd = msvcrt.open_osfhandle(handle, os.O_RDONLY)
        with open(fd, "rb") as f:
            return f.read()
    except Exception:
        return file_path.read_bytes()


def atomic_write_bytes(target_path: Path, content: bytes) -> bool:
    """Write bytes atomically to target path to prevent file corruption."""
    temp_path = target_path.with_suffix(f".tmp.{os.getpid()}_{int(time.time()*1000)}")
    try:
        temp_path.write_bytes(content)
        if target_path.exists():
            try:
                target_path.unlink()
            except Exception:
                pass
        temp_path.replace(target_path)
        return True
    except Exception as e:
        logger.error("Atomic write failed for %s: %s", target_path, e)
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        return False


def get_file_hash(path: Path) -> str:
    """Return SHA256 hex digest of file bytes."""
    try:
        data = read_file_safe_bytes(path)
        return hashlib.sha256(data).hexdigest()
    except Exception:
        return ""


def find_source_files(search_dir: Path = DATA_DIR) -> Tuple[Optional[Path], Optional[Path]]:
    """Locate the most recent Functional DMB Review workbook and Strategic Execution workbook."""
    candidates = [p for p in search_dir.glob("*.xlsx") if not p.name.startswith("~$")]

    func_file = None
    strat_file = None

    func_matches = [
        p for p in candidates
        if ("functional" in p.name.lower() or "review" in p.name.lower())
        and "master" not in p.name.lower()
    ]
    if func_matches:
        func_file = max(func_matches, key=lambda p: p.stat().st_mtime)

    strat_matches = [
        p for p in candidates
        if ("strategic" in p.name.lower() or "aop" in p.name.lower())
        and "master" not in p.name.lower()
    ]
    if strat_matches:
        strat_file = max(strat_matches, key=lambda p: p.stat().st_mtime)

    return func_file, strat_file


from copy import copy


def copy_sheet_content(src_ws: openpyxl.worksheet.worksheet.Worksheet, dst_ws: openpyxl.worksheet.worksheet.Worksheet):
    """Copy all cell values, formulas, and exact styles (fonts, fills, borders, alignments, dimensions) from src_ws to dst_ws."""
    for row in src_ws.iter_rows():
        for cell in row:
            dst_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                try:
                    if cell.font:
                        dst_cell.font = copy(cell.font)
                    if cell.fill:
                        dst_cell.fill = copy(cell.fill)
                    if cell.border:
                        dst_cell.border = copy(cell.border)
                    if cell.alignment:
                        dst_cell.alignment = copy(cell.alignment)
                    if cell.number_format:
                        dst_cell.number_format = cell.number_format
                    if cell.protection:
                        dst_cell.protection = copy(cell.protection)
                except Exception:
                    pass

    # Copy merged cells
    for merged_cell in src_ws.merged_cells.ranges:
        try:
            dst_ws.merge_cells(str(merged_cell))
        except Exception:
            pass

    # Copy row heights
    for row_idx, row_dim in src_ws.row_dimensions.items():
        if row_dim.height is not None:
            dst_ws.row_dimensions[row_idx].height = row_dim.height

    # Copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            dst_ws.column_dimensions[col_letter].width = dim.width
        if dim.hidden:
            dst_ws.column_dimensions[col_letter].hidden = True

    # Ensure grid lines are visible
    try:
        if dst_ws.views.sheetView:
            dst_ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass


def sync_masterfile(
    func_path: Optional[Path] = None,
    strat_path: Optional[Path] = None,
    target_path: Path = TARGET_MASTERFILE,
) -> dict:
    """
    Syncs the Functional DMB Review file and Strategic Execution file into Masterfile_DMB_Dashboard.xlsx.
    """
    if func_path is None or strat_path is None:
        detected_func, detected_strat = find_source_files()
        func_path = func_path or detected_func
        strat_path = strat_path or detected_strat

    if not func_path or not func_path.exists():
        return {"success": False, "message": f"Functional DMB file not found: {func_path}"}
    if not strat_path or not strat_path.exists():
        return {"success": False, "message": f"Strategic Execution file not found: {strat_path}"}

    logger.info("Reading Functional DMB Workbook: %s", func_path.name)
    logger.info("Reading Strategic Execution Workbook: %s", strat_path.name)

    # 1. Load Functional Workbook
    func_bytes = read_file_safe_bytes(func_path)
    func_wb = openpyxl.load_workbook(io.BytesIO(func_bytes), data_only=True)

    # 2. Load Strategic Workbook
    strat_bytes = read_file_safe_bytes(strat_path)
    strat_wb = openpyxl.load_workbook(io.BytesIO(strat_bytes), data_only=True)

    # 3. Create or Load Master Workbook
    out_wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = out_wb.active

    # --- Sync MPR Masterfile ---
    mpr_src_sheet = None
    if "Mastersheet" in strat_wb.sheetnames:
        mpr_src_sheet = strat_wb["Mastersheet"]
    elif "MPR Masterfile" in strat_wb.sheetnames:
        mpr_src_sheet = strat_wb["MPR Masterfile"]
    elif "AOP Critical" in strat_wb.sheetnames:
        mpr_src_sheet = strat_wb["AOP Critical"]
    else:
        mpr_src_sheet = strat_wb.active

    mpr_dst_sheet = out_wb.create_sheet(title="MPR Masterfile")
    if mpr_src_sheet:
        copy_sheet_content(mpr_src_sheet, mpr_dst_sheet)
        logger.info("Synced 'MPR Masterfile' from '%s' (%d rows).", mpr_src_sheet.title, mpr_src_sheet.max_row)

    # --- Sync DMB Masterfile ---
    dmb_src_sheet = None
    if "MasterSheet" in func_wb.sheetnames:
        dmb_src_sheet = func_wb["MasterSheet"]
    elif "Mastersheet" in func_wb.sheetnames:
        dmb_src_sheet = func_wb["Mastersheet"]
    elif "DMB Masterfile" in func_wb.sheetnames:
        dmb_src_sheet = func_wb["DMB Masterfile"]
    else:
        dmb_src_sheet = func_wb.active

    dmb_dst_sheet = out_wb.create_sheet(title="DMB Masterfile")
    if dmb_src_sheet:
        copy_sheet_content(dmb_src_sheet, dmb_dst_sheet)
        logger.info("Synced 'DMB Masterfile' from '%s' (%d rows).", dmb_src_sheet.title, dmb_src_sheet.max_row)

    # Remove the blank initial sheet
    if default_sheet and default_sheet in out_wb.worksheets:
        out_wb.remove(default_sheet)

    # Save to memory and write atomically
    out_buf = io.BytesIO()
    out_wb.save(out_buf)
    content_bytes = out_buf.getvalue()

    success = atomic_write_bytes(target_path, content_bytes)
    if success:
        logger.info("Successfully updated '%s' (%d bytes).", target_path.name, len(content_bytes))
        return {
            "success": True,
            "target": str(target_path),
            "func_rows": dmb_src_sheet.max_row if dmb_src_sheet else 0,
            "mpr_rows": mpr_src_sheet.max_row if mpr_src_sheet else 0,
            "message": f"Masterfile updated successfully with {dmb_src_sheet.max_row if dmb_src_sheet else 0} DMB rows and {mpr_src_sheet.max_row if mpr_src_sheet else 0} MPR rows.",
        }
    else:
        return {"success": False, "message": "Failed to write target masterfile."}


def watch_and_sync(interval_seconds: float = 3.0):
    """
    Continuous background watcher that detects changes in Functional and Strategic files
    and immediately synchronizes them into Masterfile_DMB_Dashboard.xlsx.
    """
    logger.info("Starting Masterfile Live Watcher (checking every %.1f seconds)...", interval_seconds)
    last_hashes: Dict[str, str] = {}

    # Initial sync
    try:
        res = sync_masterfile()
        logger.info("Initial sync status: %s", res.get("message"))
    except Exception as e:
        logger.error("Initial sync error: %s", e)

    func_file, strat_file = find_source_files()
    if func_file:
        last_hashes[str(func_file)] = get_file_hash(func_file)
    if strat_file:
        last_hashes[str(strat_file)] = get_file_hash(strat_file)

    while True:
        try:
            time.sleep(interval_seconds)
            func_file, strat_file = find_source_files()
            changed = False

            for f in [func_file, strat_file]:
                if not f or not f.exists():
                    continue
                current_hash = get_file_hash(f)
                prev_hash = last_hashes.get(str(f))
                if current_hash != prev_hash:
                    logger.info("Change detected in '%s'! Triggering sync...", f.name)
                    last_hashes[str(f)] = current_hash
                    changed = True

            if changed:
                res = sync_masterfile(func_file, strat_file)
                logger.info("Sync completed: %s", res.get("message"))

        except KeyboardInterrupt:
            logger.info("Masterfile watcher stopped by user.")
            break
        except Exception as e:
            logger.error("Watcher loop error: %s", e)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Synchronize Functional & Strategic Excel files into Masterfile.")
    parser.add_argument("--watch", action="store_true", help="Run in continuous watch mode (auto-sync on file change).")
    parser.add_argument("--interval", type=float, default=3.0, help="Watcher poll interval in seconds.")
    args = parser.parse_args()

    if args.watch:
        watch_and_sync(args.interval)
    else:
        result = sync_masterfile()
        print(result)
