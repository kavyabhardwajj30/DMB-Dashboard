import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
import pandas as pd
import sharepoint_sync
import data_loader
import app

def test_integration():
    print("--- 1. Testing Sync Engine ---")
    sync_mgr = sharepoint_sync.get_sync_manager()
    status = sync_mgr.get_status()
    print(f"Status: {status['status']}")
    print(f"Local files detected: {status['local_files_detected']}")
    
    sync_res = sharepoint_sync.sync_now()
    print(f"Sync Now Result: {sync_res}")
    
    print("\n--- 2. Testing Data Store Dynamic Reload ---")
    store = data_loader.get_data_store()
    reloaded = store.reload(force=True)
    print(f"Data Store Reloaded: {reloaded}")
    mpr_df, dmb_df, strat_df = store.get_data()
    print(f"MPR Data Shape: {mpr_df.shape}")
    print(f"DMB Data Shape: {dmb_df.shape}")
    print(f"Strategic RCA Shape: {strat_df.shape}")
    
    print("\n--- 3. Testing Dash Callbacks ---")
    # Live sync callback
    sync_store_data = app.handle_live_sync_trigger(1)
    print(f"Sync Callback Output: data={sync_store_data}")
    
    # Month callbacks
    default_m = "2026-08-01"
    mpr_res = app.update_mpr_dashboard(default_m, default_m, sync_store_data)
    print(f"MPR Callback Title: {mpr_res[0]}, KPIs total: {mpr_res[1]}")
    
    dmb_res = app.update_dmb_function_cards(default_m, default_m, sync_store_data)
    print(f"DMB Callback Functions count: {len(dmb_res)}")
    
    rca_res = app.update_rca_table(default_m, default_m, sync_store_data)
    print(f"RCA Table Month: {rca_res[1]}")
    
    # Test Function KPIs 2-row Table Detail
    fn_detail = app.create_function_kpis_table_detail("Quality", default_m)
    print(f"Function Detail Rendered: type={type(fn_detail)}")
    
    print("\n>>> ALL TESTS PASSED SUCCESSFULLY! <<<")

if __name__ == "__main__":
    test_integration()
