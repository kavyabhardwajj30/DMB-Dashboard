"""
DMB-Dashboard Poller Integration
================================
Include this in your DMB-Dashboard app.py to enable automatic
background polling of the mailbox and hot-reloading data.
"""

import threading
import time
import email_sync

def start_dmb_mailbox_poller(reload_callback=None):
    """
    Starts a background daemon thread that polls the mailbox every 60 seconds.
    Whenever a new Excel file arrives:
      1. email_sync downloads it into data/ (Functional DMB Review Sheets / Masterfile)
      2. email_sync commits it to GitHub (kavyabhardwajj30/DMB-Dashboard)
      3. reload_callback() is called to refresh KPI calculations / data in memory.
    """
    def _worker():
        # Initial check at app startup
        try:
            initial_count = email_sync.sync_once()
            if initial_count > 0 and reload_callback:
                reload_callback()
        except Exception as e:
            print(f"[DMB Poller Startup Notice] {e}")

        # Continuous background polling
        while True:
            time.sleep(60)
            try:
                updated_count = email_sync.sync_once()
                if updated_count > 0:
                    print(f"[DMB] {updated_count} new workbook(s) synced! Triggering data reload...")
                    if reload_callback:
                        reload_callback()
            except Exception as e:
                print(f"[DMB Poller Error] {e}")

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
    print("[DMB] 24/7 Mailbox poller successfully started.")
