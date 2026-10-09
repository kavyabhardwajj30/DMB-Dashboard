import unittest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from app import populate_continuous_red_modal, populate_function_gauge_modal

class TestDualModals(unittest.TestCase):
    def test_rca_and_gauge_modals_independent(self):
        # 1. Populate RCA modal for Customer Service
        rca_state = {"is_open": True, "function": "Customer Service", "month": "2026-07-01"}
        c_title, c_month, c_body = populate_continuous_red_modal(
            active_state=rca_state,
            month_filter_val="2026-07-01",
            sync_data=None,
        )
        self.assertIn("Customer Service", c_title)
        self.assertIsNotNone(c_body)
        print("Successfully tested RCA modal opening for Customer Service.")

        # 2. Populate Gauge modal for R&D
        gauge_state = {"is_open": True, "function": "R&D", "month": "2026-07-01"}
        g_title, g_subtitle, g_body = populate_function_gauge_modal(
            active_state=gauge_state,
            month_filter_val="2026-07-01",
            sync_data=None,
        )
        self.assertIn("R&D", g_title)
        self.assertIsNotNone(g_body)
        print("Successfully tested Gauge KPI & Trend modal opening for R&D.")

if __name__ == "__main__":
    unittest.main()
