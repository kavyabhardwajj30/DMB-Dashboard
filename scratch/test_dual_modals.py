import unittest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from app import get_continuous_red_modal_content, get_function_kpis_trend_modal_content

class TestDualModals(unittest.TestCase):
    def test_rca_and_gauge_modals_independent(self):
        # 1. Populate RCA modal for Customer Service
        c_class, c_title, c_month, c_body = get_continuous_red_modal_content(
            "Customer Service", "2026-08-01"
        )
        self.assertEqual(c_class, "continuous-red-modal")
        self.assertIn("Customer Service", c_title)
        self.assertIn("August 2026", c_month)
        self.assertIsNotNone(c_body)
        print("Successfully tested RCA modal content generation for Customer Service.")

        # 2. Populate Gauge modal for R&D
        g_body = get_function_kpis_trend_modal_content(
            "R&D", "2026-08-01"
        )
        self.assertIsNotNone(g_body)
        print("Successfully tested Gauge KPI & Trend modal content generation for R&D.")

if __name__ == "__main__":
    unittest.main()
