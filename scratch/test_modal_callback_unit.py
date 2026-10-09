import unittest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from unittest.mock import patch, MagicMock
from dash.exceptions import PreventUpdate
import pandas as pd
from app import populate_continuous_red_modal, populate_function_gauge_modal, default_month

class TestModalCallbacks(unittest.TestCase):
    def test_populate_rca_modal_when_open(self):
        active_state = {"is_open": True, "function": "Marketing", "month": "2026-07-01"}
        c_title, c_month, c_body = populate_continuous_red_modal(
            active_state=active_state,
            month_filter_val="2026-07-01",
            sync_data={"ts": 100},
        )
        print("Populated Marketing RCA Modal:", c_title, c_month)
        self.assertIn("Marketing", c_title)
        self.assertIn("July 2026", c_month)
        self.assertIsNotNone(c_body)

    def test_rca_modal_persists_during_sync(self):
        # When background sync happens and modal is open, content updates without closing
        active_state = {"is_open": True, "function": "Marketing", "month": "2026-07-01"}
        c_title, c_month, c_body = populate_continuous_red_modal(
            active_state=active_state,
            month_filter_val="2026-07-01",
            sync_data={"ts": 200, "auto": True},
        )
        self.assertIn("Marketing", c_title)
        self.assertIsNotNone(c_body)

    def test_rca_modal_closed_prevents_update(self):
        # When modal is closed, background sync does not touch the modal
        active_state = {"is_open": False, "function": None, "month": "2026-07-01"}
        with self.assertRaises(PreventUpdate):
            populate_continuous_red_modal(
                active_state=active_state,
                month_filter_val="2026-07-01",
                sync_data={"ts": 300},
            )

    def test_populate_gauge_modal_when_open(self):
        active_state = {"is_open": True, "function": "Quality", "month": "2026-07-01"}
        g_title, g_subtitle, g_body = populate_function_gauge_modal(
            active_state=active_state,
            month_filter_val="2026-07-01",
            sync_data={"ts": 100},
        )
        print("Populated Quality Gauge Modal:", g_title, g_subtitle)
        self.assertIn("Quality", g_title)
        self.assertIn("July 2026", g_subtitle)
        self.assertIsNotNone(g_body)

    def test_gauge_modal_closed_prevents_update(self):
        active_state = {"is_open": False, "function": None, "month": "2026-07-01"}
        with self.assertRaises(PreventUpdate):
            populate_function_gauge_modal(
                active_state=active_state,
                month_filter_val="2026-07-01",
                sync_data={"ts": 300},
            )

if __name__ == "__main__":
    unittest.main()
