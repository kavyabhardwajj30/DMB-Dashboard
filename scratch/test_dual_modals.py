import unittest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from unittest.mock import patch, MagicMock
import json
import dash
from dash.exceptions import PreventUpdate
import pandas as pd
from app import manage_continuous_red_modal, manage_function_gauge_modal

class TestDualModals(unittest.TestCase):
    def test_rca_and_gauge_modals_independent(self):
        # 1. Click RCA card for Customer Service
        mock_ctx_rca = MagicMock()
        mock_ctx_rca.triggered = [
            {"prop_id": '{"function":"Customer Service","type":"continuous-red-card"}.n_clicks', "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx_rca), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "continuous-red-card", "function": "Customer Service"}
            mock_dash_ctx.triggered = mock_ctx_rca.triggered
            
            rca_res = manage_continuous_red_modal(
                card_clicks=[0, 0, 0, 0, 1, 0],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data=None,
                current_modal_state={"is_open": False, "function": None, "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            c_class, c_title, c_month, c_body, state = rca_res
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertIn("Customer Service", c_title)
            self.assertTrue(state["is_open"])
            self.assertEqual(state["function"], "Customer Service")
            print("Successfully tested RCA modal opening for Customer Service.")

        # 2. Click Gauge card for R&D
        mock_ctx_gauge = MagicMock()
        mock_ctx_gauge.triggered = [
            {"prop_id": '{"function":"R&D","type":"function-gauge-card"}.n_clicks', "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx_gauge), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "function-gauge-card", "function": "R&D"}
            mock_dash_ctx.triggered = mock_ctx_gauge.triggered
            
            g_res = manage_function_gauge_modal(
                gauge_card_clicks=[0, 0, 0, 1, 0, 0],
                kpi_tab_clicks=[],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data=None,
                current_modal_state={"is_open": False, "function": None, "selected_kpi": None, "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            g_class, g_title, g_subtitle, g_body, g_state = g_res
            self.assertEqual(g_class, "function-gauge-modal")
            self.assertIn("R&D", g_subtitle)
            self.assertTrue(g_state["is_open"])
            self.assertEqual(g_state["function"], "R&D")
            print("Successfully tested Gauge KPI & Trend modal opening for R&D.")

if __name__ == "__main__":
    unittest.main()
