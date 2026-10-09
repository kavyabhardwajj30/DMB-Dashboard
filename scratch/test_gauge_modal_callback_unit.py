import unittest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from unittest.mock import patch, MagicMock
import dash
from dash.exceptions import PreventUpdate
import pandas as pd
from app import manage_function_gauge_modal

class TestManageFunctionGaugeModal(unittest.TestCase):
    def test_open_gauge_modal(self):
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": '{"function":"Quality","type":"function-gauge-card"}.n_clicks', "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "function-gauge-card", "function": "Quality"}
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            res = manage_function_gauge_modal(
                gauge_card_clicks=[1, 0, 0, 0, 0, 0],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 100},
                current_modal_state={"is_open": False, "function": None, "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            g_class, g_title, g_subtitle, g_body, state = res
            print("Opened Quality Gauge Modal:", g_title, state)
            self.assertEqual(g_class, "function-gauge-modal")
            self.assertIn("Quality", g_title)
            self.assertEqual(state["function"], "Quality")
            self.assertTrue(state["is_open"])

    def test_close_gauge_modal(self):
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": "close-function-gauge-modal.n_clicks", "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "close-function-gauge-modal"
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            res = manage_function_gauge_modal(
                gauge_card_clicks=[0],
                close_clicks=1,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 456},
                current_modal_state={"is_open": True, "function": "Quality", "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            g_class, g_title, g_subtitle, g_body, state = res
            print("Close Gauge Modal:", g_class, state)
            self.assertIn("function-gauge-modal-hidden", g_class)
            self.assertFalse(state["is_open"])

if __name__ == "__main__":
    unittest.main()
