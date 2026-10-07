import unittest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from unittest.mock import patch, MagicMock
import json
import dash
from dash.exceptions import PreventUpdate
import pandas as pd
from app import manage_continuous_red_modal, default_month

class TestManageContinuousRedModal(unittest.TestCase):
    def test_single_click_card_dash_ctx(self):
        # Scenario 1: dash.ctx.triggered_id is a dict with {"type": "continuous-red-card", "function": "ISC & Procurement"}
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": '{"function":"ISC & Procurement","type":"continuous-red-card"}.n_clicks', "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "continuous-red-card", "function": "ISC & Procurement"}
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            res = manage_continuous_red_modal(
                card_clicks=[1],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 123},
                current_modal_state={"is_open": False, "function": None, "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            c_class, c_title, c_month, c_body, state = res
            print("Scenario 1 Result Class:", c_class)
            print("Scenario 1 Title:", c_title)
            print("Scenario 1 State:", state)
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertTrue("ISC & Procurement" in c_title)
            self.assertTrue(state["is_open"])
            self.assertEqual(state["function"], "ISC & Procurement")

    def test_multi_trigger_with_sync(self):
        # Scenario 2: live sync state changed at the same time as card click
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": "live-sync-state-store.data", "value": {"ts": 456}},
            {"prop_id": '{"function":"Regulatory","type":"continuous-red-card"}.n_clicks', "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = None
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            res = manage_continuous_red_modal(
                card_clicks=[1],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 456},
                current_modal_state={"is_open": False, "function": None, "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            c_class, c_title, c_month, c_body, state = res
            print("Scenario 2 Result:", c_title, state)
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertTrue("Regulatory" in c_title)
            self.assertTrue(state["is_open"])

    def test_close_modal(self):
        # Scenario 3: user clicks close button
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": "close-continuous-red-modal.n_clicks", "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "close-continuous-red-modal"
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            res = manage_continuous_red_modal(
                card_clicks=[0],
                close_clicks=1,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 456},
                current_modal_state={"is_open": True, "function": "Regulatory", "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            c_class, c_title, c_month, c_body, state = res
            print("Scenario 3 Close Result:", c_class, state)
            self.assertIn("continuous-red-modal-hidden", c_class)
            self.assertFalse(state["is_open"])

    def test_background_sync_while_open(self):
        # Scenario 4: live sync fires while modal is OPEN -> refreshes without closing
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": "live-sync-state-store.data", "value": {"ts": 789}}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "live-sync-state-store"
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            res = manage_continuous_red_modal(
                card_clicks=[0],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 789},
                current_modal_state={"is_open": True, "function": "ISC & Procurement", "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            c_class, c_title, c_month, c_body, state = res
            print("Scenario 4 Open Refresh Result:", c_title, state)
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertTrue(state["is_open"])

    def test_background_sync_while_closed(self):
        # Scenario 5: live sync fires while modal is CLOSED -> PreventUpdate
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": "live-sync-state-store.data", "value": {"ts": 789}}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "live-sync-state-store"
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            with self.assertRaises(PreventUpdate):
                manage_continuous_red_modal(
                    card_clicks=[0],
                    close_clicks=0,
                    backdrop_clicks=0,
                    month_filter_val="2026-07-01",
                    sync_data={"ts": 789},
                    current_modal_state={"is_open": False, "function": None, "month": "2026-07-01"},
                    month_state="2026-07-01",
                )
            print("Scenario 5 Successfully Prevented Update when modal was closed!")

if __name__ == "__main__":
    unittest.main()
