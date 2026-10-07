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
    def test_open_marketing_and_background_sync(self):
        # Scenario 1: User clicks Marketing card
        mock_ctx = MagicMock()
        mock_ctx.triggered = [
            {"prop_id": '{"function":"Marketing","type":"continuous-red-card"}.n_clicks', "value": 1}
        ]
        with patch("dash.callback_context", mock_ctx), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = {"type": "continuous-red-card", "function": "Marketing"}
            mock_dash_ctx.triggered = mock_ctx.triggered
            
            res = manage_continuous_red_modal(
                card_clicks=[0, 0, 0, 0, 0, 1],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 100},
                current_modal_state={"is_open": False, "function": None, "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            c_class, c_title, c_month, c_body, state = res
            print("Opened Marketing Modal:", c_title, state)
            self.assertEqual(c_class, "continuous-red-modal")
            self.assertIn("Marketing", c_title)
            self.assertEqual(state["function"], "Marketing")
            self.assertTrue(state["is_open"])

        # Scenario 2: 20s Background Sync fires while Marketing is OPEN -> MUST STAY ON MARKETING!
        mock_ctx_sync = MagicMock()
        mock_ctx_sync.triggered = [
            {"prop_id": "live-sync-state-store.data", "value": {"ts": 120}},
            {"prop_id": '{"function":"Quality","type":"continuous-red-card"}.n_clicks', "value": 0},
            {"prop_id": '{"function":"Regulatory","type":"continuous-red-card"}.n_clicks', "value": 0},
            {"prop_id": '{"function":"ISC & Procurement","type":"continuous-red-card"}.n_clicks', "value": 0},
            {"prop_id": '{"function":"R&D","type":"continuous-red-card"}.n_clicks', "value": 0},
            {"prop_id": '{"function":"Customer Service","type":"continuous-red-card"}.n_clicks', "value": 0},
            {"prop_id": '{"function":"Marketing","type":"continuous-red-card"}.n_clicks', "value": 1},
        ]
        with patch("dash.callback_context", mock_ctx_sync), patch("dash.ctx") as mock_dash_ctx:
            mock_dash_ctx.triggered_id = "live-sync-state-store"
            mock_dash_ctx.triggered = mock_ctx_sync.triggered
            
            res_sync = manage_continuous_red_modal(
                card_clicks=[0, 0, 0, 0, 0, 1],
                close_clicks=0,
                backdrop_clicks=0,
                month_filter_val="2026-07-01",
                sync_data={"ts": 120},
                current_modal_state=state,  # currently open with Marketing
                month_state="2026-07-01",
            )
            s_class, s_title, s_month, s_body, s_state = res_sync
            print("After Background Sync with Marketing Open:", s_title, s_state)
            self.assertEqual(s_class, "continuous-red-modal")
            self.assertIn("Marketing", s_title)
            self.assertNotIn("Quality", s_title)
            self.assertEqual(s_state["function"], "Marketing")
            self.assertTrue(s_state["is_open"])

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
                current_modal_state={"is_open": True, "function": "Marketing", "month": "2026-07-01"},
                month_state="2026-07-01",
            )
            c_class, c_title, c_month, c_body, state = res
            print("Close Result:", c_class, state)
            self.assertIn("continuous-red-modal-hidden", c_class)
            self.assertFalse(state["is_open"])

    def test_background_sync_while_closed(self):
        # Scenario 4: live sync fires while modal is CLOSED -> PreventUpdate
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
            print("Scenario 4 Successfully Prevented Update when modal was closed!")

if __name__ == "__main__":
    unittest.main()
