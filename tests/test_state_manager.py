import os
import unittest
import tempfile
import json
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from state_manager import StateManager


class TestStateManager(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.state_file = os.path.join(self.temp_dir.name, "sub_dir", "app_state.json")
        self.mgr = StateManager(self.state_file)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_load_empty_on_missing_file(self):
        self.assertEqual(self.mgr.load_state(), {})

    def test_save_and_load_state(self):
        test_data = {
            "target_month": "SEPTEMBER 2026",
            "invoice_date": "10-09-2026",
            "starting_inv_no": 1800,
            "agency_inv_map": {"MAHALASA": 1805},
            "agency_meta_overrides": {
                "MAHALASA": {
                    "agency_name": "MAHALASA GAS ENTERPRISES",
                    "pan_no": "ABCDE1234F",
                    "vendor_code": "9999",
                    "gst_no": "29ABCDE1234F1Z5",
                    "gstin": "GSTIN999",
                    "address": ["Plot 42", "Industrial Area", "Bangalore 560001"]
                }
            },
            "agency_loads": {
                "MAHALASA": [
                    {"sl": 1, "vehicle_no": "KA 19 D 5917", "rate": 600.0, "loads": 25.0, "total_amount": 15000.0}
                ]
            },
            "custom_agencies": [
                {"sheet_name": "CUSTOM_NEW", "agency_name": "NEW CUSTOM AGENCY"}
            ]
        }
        ok = self.mgr.save_state(test_data)
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(self.state_file))

        loaded = self.mgr.load_state()
        self.assertEqual(loaded, test_data)
        self.assertEqual(loaded["agency_meta_overrides"]["MAHALASA"]["agency_name"], "MAHALASA GAS ENTERPRISES")
        self.assertEqual(loaded["agency_loads"]["MAHALASA"][0]["rate"], 600.0)

    def test_clear_state(self):
        self.mgr.save_state({"test": 123})
        self.assertTrue(os.path.exists(self.state_file))
        self.mgr.clear_state()
        self.assertFalse(os.path.exists(self.state_file))
        self.assertEqual(self.mgr.load_state(), {})

    def test_corrupted_json_handling(self):
        with open(self.state_file, "w") as f:
            f.write("{invalid json content")
        loaded = self.mgr.load_state()
        self.assertEqual(loaded, {})


if __name__ == "__main__":
    unittest.main()
