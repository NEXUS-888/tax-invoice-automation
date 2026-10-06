import os
import sys
import unittest
import tempfile
import shutil
from unittest.mock import MagicMock, patch

# Add src to python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from whatsapp_automator import WhatsAppAutomator


class TestWhatsAppAutomator(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.automator = WhatsAppAutomator(self.test_dir)

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_format_phone(self):
        # 10-digit valid numbers
        self.assertEqual(self.automator.format_phone("9876543210"), "919876543210")
        self.assertEqual(self.automator.format_phone("+91 9141844840"), "919141844840")
        self.assertEqual(self.automator.format_phone("91-98765-43210"), "919876543210")
        
        # Invalid / short numbers
        self.assertEqual(self.automator.format_phone("12345"), "")
        self.assertEqual(self.automator.format_phone("1234567890"), "")  # doesn't start with 6789
        self.assertEqual(self.automator.format_phone(""), "")

    def test_parse_phones(self):
        # String with multiple phone numbers separated by comma, slash, newline
        input_str = "9876543210, 919141844840 / 9742108907\n9876543210"
        result = self.automator.parse_phones(input_str)
        self.assertEqual(result, ["919876543210", "919141844840", "919742108907"])

        # List input
        input_list = ["9876543210", "919141844840", "invalid"]
        result = self.automator.parse_phones(input_list)
        self.assertEqual(result, ["919876543210", "919141844840"])

    def test_cleanup_lockfiles(self):
        # Create mock lockfiles
        lockfiles = ["lockfile", "SingletonLock", "SingletonCookie", "SingletonSocket", "DevToolsActivePort"]
        for f in lockfiles:
            open(os.path.join(self.test_dir, f), 'w').close()

        self.automator._clean_session_lockfiles()

        for f in lockfiles:
            self.assertFalse(os.path.exists(os.path.join(self.test_dir, f)), f"{f} was not cleaned up")

    def test_send_invoice_pdf_validation(self):
        # When page is None
        ok, msg = self.automator.send_invoice_pdf("9876543210", "Test Agency", "101", "AUG 2026", "fake.pdf")
        self.assertFalse(ok)
        self.assertIn("not connected", msg)

        # Mock page connected
        self.automator.page = MagicMock()

        # Invalid phone
        ok, msg = self.automator.send_invoice_pdf("invalid", "Test Agency", "101", "AUG 2026", "fake.pdf")
        self.assertFalse(ok)
        self.assertIn("No valid phone", msg)

        # Non-existent PDF
        ok, msg = self.automator.send_invoice_pdf("9876543210", "Test Agency", "101", "AUG 2026", "non_existent.pdf")
        self.assertFalse(ok)
        self.assertIn("PDF not found", msg)

    def test_wait_for_chat_invalid_phone_detection(self):
        mock_page = MagicMock()
        self.automator.page = mock_page

        # Simulate visible popup for invalid number
        mock_popup = MagicMock()
        mock_popup.is_visible.return_value = True
        mock_popup.inner_text.return_value = "Phone number shared via url is invalid."
        
        mock_page.query_selector.side_effect = lambda sel: mock_popup if "dialog" in sel or "popup" in sel else None

        res = self.automator._wait_for_chat(timeout=1000)
        self.assertEqual(res, "INVALID_PHONE")

    @patch("playwright.sync_api.sync_playwright")
    def test_launch_session_stealth_configuration(self, mock_sync_playwright):
        mock_pw = MagicMock()
        mock_sync_playwright.return_value.start.return_value = mock_pw
        mock_context = MagicMock()
        mock_pw.chromium.launch_persistent_context.return_value = mock_context
        mock_page = MagicMock()
        mock_context.pages = [mock_page]

        res = self.automator.launch_session(headless=False)
        self.assertTrue(res)

        # Check launch parameters
        call_kwargs = mock_pw.chromium.launch_persistent_context.call_args[1]
        self.assertIn("user_agent", call_kwargs)
        self.assertIn("Chrome/", call_kwargs["user_agent"])
        self.assertIn("args", call_kwargs)
        self.assertIn("--disable-blink-features=AutomationControlled", call_kwargs["args"])

        # Check stealth init script
        mock_context.add_init_script.assert_called_once()
        init_script = mock_context.add_init_script.call_args[0][0]
        self.assertIn("navigator", init_script)
        self.assertIn("webdriver", init_script)

    def test_open_chat_for_number_spa(self):
        mock_page = MagicMock()
        mock_page.url = "https://web.whatsapp.com"
        self.automator.page = mock_page

        res = self.automator._open_chat_for_number("919876543210")
        self.assertTrue(res)
        mock_page.evaluate.assert_called_once_with(
            "(url) => { window.location.href = url; }",
            "https://web.whatsapp.com/send?phone=919876543210"
        )


if __name__ == "__main__":
    unittest.main()

