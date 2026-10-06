import os
import unittest
import tempfile
from unittest.mock import patch

import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from whatsapp_dispatcher import WhatsAppDispatcher


class TestWhatsAppDispatcher(unittest.TestCase):
    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()

    def test_format_phone_number(self):
        self.assertEqual(self.dispatcher.format_phone_number("9876543210"), "919876543210")
        self.assertEqual(self.dispatcher.format_phone_number("+91 98765 43210"), "919876543210")
        self.assertEqual(self.dispatcher.format_phone_number("919876543210"), "919876543210")
        self.assertEqual(self.dispatcher.format_phone_number("invalid123"), "")
        self.assertEqual(self.dispatcher.format_phone_number(""), "")

    def test_build_invoice_message(self):
        msg = self.dispatcher.build_invoice_message("TEST AGENCY", "INV-101", "SEPTEMBER 2026")
        self.assertIn("TEST AGENCY", msg)
        self.assertIn("INV-101", msg)
        self.assertIn("SEPTEMBER 2026", msg)
        self.assertIn("Note : please complete the payment before 10th of this month.", msg)
        self.assertIn("ANANYA ENTERPRISES", msg)

    def test_copy_pdf_to_clipboard(self):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            tf.write(b"%PDF-1.4 dummy")
            temp_pdf = tf.name

        try:
            self.assertFalse(self.dispatcher.copy_pdf_to_clipboard("non_existent_file.pdf"))
            result = self.dispatcher.copy_pdf_to_clipboard(temp_pdf)
            self.assertIsInstance(result, bool)
        finally:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)

    @patch("whatsapp_dispatcher.os.startfile")
    @patch("whatsapp_dispatcher.subprocess.Popen")
    def test_share_invoice_to_whatsapp_with_phone(self, mock_popen, mock_startfile):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            tf.write(b"%PDF-1.4 dummy")
            temp_pdf = tf.name

        try:
            ok, msg = self.dispatcher.share_invoice_to_whatsapp(
                agency_name="ACME FUELS",
                phone="9876543210",
                invoice_no="105",
                month_desc="SEP 2026",
                pdf_path=temp_pdf,
                open_explorer=False
            )
            self.assertTrue(ok)
            self.assertIn("ACME FUELS", msg)
            self.assertIn("919876543210", msg)
            self.assertTrue(mock_startfile.called)
            called_url = mock_startfile.call_args[0][0]
            self.assertTrue(called_url.startswith("whatsapp://send?phone=919876543210"))
            self.assertIn("Note", called_url)
        finally:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)

    @patch("whatsapp_dispatcher.WhatsAppDispatcher.auto_attach_pdf_in_background")
    @patch("whatsapp_dispatcher.os.startfile")
    def test_share_triggers_auto_attach(self, mock_startfile, mock_auto_attach):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            tf.write(b"%PDF-1.4 dummy")
            temp_pdf = tf.name

        try:
            ok, msg = self.dispatcher.share_invoice_to_whatsapp(
                agency_name="ACME FUELS",
                phone="9876543210",
                invoice_no="105",
                month_desc="SEP 2026",
                pdf_path=temp_pdf,
                target="desktop"
            )
            self.assertTrue(ok)
            self.assertTrue(mock_auto_attach.called)
            self.assertIn("attached directly to chat", msg)
        finally:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)

    def test_activate_and_send_paste_handles_invalid_hwnd(self):
        # Should return safely without raising exceptions
        result = self.dispatcher.activate_and_send_paste(999999999)
        self.assertIsInstance(result, bool)


if __name__ == "__main__":
    unittest.main()

