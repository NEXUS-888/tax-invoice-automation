import os
import sys
import tempfile
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from whatsapp_desktop_bridge import WhatsAppDesktopBridge, ENV_VAR
from whatsapp_automator import WhatsAppAutomator

WRAPPER_URL = ("https://web.whatsapp.com/?osBuild=26200&windowsBuild=263802&launchContext=Launch"
               "&windows=1")


class TestChatUrl(unittest.TestCase):
    def test_keeps_desktop_wrapper_parameters(self):
        url = WhatsAppDesktopBridge.chat_url(WRAPPER_URL, "919876543210")
        self.assertTrue(url.startswith("https://web.whatsapp.com/send?phone=919876543210&"))
        for part in ("osBuild=26200", "windowsBuild=263802", "launchContext=Launch", "windows=1"):
            self.assertIn(part, url)

    def test_drops_old_phone_and_text(self):
        url = WhatsAppDesktopBridge.chat_url(
            "https://web.whatsapp.com/send?phone=911111111111&text=hello&windows=1", "919876543210")
        self.assertNotIn("text=", url)
        self.assertNotIn("911111111111", url)
        self.assertEqual(url.count("phone="), 1)


class TestAvailability(unittest.TestCase):
    def setUp(self):
        self.bridge = WhatsAppDesktopBridge()

    def test_available_when_whatsapp_page_listed(self):
        targets = [{"type": "service_worker", "url": "https://web.whatsapp.com/sw.js"},
                   {"type": "page", "url": WRAPPER_URL}]
        with patch.object(self.bridge, "_get_json", return_value=targets):
            self.assertTrue(self.bridge.is_available())

    def test_not_available_without_page_or_port(self):
        with patch.object(self.bridge, "_get_json",
                          return_value=[{"type": "service_worker", "url": "https://web.whatsapp.com/sw.js"}]):
            self.assertFalse(self.bridge.is_available())
        with patch.object(self.bridge, "_get_json", side_effect=OSError("refused")):
            self.assertFalse(self.bridge.is_available())


@patch("whatsapp_desktop_bridge.time.sleep")
@patch("whatsapp_desktop_bridge.os.startfile", create=True)
@patch("whatsapp_desktop_bridge.subprocess.run")
class TestEnable(unittest.TestCase):
    def setUp(self):
        self.bridge = WhatsAppDesktopBridge()
        p = patch.object(WhatsAppDesktopBridge, "_set_user_env")
        self.set_env = p.start()
        self.addCleanup(p.stop)

    def test_already_connected_does_not_restart(self, mock_run, mock_start, _sleep):
        with patch.object(self.bridge, "is_available", return_value=True):
            self.assertEqual(self.bridge.enable(), (True, "already connected"))
        mock_run.assert_not_called()
        mock_start.assert_not_called()
        self.set_env.assert_not_called()

    def test_not_installed(self, mock_run, mock_start, _sleep):
        with patch.object(self.bridge, "is_available", return_value=False), \
                patch.object(self.bridge, "is_whatsapp_installed", return_value=False):
            ok, msg = self.bridge.enable()
        self.assertFalse(ok)
        self.assertIn("not installed", msg)
        mock_run.assert_not_called()

    def test_restarts_whatsapp_then_removes_switch(self, mock_run, mock_start, _sleep):
        mock_run.return_value = MagicMock(stdout="INFO: No tasks are running")
        with patch.object(self.bridge, "is_available", side_effect=[False, False, True]), \
                patch.object(self.bridge, "is_whatsapp_installed", return_value=True):
            self.assertEqual(self.bridge.enable(), (True, "connected"))
        self.assertEqual(self.set_env.call_args_list[0][0][0], f"--remote-debugging-port={self.bridge.port}")
        self.assertIsNone(self.set_env.call_args_list[-1][0][0])  # switch removed again
        self.assertIn("taskkill", mock_run.call_args_list[0][0][0])
        self.assertIn("WhatsAppDesktop", mock_start.call_args[0][0])

    def test_switch_removed_even_when_restart_fails(self, mock_run, mock_start, _sleep):
        mock_run.side_effect = OSError("taskkill missing")
        with patch.object(self.bridge, "is_available", return_value=False), \
                patch.object(self.bridge, "is_whatsapp_installed", return_value=True):
            ok, _ = self.bridge.enable()
        self.assertFalse(ok)
        self.assertIsNone(self.set_env.call_args_list[-1][0][0])

    def test_timeout_reported(self, mock_run, mock_start, _sleep):
        mock_run.return_value = MagicMock(stdout="")
        with patch.object(self.bridge, "is_available", return_value=False), \
                patch.object(self.bridge, "is_whatsapp_installed", return_value=True), \
                patch("whatsapp_desktop_bridge.time.time", side_effect=[0, 0] + [100] * 5):
            ok, msg = self.bridge.enable(timeout=5)
        self.assertFalse(ok)
        self.assertIn("did not accept", msg)


@unittest.skipUnless(sys.platform == "win32", "Windows only")
class TestUserEnvironment(unittest.TestCase):
    def test_set_and_remove_round_trip(self):
        import winreg

        def read():
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
                try:
                    return winreg.QueryValueEx(k, ENV_VAR)[0]
                except FileNotFoundError:
                    return None

        if read() is not None:
            self.skipTest("variable already set by the user; leave it alone")
        try:
            WhatsAppDesktopBridge._set_user_env("--test")
            self.assertEqual(read(), "--test")
        finally:
            WhatsAppDesktopBridge._set_user_env(None)
        self.assertIsNone(read())


class TestAttachInvoice(unittest.TestCase):
    """attach_invoice against a fake WhatsApp page: never touches the real app."""

    def setUp(self):
        self.bridge = WhatsAppDesktopBridge()
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
            tf.write(b"%PDF-1.4")
            self.pdf = tf.name
        self.page = MagicMock()
        self.page.url = WRAPPER_URL
        context = MagicMock(pages=[self.page])
        browser = MagicMock(contexts=[context])
        pw = MagicMock()
        pw.chromium.connect_over_cdp.return_value = browser
        self.pw = pw
        starter = MagicMock()
        starter.start.return_value = pw
        p = patch("playwright.sync_api.sync_playwright", return_value=starter)
        p.start()
        self.addCleanup(p.stop)
        self.helpers = {}
        for name, value in (("is_logged_in", True), ("_wait_for_chat", True), ("_dismiss_failed_popup", False),
                            ("_attach_pdf_document", True), ("_wait_for_document_preview", True),
                            ("_type_caption_in_preview", True), ("_click_send", True),
                            ("_wait_for_upload_and_delivery", True)):
            hp = patch.object(WhatsAppAutomator, name, return_value=value)
            self.helpers[name] = hp.start()
            self.addCleanup(hp.stop)
        bp = patch.object(WhatsAppDesktopBridge, "bring_to_front")
        self.front = bp.start()
        self.addCleanup(bp.stop)

    def tearDown(self):
        os.remove(self.pdf)

    def test_attaches_with_caption_and_never_sends(self):
        ok, msg = self.bridge.attach_invoice("919876543210", "invoice message", self.pdf)
        self.assertTrue(ok)
        self.assertIn("press Send", msg)
        goto_url = self.page.goto.call_args[0][0]
        self.assertTrue(goto_url.startswith("https://web.whatsapp.com/send?phone=919876543210"))
        self.helpers["_attach_pdf_document"].assert_called_once_with(os.path.abspath(self.pdf))
        self.helpers["_type_caption_in_preview"].assert_called_once_with("invoice message")
        self.helpers["_click_send"].assert_not_called()
        self.pw.stop.assert_called_once()  # connection dropped, WhatsApp left running

    def test_whatsapp_shown_before_driving_it(self):
        order = []
        self.front.side_effect = lambda: order.append("front")
        self.page.goto.side_effect = lambda *a, **k: order.append("goto")
        self.bridge.attach_invoice("919876543210", "m", self.pdf)
        self.assertEqual(order[:2], ["front", "goto"])

    def test_invalid_number(self):
        self.helpers["_wait_for_chat"].return_value = "INVALID_PHONE"
        ok, msg = self.bridge.attach_invoice("919876543210", "m", self.pdf)
        self.assertFalse(ok)
        self.assertIn("not on WhatsApp", msg)
        self.helpers["_attach_pdf_document"].assert_not_called()

    def test_attach_rejected(self):
        self.helpers["_attach_pdf_document"].return_value = False
        ok, _ = self.bridge.attach_invoice("919876543210", "m", self.pdf)
        self.assertFalse(ok)
        self.helpers["_type_caption_in_preview"].assert_not_called()

    def test_caption_failure_still_attached(self):
        self.helpers["_type_caption_in_preview"].return_value = False
        ok, msg = self.bridge.attach_invoice("919876543210", "m", self.pdf)
        self.assertTrue(ok)
        self.assertIn("type the message yourself", msg)

    def test_no_whatsapp_page(self):
        self.page.url = "about:blank"
        ok, msg = self.bridge.attach_invoice("919876543210", "m", self.pdf)
        self.assertFalse(ok)
        self.assertIn("not found", msg)

    def test_connection_error_is_reported_not_raised(self):
        self.pw.chromium.connect_over_cdp.side_effect = RuntimeError("ECONNREFUSED")
        ok, msg = self.bridge.attach_invoice("919876543210", "m", self.pdf)
        self.assertFalse(ok)
        self.assertIn("ECONNREFUSED", msg)


if __name__ == "__main__":
    unittest.main()
