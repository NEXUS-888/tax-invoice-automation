import os
import sys
import unittest
import tempfile
import threading
import time
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from whatsapp_dispatcher import WhatsAppDispatcher
from whatsapp_automator import WhatsAppAutomator, build_invoice_message

REMINDER = "Note : please complete the payment before 10th of this month."


def _make_pdf():
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
        tf.write(b"%PDF-1.4 dummy")
        return tf.name


class TestFormatting(unittest.TestCase):
    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()

    def test_format_phone_number(self):
        fmt = self.dispatcher.format_phone_number
        self.assertEqual(fmt("9876543210"), "919876543210")
        self.assertEqual(fmt("+91 98765 43210"), "919876543210")
        self.assertEqual(fmt("919876543210"), "919876543210")
        self.assertEqual(fmt("098765-43210"), "919876543210")   # domestic trunk 0
        self.assertEqual(fmt("0091 98765 43210"), "919876543210")
        self.assertEqual(fmt("+1 415 555 2671"), "14155552671")  # other countries need + or 00
        self.assertEqual(fmt("4155552671"), "")
        self.assertEqual(fmt("+91 12345 67890"), "")             # Indian numbers still need a mobile prefix
        self.assertEqual(fmt(None), "")
        self.assertEqual(fmt("5876543210"), "")          # not a valid Indian mobile prefix
        self.assertEqual(fmt("invalid123"), "")
        self.assertEqual(fmt(""), "")
        self.assertEqual(fmt(9876543210), "919876543210")

    def test_build_invoice_message_has_exact_reminder(self):
        msg = self.dispatcher.build_invoice_message("TEST AGENCY", "INV-101", "SEPTEMBER 2026")
        self.assertIn("Dear TEST AGENCY,", msg)
        self.assertIn("(Invoice No: INV-101) for SEPTEMBER 2026.", msg)
        self.assertIn(REMINDER, msg)
        self.assertTrue(msg.endswith("Thank you,\nANANYA ENTERPRISES"))

    def test_dispatcher_and_automator_messages_match(self):
        self.assertEqual(
            self.dispatcher.build_invoice_message("A", "1", "M"),
            build_invoice_message("A", "1", "M"),
        )

    def test_build_share_url_variants(self):
        url = self.dispatcher.build_share_url("desktop", "919876543210", "Hi & bye")
        self.assertEqual(url, "whatsapp://send?phone=919876543210&text=Hi%20%26%20bye")
        self.assertEqual(self.dispatcher.build_share_url("app", "", "x"), "whatsapp://send?text=x")
        self.assertTrue(self.dispatcher.build_share_url("web", "919876543210", "x")
                        .startswith("https://web.whatsapp.com/send?phone=919876543210&text="))
        self.assertEqual(self.dispatcher.build_share_url("web", "", "x"), "https://web.whatsapp.com/")
        self.assertTrue(self.dispatcher.build_share_url("auto", "919876543210", "x")
                        .startswith("https://api.whatsapp.com/send?phone=919876543210"))


class TestClipboard(unittest.TestCase):
    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()
        self.pdf = _make_pdf()

    def tearDown(self):
        os.remove(self.pdf)

    def test_missing_file_returns_false(self):
        self.assertFalse(self.dispatcher.copy_pdf_to_clipboard("non_existent_file.pdf"))
        self.assertFalse(self.dispatcher.copy_pdf_to_clipboard(""))

    @unittest.skipUnless(sys.platform == "win32", "Win32 clipboard")
    def test_hdrop_round_trip(self):
        """The clipboard must contain exactly our file as CF_HDROP."""
        import ctypes
        from ctypes import wintypes
        self.assertTrue(self.dispatcher.copy_pdf_to_clipboard(self.pdf))

        user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
        user32.GetClipboardData.restype = wintypes.HANDLE
        shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
        self.assertTrue(user32.OpenClipboard(None))
        try:
            h = user32.GetClipboardData(15)
            self.assertTrue(h)
            self.assertEqual(shell32.DragQueryFileW(h, 0xFFFFFFFF, None, 0), 1)
            buf = ctypes.create_unicode_buffer(1024)
            shell32.DragQueryFileW(h, 0, buf, 1024)
            self.assertEqual(os.path.normcase(buf.value), os.path.normcase(os.path.abspath(self.pdf)))
        finally:
            user32.CloseClipboard()

    @patch("whatsapp_dispatcher.subprocess.run")
    @patch.object(WhatsAppDispatcher, "_set_clipboard_hdrop", return_value=False)
    def test_powershell_fallback_escapes_quotes(self, _hdrop, mock_run):
        if sys.platform != "win32":
            self.skipTest("PowerShell fallback is Windows-only")
        weird = os.path.join(tempfile.gettempdir(), "o'brien invoice.pdf")
        with open(weird, "wb") as f:
            f.write(b"%PDF")
        try:
            self.assertTrue(self.dispatcher.copy_pdf_to_clipboard(weird))
            cmd = mock_run.call_args[0][0][-1]
            self.assertIn("o''brien invoice.pdf'", cmd)
        finally:
            os.remove(weird)

    @patch("whatsapp_dispatcher.subprocess.run", side_effect=OSError)
    @patch.object(WhatsAppDispatcher, "_set_clipboard_hdrop", return_value=False)
    def test_qt_clipboard_not_touched_off_main_thread(self, _hdrop, _run):
        fake_qapp = MagicMock()
        results = []
        with patch("whatsapp_dispatcher.QApplication", fake_qapp):
            t = threading.Thread(target=lambda: results.append(self.dispatcher.copy_pdf_to_clipboard(self.pdf)))
            t.start()
            t.join()
        self.assertEqual(results, [False])
        fake_qapp.instance.assert_not_called()


@patch("whatsapp_dispatcher.subprocess.Popen")
class TestFallback(unittest.TestCase):
    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()
        self.pdf = _make_pdf()

    def tearDown(self):
        os.remove(self.pdf)

    @patch.object(WhatsAppDispatcher, "copy_pdf_to_clipboard", return_value=True)
    def test_fallback_copies_and_selects_in_explorer(self, mock_copy, mock_popen):
        msg = self.dispatcher.run_fallback(self.pdf, reason="Timed out.")
        mock_copy.assert_called_once_with(self.pdf)
        if sys.platform == "win32":
            self.assertEqual(mock_popen.call_args[0][0], f'explorer /select,"{self.pdf}"')
        self.assertIn("Ctrl+V", msg)
        self.assertTrue(msg.startswith("Timed out."))

    @patch.object(WhatsAppDispatcher, "copy_pdf_to_clipboard", return_value=False)
    def test_fallback_without_clipboard_still_says_drag(self, _copy, mock_popen):
        msg = self.dispatcher.run_fallback(self.pdf)
        self.assertIn("drag", msg)
        self.assertNotIn("Ctrl+V", msg)


@patch("whatsapp_dispatcher.subprocess.Popen")
class TestShareEntryPoint(unittest.TestCase):
    """Share App: PDF copied to the clipboard + WhatsApp opened with the message; the user presses Ctrl+V."""

    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()
        self.pdf = _make_pdf()
        p = patch.object(WhatsAppDispatcher, "copy_pdf_to_clipboard", return_value=True)
        self.copy = p.start()
        self.addCleanup(p.stop)
        m = patch.object(WhatsAppDispatcher, "maximize_whatsapp_soon")  # never touch the real window
        self.maximize = m.start()
        self.addCleanup(m.stop)

    def tearDown(self):
        os.remove(self.pdf)

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_app_with_phone_copies_pdf_and_opens_chat_with_message(self, mock_open, mock_popen):
        ok, msg = self.dispatcher.share_invoice_to_whatsapp(
            "ACME FUELS", "9876543210", "105", "SEP 2026", self.pdf, target="desktop")
        self.assertTrue(ok)
        self.copy.assert_called_once_with(os.path.abspath(self.pdf))
        url = mock_open.call_args[0][0]
        self.assertTrue(url.startswith("whatsapp://send?phone=919876543210&text="))
        self.assertIn("Note%20%3A%20please%20complete%20the%20payment", url)
        self.assertEqual(mock_open.call_count, 1)  # opened exactly once: no duplicate message
        self.assertIn("Ctrl+V in the chat", msg)
        mock_popen.assert_not_called()  # no Explorer window

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_app_maximizes_whatsapp_after_opening(self, _open, _popen):
        self.dispatcher.share_invoice_to_whatsapp("ACME", "9876543210", "1", "M", self.pdf, target="desktop")
        self.maximize.assert_called_once()

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_web_does_not_touch_whatsapp_window(self, _open, _popen):
        self.dispatcher.share_invoice_to_whatsapp("ACME", "9876543210", "1", "M", self.pdf, target="web")
        self.maximize.assert_not_called()

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_app_without_phone_opens_contact_picker(self, mock_open, mock_popen):
        ok, msg = self.dispatcher.share_invoice_to_whatsapp("ACME", None, "1", "M", self.pdf, target="desktop")
        self.assertTrue(ok)
        self.assertTrue(mock_open.call_args[0][0].startswith("whatsapp://send?text="))
        self.assertIn("after picking the contact", msg)
        self.copy.assert_called_once()
        mock_popen.assert_not_called()

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_pdf_copied_before_whatsapp_opens(self, mock_open, _popen):
        order = []
        self.copy.side_effect = lambda *_: order.append("copy") or True
        mock_open.side_effect = lambda *_: order.append("open") or True
        self.dispatcher.share_invoice_to_whatsapp("ACME", "9876543210", "1", "M", self.pdf, target="desktop")
        self.assertEqual(order, ["copy", "open"])

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_invalid_number_refuses_instead_of_opening_picker(self, mock_open, _popen):
        ok, msg = self.dispatcher.share_invoice_to_whatsapp("ACME", "12345", "1", "M", self.pdf, target="desktop")
        self.assertFalse(ok)
        self.assertIn("not a valid WhatsApp number", msg)
        mock_open.assert_not_called()
        self.copy.assert_not_called()

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="hint")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_copy_failure_falls_back_to_explorer(self, _open, mock_fallback, _popen):
        self.copy.return_value = False
        ok, msg = self.dispatcher.share_invoice_to_whatsapp("ACME", "9876543210", "1", "M", self.pdf)
        self.assertTrue(ok)
        self.assertIn("hint", msg)
        mock_fallback.assert_called_once()

    @patch.object(WhatsAppDispatcher, "_open_url", side_effect=[False, True])
    def test_desktop_not_installed_falls_back_to_web(self, mock_open, _popen):
        ok, msg = self.dispatcher.share_invoice_to_whatsapp(
            "ACME", "9876543210", "1", "M", self.pdf, target="desktop")
        self.assertTrue(ok)
        self.assertTrue(mock_open.call_args_list[1][0][0].startswith("https://web.whatsapp.com/send?phone="))
        self.assertIn("WhatsApp Web", msg)
        self.assertIn("Ctrl+V", msg)

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_web_target_copies_pdf_too(self, mock_open, mock_popen):
        ok, msg = self.dispatcher.share_invoice_to_whatsapp("ACME", "9876543210", "1", "M", self.pdf, target="web")
        self.assertTrue(ok)
        self.assertTrue(mock_open.call_args[0][0].startswith("https://web.whatsapp.com/send?phone="))
        self.copy.assert_called_once()
        mock_popen.assert_not_called()

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="hint")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=False)
    def test_nothing_opens_reports_failure(self, _open, _fallback, _popen):
        ok, msg = self.dispatcher.share_invoice_to_whatsapp("ACME", "9876543210", "1", "M", self.pdf, target="web")
        self.assertFalse(ok)
        self.assertIn("hint", msg)

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_legacy_alias_returns_pair(self, _open, _popen):
        result = self.dispatcher.send_agency_invoice("ACME", "9876543210", "1", "M", self.pdf)
        self.assertEqual(len(result), 2)
        self.assertTrue(result[0])


@unittest.skipUnless(sys.platform == "win32", "Windows only")
class TestMaximizeWhatsApp(unittest.TestCase):
    def _run(self, zoomed, iconic=lambda i: False, windows=(111,), timeout=30):
        """zoomed(i)/iconic(i): window state on the i-th check (one check per 0.5 s loop)."""
        calls = {"n": 0}
        user32 = MagicMock()
        user32.IsZoomed.side_effect = lambda h: zoomed(calls["n"])
        user32.IsIconic.side_effect = lambda h: iconic(calls["n"])

        def sleep(_):
            calls["n"] += 1
        clock = iter(range(0, 10000))
        with patch.object(WhatsAppDispatcher, "_whatsapp_windows", return_value=list(windows)),                 patch("ctypes.windll.user32", user32, create=True),                 patch("whatsapp_dispatcher.time.sleep", side_effect=sleep),                 patch("whatsapp_dispatcher.time.time", side_effect=lambda: next(clock)):
            WhatsAppDispatcher().maximize_whatsapp_soon(timeout=timeout).join(5)
        return user32

    def test_maximizes_restored_window(self):
        user32 = self._run(lambda i: i > 0)
        user32.ShowWindow.assert_called_once_with(111, 3)

    def test_re_maximizes_when_whatsapp_shrinks_late(self):
        """WhatsApp shrinking 20 s after opening (slow laptop) is still caught."""
        user32 = self._run(lambda i: not (i == 0 or i == 20))
        self.assertEqual(user32.ShowWindow.call_count, 2)

    def test_leaves_maximized_window_alone(self):
        user32 = self._run(lambda i: True)
        user32.ShowWindow.assert_not_called()

    def test_leaves_window_minimized_by_user_alone(self):
        user32 = self._run(lambda i: False, iconic=lambda i: True)
        user32.ShowWindow.assert_not_called()

    def test_no_whatsapp_window_gives_up_quietly(self):
        user32 = self._run(lambda i: False, windows=())
        user32.ShowWindow.assert_not_called()


class TestPlaywrightSharePreview(unittest.TestCase):
    """prepare_invoice_share: deterministic attach via the connected WhatsApp Web session."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.auto = WhatsAppAutomator(self.tmp, log_callback=lambda m: None)
        self.auto.page = MagicMock()
        self.pdf = _make_pdf()
        self.helpers = {}
        for name, value in (
            ("is_logged_in", True), ("_open_chat_for_number", True), ("_wait_for_chat", True),
            ("_dismiss_failed_popup", False), ("_attach_pdf_document", True),
            ("_wait_for_document_preview", True), ("_type_caption_in_preview", True),
            ("_click_send", True), ("_wait_for_upload_and_delivery", True),
        ):
            p = patch.object(WhatsAppAutomator, name, return_value=value)
            self.helpers[name] = p.start()
            self.addCleanup(p.stop)

    def tearDown(self):
        os.remove(self.pdf)

    def test_attaches_and_leaves_send_to_user(self):
        ok, msg = self.auto.prepare_invoice_share("9876543210", "ACME", "105", "SEP", self.pdf)
        self.assertTrue(ok)
        self.helpers["_open_chat_for_number"].assert_called_once_with("919876543210")
        self.helpers["_attach_pdf_document"].assert_called_once_with(os.path.abspath(self.pdf))
        caption = self.helpers["_type_caption_in_preview"].call_args[0][0]
        self.assertIn(REMINDER, caption)
        self.helpers["_click_send"].assert_not_called()

    def test_auto_send_clicks_send(self):
        ok, msg = self.auto.prepare_invoice_share("9876543210", "ACME", "105", "SEP", self.pdf, auto_send=True)
        self.assertTrue(ok)
        self.helpers["_click_send"].assert_called_once()

    def test_invalid_phone_popup(self):
        self.helpers["_wait_for_chat"].return_value = "INVALID_PHONE"
        ok, msg = self.auto.prepare_invoice_share("9876543210", "ACME", "1", "M", self.pdf)
        self.assertFalse(ok)
        self.assertIn("not on WhatsApp", msg)
        self.helpers["_attach_pdf_document"].assert_not_called()

    def test_attach_failure_reported(self):
        self.helpers["_attach_pdf_document"].return_value = False
        ok, msg = self.auto.prepare_invoice_share("9876543210", "ACME", "1", "M", self.pdf)
        self.assertFalse(ok)

    def test_rejects_bad_inputs_without_touching_page(self):
        self.assertFalse(self.auto.prepare_invoice_share("123", "ACME", "1", "M", self.pdf)[0])
        self.assertFalse(self.auto.prepare_invoice_share("9876543210", "ACME", "1", "M", "missing.pdf")[0])
        self.helpers["_open_chat_for_number"].assert_not_called()

    def test_not_connected(self):
        self.auto.page = None
        self.assertEqual(self.auto.prepare_invoice_share("9876543210", "A", "1", "M", self.pdf),
                         (False, "WhatsApp not connected."))


class TestCaptionTyping(unittest.TestCase):
    def _auto(self, box_texts):
        auto = WhatsAppAutomator(tempfile.mkdtemp(), log_callback=lambda m: None)
        auto.page = MagicMock()
        el = MagicMock()
        el.is_visible.return_value = True
        el.inner_text.side_effect = list(box_texts)  # before typing, after typing
        auto.page.query_selector_all.return_value = [el]
        return auto, el

    def test_newlines_use_shift_enter_not_enter(self):
        auto, el = self._auto(["", "line1\n\nline3"])
        with patch("whatsapp_automator.time.sleep"):
            self.assertTrue(auto._type_caption_in_preview("line1\n\nline3"))
        el.focus.assert_called_once()
        el.click.assert_not_called()
        pressed = [c[0][0] for c in auto.page.keyboard.press.call_args_list]
        self.assertEqual(pressed, ["Shift+Enter", "Shift+Enter"])
        inserted = "".join(c[0][0] for c in auto.page.keyboard.insert_text.call_args_list)
        self.assertEqual(inserted, "line1line3")

    def test_half_typed_caption_is_replaced_not_appended(self):
        auto, _ = self._auto(["line1 partial", "line1\nline3"])
        with patch("whatsapp_automator.time.sleep"):
            self.assertTrue(auto._type_caption_in_preview("line1\nline3"))
        pressed = [c[0][0] for c in auto.page.keyboard.press.call_args_list]
        self.assertEqual(pressed[:2], ["Control+A", "Delete"])

    def test_reports_failure_when_caption_did_not_stick(self):
        auto, _ = self._auto(["", ""])
        with patch("whatsapp_automator.time.sleep"):
            self.assertFalse(auto._type_caption_in_preview("line1\nline3"))

    def test_waits_for_caption_box_to_appear(self):
        auto, el = self._auto(["", "line1\nline3"])
        auto.page.query_selector_all.side_effect = lambda sel: [el] if calls.append(sel) or len(calls) > 7 else []
        calls = []
        with patch("whatsapp_automator.time.sleep"):
            self.assertTrue(auto._type_caption_in_preview("line1\nline3"))

    def test_no_caption_box_gives_up(self):
        auto, _ = self._auto([])
        auto.page.query_selector_all.return_value = []
        clock = iter(range(0, 100))
        with patch("whatsapp_automator.time.sleep"), patch("whatsapp_automator.time.time", side_effect=lambda: next(clock)):
            self.assertFalse(auto._type_caption_in_preview("line1", timeout=5))

    def test_reports_failure_on_duplicated_caption(self):
        auto, _ = self._auto(["", "line1\nline3\nline1\nline3"])
        with patch("whatsapp_automator.time.sleep"):
            self.assertFalse(auto._type_caption_in_preview("line1\nline3"))


if __name__ == "__main__":
    unittest.main()
