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


class TestDesktopAttach(unittest.TestCase):
    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()
        self.pdf = _make_pdf()

    def tearDown(self):
        os.remove(self.pdf)

    def _run(self, bridge):
        done = []
        with patch("whatsapp_dispatcher.WhatsAppDesktopBridge", return_value=bridge):
            t = self.dispatcher.attach_in_background("919876543210", "msg", self.pdf,
                                                     on_complete=lambda ok, m: done.append((ok, m)))
            t.join(5)
        return done

    @patch.object(WhatsAppDispatcher, "run_fallback")
    def test_success_attaches_inside_whatsapp(self, mock_fallback):
        bridge = MagicMock()
        bridge.enable.return_value = (True, "connected")
        bridge.attach_invoice.return_value = (True, "PDF attached")
        self.assertEqual(self._run(bridge), [(True, "PDF attached")])
        bridge.attach_invoice.assert_called_once_with("919876543210", "msg", self.pdf)
        mock_fallback.assert_not_called()

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="fallback hint")
    def test_connect_failure_skips_attach_and_falls_back(self, mock_fallback):
        bridge = MagicMock()
        bridge.enable.return_value = (False, "not installed")
        self.assertEqual(self._run(bridge), [(False, "fallback hint")])
        bridge.attach_invoice.assert_not_called()
        self.assertIn("not installed", mock_fallback.call_args.kwargs["reason"])

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="fallback hint")
    def test_attach_failure_falls_back(self, mock_fallback):
        bridge = MagicMock()
        bridge.enable.return_value = (True, "connected")
        bridge.attach_invoice.return_value = (False, "chat did not open")
        self.assertEqual(self._run(bridge), [(False, "fallback hint")])

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="fallback hint")
    def test_exception_never_escapes_thread(self, _fallback):
        bridge = MagicMock()
        bridge.enable.side_effect = RuntimeError("boom")
        self.assertEqual(self._run(bridge), [(False, "fallback hint")])

    @patch.object(WhatsAppDispatcher, "run_fallback")
    def test_shares_run_one_at_a_time(self, _fallback):
        active, peak = [0], [0]
        lock = threading.Lock()

        def slow_attach(*_):
            with lock:
                active[0] += 1
                peak[0] = max(peak[0], active[0])
            time.sleep(0.2)
            with lock:
                active[0] -= 1
            return True, "ok"

        bridge = MagicMock()
        bridge.enable.return_value = (True, "connected")
        bridge.attach_invoice.side_effect = slow_attach
        with patch("whatsapp_dispatcher.WhatsAppDesktopBridge", return_value=bridge):
            threads = [self.dispatcher.attach_in_background("919876543210", "m", self.pdf) for _ in range(3)]
            for t in threads:
                t.join(5)
        self.assertEqual(peak[0], 1)
        self.assertEqual(bridge.attach_invoice.call_count, 3)

    def test_find_windows_excludes_non_whatsapp_processes(self):
        """This test process (python.exe) owns no WhatsApp windows, so nothing of ours may match."""
        own_pid = os.getpid()
        if sys.platform != "win32":
            self.assertEqual(self.dispatcher.find_whatsapp_desktop_windows(), [])
            return
        import ctypes
        from ctypes import wintypes
        for hwnd, _title, exe in self.dispatcher.find_whatsapp_desktop_windows():
            self.assertTrue(exe.startswith("whatsapp"))
            pid = wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            self.assertNotEqual(pid.value, own_pid)

    def test_force_foreground_invalid_hwnd(self):
        self.assertFalse(self.dispatcher.force_foreground(0))


@patch("whatsapp_dispatcher.subprocess.Popen")
@patch("whatsapp_dispatcher.webbrowser.open", return_value=True)
class TestShareEntryPoint(unittest.TestCase):
    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()
        self.pdf = _make_pdf()
        for name, value in (("find_whatsapp_desktop_windows", []), ("copy_pdf_to_clipboard", True)):
            p = patch.object(WhatsAppDispatcher, name, return_value=value)
            p.start()
            self.addCleanup(p.stop)

    def tearDown(self):
        os.remove(self.pdf)

    @patch.object(WhatsAppDispatcher, "attach_in_background")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_desktop_with_phone_attaches_without_whatsapp_link(self, mock_open, mock_bg, _wb, _popen):
        cb = MagicMock()
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME FUELS", "9876543210", "105", "SEP 2026", self.pdf, target="desktop", on_complete=cb)
        self.assertTrue(ok)
        self.assertTrue(pending)
        # A whatsapp:// link would open the "Send to" picker and send the text separately (twice).
        mock_open.assert_not_called()
        phone, message, pdf = mock_bg.call_args[0]
        self.assertEqual(phone, "919876543210")
        self.assertIn(REMINDER, message)
        self.assertEqual(pdf, os.path.abspath(self.pdf))
        self.assertIs(mock_bg.call_args.kwargs["on_complete"], cb)
        self.assertNotIn("attached", msg.lower().replace("attaching", ""))  # never claims success up front

    @patch.object(WhatsAppDispatcher, "attach_in_background")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_invalid_number_refuses_instead_of_opening_picker(self, mock_open, mock_bg, _wb, _popen):
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME", "12345", "1", "M", self.pdf, target="desktop")
        self.assertFalse(ok)
        self.assertFalse(pending)
        self.assertIn("not a valid WhatsApp number", msg)
        mock_open.assert_not_called()
        mock_bg.assert_not_called()

    @patch.object(WhatsAppDispatcher, "attach_in_background")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_desktop_without_phone_does_not_paste_blind(self, _open, mock_bg, _wb, _popen):
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME", None, "1", "M", self.pdf, target="desktop")
        self.assertTrue(ok)
        self.assertFalse(pending)
        mock_bg.assert_not_called()
        self.assertIn("Ctrl+V", msg)

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="hint")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_browser_web_runs_fallback_immediately(self, _open, mock_fallback, _wb, _popen):
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME", "9876543210", "1", "M", self.pdf, target="web")
        self.assertTrue(ok)
        self.assertFalse(pending)
        mock_fallback.assert_called_once_with(os.path.abspath(self.pdf))
        self.assertIn("hint", msg)

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="hint")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=False)
    def test_nothing_opens_reports_failure(self, _open, _fallback, _wb, _popen):
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME", "9876543210", "1", "M", self.pdf, target="web")
        self.assertFalse(ok)
        self.assertFalse(pending)

    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_legacy_alias_returns_pair(self, _open, _wb, _popen):
        with patch.object(WhatsAppDispatcher, "attach_in_background"):
            result = self.dispatcher.send_agency_invoice("ACME", "9876543210", "1", "M", self.pdf)
        self.assertEqual(len(result), 2)


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
    def test_newlines_use_shift_enter_not_enter(self):
        auto = WhatsAppAutomator(tempfile.mkdtemp(), log_callback=lambda m: None)
        auto.page = MagicMock()
        el = MagicMock()
        el.is_visible.return_value = True
        auto.page.query_selector_all.return_value = [el]
        with patch("whatsapp_automator.time.sleep"):
            self.assertTrue(auto._type_caption_in_preview("line1\n\nline3"))
        pressed = [c[0][0] for c in auto.page.keyboard.press.call_args_list]
        self.assertEqual(pressed, ["Shift+Enter", "Shift+Enter"])
        typed = "".join(c[0][0] for c in auto.page.keyboard.type.call_args_list)
        self.assertNotIn("\n", typed)


if __name__ == "__main__":
    unittest.main()
