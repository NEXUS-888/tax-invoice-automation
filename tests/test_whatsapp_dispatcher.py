import os
import sys
import unittest
import tempfile
import threading
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
        self.assertEqual(fmt("098765-43210"), "")       # leading trunk zero -> 11 digits, rejected
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

    def _patch(self, **kw):
        defaults = dict(
            find_whatsapp_desktop_windows=[(1234, "WhatsApp", "whatsapp.root.exe")],
            copy_pdf_to_clipboard=True,
            force_foreground=True,
            send_ctrl_v=True,
        )
        defaults.update(kw)
        patches = [patch.object(WhatsAppDispatcher, k, return_value=v) for k, v in defaults.items()]
        mocks = {}
        for name, p in zip(defaults, patches):
            mocks[name] = p.start()
            self.addCleanup(p.stop)
        sleep = patch("whatsapp_dispatcher.time.sleep")
        sleep.start()
        self.addCleanup(sleep.stop)
        return mocks

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_success_path_focuses_then_pastes(self):
        m = self._patch()
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf, was_running=True)
        self.assertTrue(ok)
        m["force_foreground"].assert_called_once_with(1234)
        m["send_ctrl_v"].assert_called_once()

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_no_window_times_out(self):
        m = self._patch(find_whatsapp_desktop_windows=[])
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf, window_timeout=0.01)
        self.assertFalse(ok)
        self.assertIn("did not appear", msg)
        m["send_ctrl_v"].assert_not_called()

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_focus_refused_does_not_paste(self):
        m = self._patch(force_foreground=False)
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf)
        self.assertFalse(ok)
        m["send_ctrl_v"].assert_not_called()

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_uipi_blocked_input_reported(self):
        self._patch(send_ctrl_v=False)
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf)
        self.assertFalse(ok)
        self.assertIn("blocked", msg)

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="fallback hint")
    @patch.object(WhatsAppDispatcher, "attach_via_desktop", return_value=(False, "nope"))
    def test_background_failure_runs_fallback_and_reports(self, _attach, mock_fallback):
        done = []
        t = self.dispatcher.auto_attach_pdf_in_background(self.pdf, on_complete=lambda ok, m: done.append((ok, m)))
        t.join(5)
        self.assertEqual(done, [(False, "fallback hint")])
        mock_fallback.assert_called_once()

    @patch.object(WhatsAppDispatcher, "run_fallback")
    @patch.object(WhatsAppDispatcher, "attach_via_desktop", return_value=(True, "pasted"))
    def test_background_success_skips_fallback(self, _attach, mock_fallback):
        done = []
        t = self.dispatcher.auto_attach_pdf_in_background(self.pdf, on_complete=lambda ok, m: done.append((ok, m)))
        t.join(5)
        self.assertEqual(done, [(True, "pasted")])
        mock_fallback.assert_not_called()

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="fallback hint")
    @patch.object(WhatsAppDispatcher, "attach_via_desktop", side_effect=RuntimeError("boom"))
    def test_background_exception_never_escapes(self, _attach, _fallback):
        done = []
        t = self.dispatcher.auto_attach_pdf_in_background(self.pdf, on_complete=lambda ok, m: done.append(ok))
        t.join(5)
        self.assertEqual(done, [False])

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

    @patch.object(WhatsAppDispatcher, "auto_attach_pdf_in_background")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_desktop_with_phone_starts_background_attach(self, mock_open, mock_bg, _wb, _popen):
        cb = MagicMock()
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME FUELS", "9876543210", "105", "SEP 2026", self.pdf, target="desktop", on_complete=cb)
        self.assertTrue(ok)
        self.assertTrue(pending)
        url = mock_open.call_args[0][0]
        self.assertTrue(url.startswith("whatsapp://send?phone=919876543210&text="))
        self.assertIn("Note%20%3A%20please%20complete%20the%20payment", url)
        mock_bg.assert_called_once()
        self.assertEqual(mock_bg.call_args[0][0], os.path.abspath(self.pdf))
        self.assertIs(mock_bg.call_args.kwargs["on_complete"], cb)
        self.assertNotIn("attached", msg.lower().replace("attaching", ""))  # never claims success up front

    @patch.object(WhatsAppDispatcher, "auto_attach_pdf_in_background")
    @patch.object(WhatsAppDispatcher, "_open_url", return_value=True)
    def test_desktop_without_phone_does_not_paste_blind(self, _open, mock_bg, _wb, _popen):
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME", None, "1", "M", self.pdf, target="desktop")
        self.assertTrue(ok)
        self.assertFalse(pending)
        mock_bg.assert_not_called()
        self.assertIn("Ctrl+V", msg)

    @patch.object(WhatsAppDispatcher, "run_fallback", return_value="hint")
    @patch.object(WhatsAppDispatcher, "auto_attach_pdf_in_background")
    @patch.object(WhatsAppDispatcher, "_open_url", side_effect=[False, True])
    def test_desktop_not_installed_falls_back_to_web(self, mock_open, mock_bg, mock_fallback, _wb, _popen):
        ok, msg, pending = self.dispatcher.share_invoice_to_whatsapp(
            "ACME", "9876543210", "1", "M", self.pdf, target="desktop")
        self.assertTrue(ok)
        self.assertFalse(pending)
        self.assertTrue(mock_open.call_args_list[1][0][0].startswith("https://web.whatsapp.com/send"))
        mock_bg.assert_not_called()
        mock_fallback.assert_called_once()

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
        with patch.object(WhatsAppDispatcher, "auto_attach_pdf_in_background"):
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
