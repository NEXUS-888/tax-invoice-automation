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
    RECT = (100, 100, 1300, 1000)

    def setUp(self):
        self.dispatcher = WhatsAppDispatcher()
        self.pdf = _make_pdf()
        sleep = patch("whatsapp_dispatcher.time.sleep")
        sleep.start()
        self.addCleanup(sleep.stop)

    def tearDown(self):
        os.remove(self.pdf)

    def _patch(self, **kw):
        mocks = {}
        for name, value in kw.items():
            p = patch.object(WhatsAppDispatcher, name, return_value=value)
            mocks[name] = p.start()
            self.addCleanup(p.stop)
        return mocks

    # ── attach_via_desktop orchestration ──

    def _patch_flow(self, paste_results):
        m = self._patch(wait_for_onscreen_window=(1234, self.RECT), wait_until_stable=None,
                        window_on_screen_rect=self.RECT, copy_pdf_to_clipboard=True)
        p = patch.object(WhatsAppDispatcher, "paste_and_verify", side_effect=list(paste_results))
        m["paste_and_verify"] = p.start()
        self.addCleanup(p.stop)
        return m

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_preview_detected_is_success(self):
        m = self._patch_flow(["attached"])
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf, was_running=True)
        self.assertTrue(ok)
        self.assertIn("attached", msg)
        m["paste_and_verify"].assert_called_once_with(1234, self.RECT)

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_nothing_happened_retries_once_then_fails(self):
        m = self._patch_flow(["nothing", "nothing"])
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf, was_running=True)
        self.assertFalse(ok)
        self.assertEqual(m["paste_and_verify"].call_count, 2)
        self.assertEqual(m["copy_pdf_to_clipboard"].call_count, 2)

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_retry_succeeds_on_second_paste(self):
        self._patch_flow(["nothing", "attached"])
        ok, _ = self.dispatcher.attach_via_desktop(self.pdf, was_running=False)
        self.assertTrue(ok)

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_hard_error_is_not_retried(self):
        m = self._patch_flow(["Another window is covering WhatsApp."])
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf, was_running=True)
        self.assertFalse(ok)
        self.assertIn("covering", msg)
        m["paste_and_verify"].assert_called_once()

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_window_never_on_screen(self):
        m = self._patch(wait_for_onscreen_window=(None, None), paste_and_verify="attached")
        ok, msg = self.dispatcher.attach_via_desktop(self.pdf)
        self.assertFalse(ok)
        self.assertIn("did not appear", msg)
        m["paste_and_verify"].assert_not_called()

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_cold_start_waits_longer(self):
        m = self._patch_flow(["attached"])
        self.dispatcher.attach_via_desktop(self.pdf, was_running=False)
        cold = m["wait_until_stable"].call_args.kwargs
        self.assertGreaterEqual(cold["min_wait"], 4.0)

    # ── paste_and_verify ──

    def _sig(self, value):
        return bytes([value, value, value, 255]) * (WhatsAppDispatcher.SAMPLE ** 2)

    def test_paste_clicks_message_box_then_detects_preview(self):
        m = self._patch(composer_point=(800, 950), window_owns_point=True, click_at=True,
                        window_on_screen_rect=self.RECT, send_ctrl_v=True)
        with patch.object(WhatsAppDispatcher, "capture_signature",
                          side_effect=[self._sig(20), self._sig(200)]):
            self.assertEqual(self.dispatcher.paste_and_verify(1234, self.RECT), "attached")
        m["click_at"].assert_called_once_with(800, 950)
        m["send_ctrl_v"].assert_called_once()

    def test_paste_with_no_screen_change_reports_nothing(self):
        self._patch(composer_point=(800, 950), window_owns_point=True, click_at=True,
                    window_on_screen_rect=self.RECT, send_ctrl_v=True)
        with patch.object(WhatsAppDispatcher, "capture_signature", return_value=self._sig(20)):
            self.assertEqual(self.dispatcher.paste_and_verify(1234, self.RECT, timeout=1.0), "nothing")

    def test_unreadable_screen_is_unverified_not_failure(self):
        self._patch(composer_point=(800, 950), window_owns_point=True, click_at=True,
                    window_on_screen_rect=self.RECT, send_ctrl_v=True, capture_signature=None)
        self.assertEqual(self.dispatcher.paste_and_verify(1234, self.RECT, timeout=1.0), "unverified")

    def test_covered_window_is_never_clicked(self):
        m = self._patch(composer_point=(800, 950), window_owns_point=False, force_foreground=False,
                        click_at=True, send_ctrl_v=True)
        result = self.dispatcher.paste_and_verify(1234, self.RECT)
        self.assertIn("covering", result)
        m["click_at"].assert_not_called()
        m["send_ctrl_v"].assert_not_called()

    def test_window_hidden_after_click_does_not_paste(self):
        m = self._patch(composer_point=(800, 950), window_owns_point=True, click_at=True,
                        window_on_screen_rect=None, send_ctrl_v=True)
        self.assertIn("hidden", self.dispatcher.paste_and_verify(1234, self.RECT))
        m["send_ctrl_v"].assert_not_called()

    def test_blocked_keystroke_reported(self):
        self._patch(composer_point=(800, 950), window_owns_point=True, click_at=True,
                    window_on_screen_rect=self.RECT, send_ctrl_v=False, capture_signature=self._sig(1))
        self.assertIn("blocked", self.dispatcher.paste_and_verify(1234, self.RECT))

    # ── geometry / window state ──

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_composer_point_is_inside_message_box(self):
        with patch("ctypes.windll.user32.GetDpiForWindow", return_value=144, create=True):
            x, y = self.dispatcher.composer_point(1234, (0, 0, 1200, 900))
        self.assertEqual((x, y), (900, 840))  # 75% across, 40 logical px (60 physical @150%) from bottom

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_composer_point_default_dpi(self):
        with patch("ctypes.windll.user32.GetDpiForWindow", return_value=0, create=True):
            x, y = self.dispatcher.composer_point(1234, (100, 50, 900, 650))
        self.assertEqual((x, y), (700, 610))

    def test_changed_fraction(self):
        a, b = self._sig(10), self._sig(10)
        self.assertEqual(WhatsAppDispatcher.changed_fraction(a, b), 0.0)
        self.assertEqual(WhatsAppDispatcher.changed_fraction(a, self._sig(200)), 1.0)
        half = self._sig(10)[: len(a) // 2] + self._sig(200)[len(a) // 2:]
        self.assertAlmostEqual(WhatsAppDispatcher.changed_fraction(a, half), 0.5)
        self.assertEqual(WhatsAppDispatcher.changed_fraction(None, a), 0.0)

    def test_wait_for_window_ignores_offscreen_and_reactivates(self):
        """A WhatsApp window parked at -32000 must not count; WhatsApp is re-activated via its URL."""
        self._patch(find_whatsapp_desktop_windows=[(1234, "WhatsApp", "whatsapp.root.exe")],
                    window_on_screen_rect=None)
        opened = self._patch(_open_url=True)["_open_url"]
        clock = iter(range(0, 100))
        with patch("whatsapp_dispatcher.time.time", side_effect=lambda: next(clock)):
            self.assertEqual(self.dispatcher.wait_for_onscreen_window(timeout=10), (None, None))
        opened.assert_called_once_with("whatsapp://")

    def test_wait_for_window_returns_largest_onscreen(self):
        self._patch(find_whatsapp_desktop_windows=[(1, "WhatsApp", "whatsapp.exe"), (2, "WhatsApp", "whatsapp.exe")])
        rects = {1: (0, 0, 400, 400), 2: (0, 0, 1200, 900)}
        with patch.object(WhatsAppDispatcher, "window_on_screen_rect", side_effect=lambda h: rects[h]):
            self.assertEqual(self.dispatcher.wait_for_onscreen_window(timeout=5), (2, rects[2]))

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_window_on_screen_rect_rejects_invalid(self):
        self.assertIsNone(self.dispatcher.window_on_screen_rect(0))

    @unittest.skipUnless(sys.platform == "win32", "Windows only")
    def test_capture_signature_real_gdi(self):
        sig = self.dispatcher.capture_signature((0, 0, 200, 200))
        self.assertIsNotNone(sig)
        self.assertEqual(len(sig), WhatsAppDispatcher.SAMPLE ** 2 * 4)

    # ── background thread ──

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
