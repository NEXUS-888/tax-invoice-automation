import os
import sys
import time
import threading
import webbrowser
import urllib.parse
import subprocess

try:
    from PyQt6.QtWidgets import QApplication
    from PyQt6.QtCore import QMimeData, QUrl
except ImportError:
    QApplication = None
    QMimeData = None
    QUrl = None

from whatsapp_automator import build_invoice_message


# Executable basenames of the native WhatsApp Desktop app (Store / WinUI and WebView2 builds).
WHATSAPP_EXE_PREFIX = "whatsapp"


class WhatsAppDispatcher:
    """
    Interactive "Share App" dispatcher (used when no Playwright session is connected).

    Strategy, in order of reliability:
      1. Connected WhatsApp Web session -> handled by WhatsAppAutomator.prepare_invoice_share
         (routed by the GUI, not this class).
      2. Native WhatsApp Desktop -> open whatsapp://send, wait for the WhatsApp *process* window,
         force it to the foreground, verify focus, then inject Ctrl+V of a CF_HDROP clipboard.
         Every step that can fail is checked; any failure triggers the fallback.
      3. Fallback -> PDF copied to clipboard + pre-selected in Explorer for a 1-second drag/paste.
    """

    def __init__(self, mode="app"):
        self.mode = mode

    # ── Formatting ───────────────────────────────────────────────

    def format_phone_number(self, phone):
        """Clean phone number string to international standard e.g. 919876543210."""
        digits = ''.join(c for c in str(phone) if c.isdigit())
        if not digits:
            return ""
        if len(digits) == 10 and digits[0] in '6789':
            return f"91{digits}"
        if len(digits) == 12 and digits.startswith('91') and digits[2] in '6789':
            return digits
        return ""

    def build_invoice_message(self, agency_name, invoice_no, month_desc):
        """Builds standard invoice text message with payment reminder note."""
        return build_invoice_message(agency_name, invoice_no, month_desc)

    def build_share_url(self, target_mode, clean_phone, message):
        """Returns the URL to open for the given target ('app'/'desktop', 'web', 'auto')."""
        encoded_msg = urllib.parse.quote(message)
        if target_mode in ("app", "desktop"):
            if clean_phone:
                return f"whatsapp://send?phone={clean_phone}&text={encoded_msg}"
            return f"whatsapp://send?text={encoded_msg}"
        if target_mode == "auto":
            if clean_phone:
                return f"https://api.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}"
            return f"https://api.whatsapp.com/send?text={encoded_msg}"
        if clean_phone:
            return f"https://web.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}"
        return "https://web.whatsapp.com/"

    # ── Clipboard ────────────────────────────────────────────────

    def _set_clipboard_hdrop(self, abs_path, retries=5):
        """Writes a CF_HDROP (file list) to the Windows clipboard. Retries if another app holds it open."""
        import ctypes
        from ctypes import wintypes

        CF_HDROP = 15
        GHND = 0x0042

        class DROPFILES(ctypes.Structure):
            _fields_ = [
                ("pFiles", wintypes.DWORD),
                ("pt", wintypes.POINT),
                ("fNC", wintypes.BOOL),
                ("fWide", wintypes.BOOL),
            ]

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalLock.restype = wintypes.LPVOID
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalFree.restype = wintypes.HGLOBAL
        user32.OpenClipboard.argtypes = [wintypes.HWND]
        user32.OpenClipboard.restype = wintypes.BOOL
        user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
        user32.SetClipboardData.restype = wintypes.HANDLE

        path_bytes = (abs_path + "\0\0").encode("utf-16le")
        dropfiles = DROPFILES()
        dropfiles.pFiles = ctypes.sizeof(DROPFILES)
        dropfiles.fWide = True

        h_global = kernel32.GlobalAlloc(GHND, ctypes.sizeof(DROPFILES) + len(path_bytes))
        if not h_global:
            return False
        ptr = kernel32.GlobalLock(h_global)
        if not ptr:
            kernel32.GlobalFree(h_global)
            return False
        ctypes.memmove(ptr, ctypes.byref(dropfiles), ctypes.sizeof(DROPFILES))
        ctypes.memmove(ptr + ctypes.sizeof(DROPFILES), path_bytes, len(path_bytes))
        kernel32.GlobalUnlock(h_global)

        for _ in range(retries):
            if user32.OpenClipboard(None):
                try:
                    user32.EmptyClipboard()
                    if user32.SetClipboardData(CF_HDROP, h_global):
                        return True  # clipboard now owns the memory
                finally:
                    user32.CloseClipboard()
                break
            time.sleep(0.05)

        kernel32.GlobalFree(h_global)
        return False

    def copy_pdf_to_clipboard(self, pdf_path):
        """
        Copies the PDF file (not its contents) to the clipboard so Ctrl+V in WhatsApp attaches it.
        Safe to call from any thread: the Qt fallback is only used on the GUI thread.
        """
        abs_path = os.path.abspath(pdf_path) if pdf_path else ""
        if not abs_path or not os.path.exists(abs_path):
            return False

        if sys.platform == "win32":
            try:
                if self._set_clipboard_hdrop(abs_path):
                    return True
            except Exception:
                pass

            try:
                flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                ps_path = abs_path.replace("'", "''")
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command", f"Set-Clipboard -LiteralPath '{ps_path}'"],
                    check=True, creationflags=flags, timeout=10,
                )
                return True
            except Exception:
                pass

        # Qt clipboard must only be touched from the GUI thread
        if QApplication and threading.current_thread() is threading.main_thread():
            try:
                app = QApplication.instance()
                if app:
                    mime = QMimeData()
                    mime.setUrls([QUrl.fromLocalFile(abs_path)])
                    app.clipboard().setMimeData(mime)
                    return True
            except Exception:
                pass

        return False

    # ── Fallback: Explorer pre-selection ─────────────────────────

    def reveal_pdf_in_file_manager(self, abs_pdf):
        """Opens the file manager with the PDF pre-selected. Returns True if launched."""
        try:
            if sys.platform == 'win32':
                subprocess.Popen(f'explorer /select,"{abs_pdf}"')
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', '-R', abs_pdf])
            else:
                subprocess.Popen(['xdg-open', os.path.dirname(abs_pdf)])
            return True
        except Exception:
            return False

    def run_fallback(self, abs_pdf, reason=""):
        """Clipboard + Explorer pre-selection. Returns a user-facing hint message."""
        copied = self.copy_pdf_to_clipboard(abs_pdf)
        self.reveal_pdf_in_file_manager(abs_pdf)
        name = os.path.basename(abs_pdf)
        prefix = f"{reason} " if reason else ""
        if copied:
            return (f"{prefix}{name} is copied and highlighted in Explorer — "
                    f"press Ctrl+V in the chat or drag it in.")
        return f"{prefix}{name} is highlighted in Explorer — drag it into the chat."

    # ── Native WhatsApp Desktop window automation (ctypes only) ──

    def _process_image_name(self, pid):
        """Returns the lowercase executable basename for a process id, or '' on failure."""
        import ctypes
        from ctypes import wintypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        kernel32.OpenProcess.restype = wintypes.HANDLE
        h_proc = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h_proc:
            return ""
        try:
            buf = ctypes.create_unicode_buffer(1024)
            size = wintypes.DWORD(len(buf))
            if kernel32.QueryFullProcessImageNameW(h_proc, 0, buf, ctypes.byref(size)):
                return os.path.basename(buf.value).lower()
            return ""
        finally:
            kernel32.CloseHandle(h_proc)

    def find_whatsapp_desktop_windows(self):
        """
        Finds visible top-level windows owned by the native WhatsApp Desktop process.
        Matching on the process executable (not the window title) excludes browser tabs,
        the Playwright Chromium window, and this app's own "...WhatsApp Dispatcher" window.
        Returns a list of (hwnd, title, exe_name).
        """
        if sys.platform != "win32":
            return []
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            GW_OWNER = 4

            found = []

            def enum_cb(hwnd, _):
                if not user32.IsWindowVisible(hwnd) or user32.GetWindow(hwnd, GW_OWNER):
                    return True
                pid = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                exe = self._process_image_name(pid.value)
                if exe.startswith(WHATSAPP_EXE_PREFIX):
                    buf = ctypes.create_unicode_buffer(512)
                    user32.GetWindowTextW(hwnd, buf, 512)
                    if buf.value:
                        found.append((hwnd, buf.value, exe))
                return True

            WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
            user32.EnumWindows(WNDENUMPROC(enum_cb), 0)
            return found
        except Exception:
            return []

    def force_foreground(self, hwnd):
        """Restores and activates hwnd; returns True only if it really became the foreground window."""
        if sys.platform != "win32":
            return False
        try:
            import ctypes
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            SW_RESTORE = 9

            if not user32.IsWindow(hwnd):
                return False
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, SW_RESTORE)

            fg_thread = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
            our_thread = kernel32.GetCurrentThreadId()
            attached = bool(fg_thread and fg_thread != our_thread and
                            user32.AttachThreadInput(our_thread, fg_thread, True))
            try:
                user32.BringWindowToTop(hwnd)
                user32.SetForegroundWindow(hwnd)
            finally:
                if attached:
                    user32.AttachThreadInput(our_thread, fg_thread, False)

            for _ in range(10):
                if user32.GetForegroundWindow() == hwnd:
                    return True
                time.sleep(0.05)
            return False
        except Exception:
            return False

    def send_ctrl_v(self):
        """
        Injects Ctrl+V with SendInput. Returns False if Windows rejected the input
        (e.g. UIPI blocks injection into an elevated window, or the desktop is locked).
        """
        if sys.platform != "win32":
            return False
        try:
            import ctypes
            from ctypes import wintypes

            ULONG_PTR = ctypes.c_size_t
            INPUT_KEYBOARD = 1
            KEYEVENTF_KEYUP = 0x0002
            VK_CONTROL, VK_V = 0x11, 0x56

            class KEYBDINPUT(ctypes.Structure):
                _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                            ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                            ("dwExtraInfo", ULONG_PTR)]

            class MOUSEINPUT(ctypes.Structure):
                _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                            ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                            ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]

            class _INPUTUNION(ctypes.Union):
                # MOUSEINPUT is the largest member; it fixes sizeof(INPUT) for SendInput
                _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT)]

            class INPUT(ctypes.Structure):
                _fields_ = [("type", wintypes.DWORD), ("u", _INPUTUNION)]

            def key(vk, up=False):
                inp = INPUT(type=INPUT_KEYBOARD)
                inp.u.ki = KEYBDINPUT(wVk=vk, dwFlags=KEYEVENTF_KEYUP if up else 0)
                return inp

            seq = (INPUT * 4)(key(VK_CONTROL), key(VK_V), key(VK_V, True), key(VK_CONTROL, True))
            sent = ctypes.windll.user32.SendInput(4, seq, ctypes.sizeof(INPUT))
            return sent == 4
        except Exception:
            return False

    def attach_via_desktop(self, abs_pdf, window_timeout=15.0, warm_settle=2.0, cold_settle=5.0,
                           was_running=None):
        """
        Blocking: waits for WhatsApp Desktop, focuses it and pastes the PDF.
        Must run on a background thread. Returns (ok, message).
        """
        if sys.platform != "win32":
            return False, "Native WhatsApp Desktop automation is only available on Windows."

        deadline = time.time() + window_timeout
        hwnd = None
        while time.time() < deadline:
            wins = self.find_whatsapp_desktop_windows()
            if wins:
                hwnd = wins[0][0]
                break
            time.sleep(0.3)
        if not hwnd:
            return False, "WhatsApp Desktop window did not appear."

        # A cold start needs longer to log in and open the chat than a warm chat switch.
        time.sleep(warm_settle if was_running else cold_settle)

        # Clipboard may have been changed by the user while waiting; re-assert it.
        if not self.copy_pdf_to_clipboard(abs_pdf):
            return False, "Could not place the PDF on the clipboard."
        if not self.force_foreground(hwnd):
            return False, "Windows did not allow WhatsApp to be brought to the front."
        time.sleep(0.3)
        if not self.send_ctrl_v():
            return False, "Windows blocked the paste keystroke (UIPI/locked desktop)."
        return True, "PDF pasted into WhatsApp Desktop — check the preview and press Send."

    def auto_attach_pdf_in_background(self, abs_pdf, on_complete=None, was_running=None):
        """Runs attach_via_desktop on a daemon thread; on failure runs the fallback. Returns the thread."""
        def worker():
            try:
                ok, msg = self.attach_via_desktop(abs_pdf, was_running=was_running)
            except Exception as e:
                ok, msg = False, f"Desktop attach error: {e}"
            if not ok:
                msg = self.run_fallback(abs_pdf, reason=f"Auto-attach failed ({msg})")
            if on_complete:
                try:
                    on_complete(ok, msg)
                except Exception:
                    pass

        t = threading.Thread(target=worker, daemon=True, name="wa-desktop-attach")
        t.start()
        return t

    # ── Entry point ──────────────────────────────────────────────

    def _open_url(self, url):
        if hasattr(os, 'startfile'):
            try:
                os.startfile(url)
                return True
            except Exception:
                pass
        try:
            return bool(webbrowser.open(url))
        except Exception:
            return False

    def share_invoice_to_whatsapp(self, agency_name, phone=None, invoice_no="", month_desc="", pdf_path="",
                                  target=None, open_explorer=False, on_complete=None):
        """
        Opens WhatsApp with the pre-filled invoice message and gets the PDF into the chat.

        Returns (ok, message, pending):
          pending=True  -> a background attach is running; on_complete(attached, message) fires later
                           (from a worker thread — marshal to the GUI thread before touching widgets).
          pending=False -> finished synchronously; message describes what the user should do.
        """
        abs_pdf = os.path.abspath(pdf_path) if pdf_path else ""
        has_pdf = bool(abs_pdf) and os.path.exists(abs_pdf)

        message = self.build_invoice_message(agency_name, invoice_no, month_desc)
        clean_phone = self.format_phone_number(phone) if phone else ""
        target_mode = (target or self.mode or "app").lower()
        target_desc = f"{agency_name} ({clean_phone})" if clean_phone else agency_name

        if target_mode in ("app", "desktop"):
            was_running = bool(self.find_whatsapp_desktop_windows())
            if has_pdf:
                self.copy_pdf_to_clipboard(abs_pdf)
            if not self._open_url(self.build_share_url("app", clean_phone, message)):
                # WhatsApp Desktop not installed -> WhatsApp Web in the default browser
                target_mode = "web"
            elif has_pdf and clean_phone:
                self.auto_attach_pdf_in_background(abs_pdf, on_complete=on_complete, was_running=was_running)
                return True, f"Opened WhatsApp Desktop for {target_desc} — attaching PDF...", True
            elif has_pdf:
                return True, (f"Opened WhatsApp Desktop — pick a contact, then press Ctrl+V "
                              f"to attach {os.path.basename(abs_pdf)}."), False
            else:
                return True, f"Opened WhatsApp Desktop for {target_desc}.", False

        opened = self._open_url(self.build_share_url(target_mode, clean_phone, message))
        dest_label = "WhatsApp Web" if target_mode == "web" else "WhatsApp"
        if not opened:
            hint = self.run_fallback(abs_pdf) if has_pdf else ""
            return False, f"Could not open {dest_label}. {hint}".strip(), False

        # A tab in the user's own browser cannot be driven safely, so hand over the PDF ready to drop.
        if has_pdf:
            return True, f"Opened {dest_label} for {target_desc}. {self.run_fallback(abs_pdf)}", False
        return True, f"Opened {dest_label} for {target_desc}.", False

    # Backwards compatibility alias
    def send_agency_invoice(self, agency_name, phone, invoice_no, month_desc, pdf_path):
        ok, msg, _ = self.share_invoice_to_whatsapp(agency_name, phone, invoice_no, month_desc, pdf_path)
        return ok, msg
