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
from contacts_manager import normalize_phone
from whatsapp_desktop_bridge import WhatsAppDesktopBridge


# Executable basenames of the native WhatsApp Desktop app (Store / WinUI and WebView2 builds).
WHATSAPP_EXE_PREFIX = "whatsapp"

_DESKTOP_LOCK = threading.Lock()


def share_helper_path():
    """ShareInvoice.exe (Windows Share helper): bundled under _MEIPASS when frozen, else in the repo's tools/."""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "tools", "share_invoice", "ShareInvoice.exe")


class WhatsAppDispatcher:
    """
    Interactive "Share App" dispatcher (used when no Playwright session is connected).

    Native WhatsApp Desktop (2025+ = web.whatsapp.com in WebView2): the chat is opened and the PDF
    attached inside WhatsApp's own page via WhatsAppDesktopBridge, with the invoice message as the
    caption; the user presses Send. Any failure falls back to: PDF copied to the clipboard and
    pre-selected in Explorer for a one-second drag/paste.
    """

    def __init__(self, mode="app"):
        self.mode = mode

    # ── Formatting ───────────────────────────────────────────────

    def format_phone_number(self, phone):
        """Clean phone number string to international standard e.g. 919876543210."""
        return normalize_phone(phone)

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
        """Asks Windows to activate hwnd; True only if it really became the foreground window."""
        if sys.platform != "win32":
            return False
        try:
            import ctypes
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            if not user32.IsWindow(hwnd):
                return False
            if user32.GetForegroundWindow() == hwnd:
                return True
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

    # ── Desktop attach (inside WhatsApp's own page) ──────────────

    def attach_in_background(self, clean_phone, message, abs_pdf, on_complete=None):
        """
        Opens the chat inside WhatsApp Desktop (clean_phone None: the user picks it in WhatsApp's
        contact list) and attaches the PDF with the message as caption, on a daemon thread
        (see whatsapp_desktop_bridge). On failure the PDF is copied and
        highlighted in Explorer instead. on_complete(attached, message) fires from the thread.
        """
        def worker():
            with _DESKTOP_LOCK:  # one share at a time: they all drive the same WhatsApp window
                try:
                    bridge = WhatsAppDesktopBridge()
                    ok, msg = bridge.enable()
                    if ok:
                        ok, msg = bridge.attach_invoice(clean_phone, message, abs_pdf)
                except Exception as e:
                    ok, msg = False, f"WhatsApp Desktop error: {e}"
            if not ok:
                msg = self.run_fallback(abs_pdf, reason=f"Could not attach automatically ({msg}).")
            if on_complete:
                try:
                    on_complete(ok, msg)
                except Exception:
                    pass

        t = threading.Thread(target=worker, daemon=True, name="wa-desktop-attach")
        t.start()
        return t

    # ── Windows Share (user picks WhatsApp and the contact) ──────

    def _set_clipboard_text(self, text):
        """Puts plain text on the Windows clipboard (CF_UNICODETEXT). Safe from any thread."""
        if sys.platform != "win32":
            return False
        try:
            import ctypes
            from ctypes import wintypes
            user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
            kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
            kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
            kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
            kernel32.GlobalLock.restype = wintypes.LPVOID
            kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
            kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
            user32.OpenClipboard.argtypes = [wintypes.HWND]
            user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
            user32.SetClipboardData.restype = wintypes.HANDLE

            data = (text.replace("\r\n", "\n").replace("\n", "\r\n") + "\0").encode("utf-16le")
            h = kernel32.GlobalAlloc(0x0042, len(data))  # GHND
            ptr = kernel32.GlobalLock(h) if h else None
            if not ptr:
                if h:
                    kernel32.GlobalFree(h)
                return False
            ctypes.memmove(ptr, data, len(data))
            kernel32.GlobalUnlock(h)
            for _ in range(5):
                if user32.OpenClipboard(None):
                    try:
                        user32.EmptyClipboard()
                        if user32.SetClipboardData(13, h):  # CF_UNICODETEXT; clipboard now owns h
                            return True
                    finally:
                        user32.CloseClipboard()
                    break
                time.sleep(0.05)
            kernel32.GlobalFree(h)
        except Exception:
            pass
        return False

    def share_with_windows(self, abs_pdf, message, title, on_complete=None):
        """
        Opens the Windows Share window (ShareInvoice.exe) with the PDF and the invoice message. The user
        clicks WhatsApp, picks the contact and presses Send — Windows hands the file to WhatsApp, so
        nothing inside WhatsApp is automated. The message is also put on the clipboard in case the
        target app drops shared text. on_complete(shared, message) fires from a worker thread.
        Returns (started, message).
        """
        exe = share_helper_path()
        if sys.platform != "win32" or not os.path.exists(exe):
            return False, "The Windows Share helper (ShareInvoice.exe) is missing."

        import tempfile
        fd, msg_file = tempfile.mkstemp(prefix="invoice_message_", suffix=".txt")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(message)
        self._set_clipboard_text(message)

        try:
            proc = subprocess.Popen([exe, abs_pdf, msg_file, title], stdout=subprocess.PIPE,
                                    stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except OSError as e:
            os.remove(msg_file)
            return False, f"Could not start Windows Share: {e}"
        try:
            import ctypes
            # We are the foreground app (the user just clicked), so let the helper come to the front:
            # Windows only shows the Share window for the foreground window.
            ctypes.windll.user32.AllowSetForegroundWindow(proc.pid)
        except Exception:
            pass

        def worker():
            try:
                out, _ = proc.communicate(timeout=360)
                outcome = (out or b"").decode("utf-8", "replace").strip().splitlines()[-1:] or [""]
                outcome = outcome[0]
            except subprocess.TimeoutExpired:
                proc.kill()
                outcome = "CANCELLED"
            except Exception as e:
                outcome = f"ERROR:{e}"
            finally:
                try:
                    os.remove(msg_file)
                except OSError:
                    pass

            if outcome.startswith("TARGET:"):
                app = outcome[len("TARGET:"):].strip() or "the app"
                ok, msg = True, (f"Shared to {app} — pick the contact and press Send. If the message is "
                                 f"not filled in, press Ctrl+V in the caption (it is on the clipboard).")
            elif outcome.startswith("ERROR:"):
                ok, msg = False, self.run_fallback(abs_pdf, reason=f"Windows Share failed ({outcome[6:]}).")
            else:
                ok, msg = False, "Share was cancelled."
            if on_complete:
                try:
                    on_complete(ok, msg)
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True, name="windows-share").start()
        return True, "Windows Share is open — click WhatsApp, pick the contact and press Send."

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
        if phone and not clean_phone:
            # Opening WhatsApp without the number shows its "Send to" picker, where the PDF cannot be
            # attached automatically — refuse instead of silently degrading.
            return False, (f"'{phone}' is not a valid WhatsApp number for {agency_name}. "
                           f"Fix it in the WhatsApp Contacts tab."), False
        target_mode = (target or self.mode or "app").lower()
        target_desc = f"{agency_name} ({clean_phone})" if clean_phone else agency_name

        if target_mode in ("app", "desktop") and has_pdf and not clean_phone:
            # "Select Contact": Windows Share hands the PDF to WhatsApp, the user picks the contact.
            started, msg = self.share_with_windows(abs_pdf, message, f"Invoice {invoice_no}".strip(),
                                                   on_complete=on_complete)
            if started:
                return True, msg, True
            return False, self.run_fallback(abs_pdf, reason=msg), False

        if target_mode in ("app", "desktop") and has_pdf:
            # No whatsapp:// link here: in the 2025+ app it opens a "Send to" picker and the
            # text ends up sent separately from the PDF.
            self.attach_in_background(clean_phone, message, abs_pdf, on_complete=on_complete)
            return True, f"Opening {target_desc} in WhatsApp Desktop and attaching the PDF...", True

        if target_mode in ("app", "desktop"):
            if has_pdf:
                self.copy_pdf_to_clipboard(abs_pdf)
            if not self._open_url(self.build_share_url("app", clean_phone, message)):
                # WhatsApp Desktop not installed -> WhatsApp Web in the default browser
                target_mode = "web"
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
