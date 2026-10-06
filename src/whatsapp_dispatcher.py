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


# Executable basenames of the native WhatsApp Desktop app (Store / WinUI and WebView2 builds).
WHATSAPP_EXE_PREFIX = "whatsapp"


class WhatsAppDispatcher:
    """
    Interactive "Share App" dispatcher (used when no Playwright session is connected).

    Native WhatsApp Desktop: open whatsapp://send (chat + pre-filled message), wait until the
    WhatsApp *process* window is really on screen and has finished loading, click into the
    "Type a message" box, paste the PDF (CF_HDROP clipboard + Ctrl+V) and confirm on screen that
    the document preview opened. Any failed step falls back to: PDF copied to the clipboard and
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

    # ── Geometry & screen helpers (physical pixels) ─────────────
    #
    # WhatsApp Desktop (WinUI 3 + WebView2) exposes nothing inside its chat to UI Automation, and
    # ShowWindow/IsIconic report it "restored" while it is still parked off-screen at -32000.
    # So readiness and success are judged from what is actually on screen.

    # Centre of the "Type a message" box: 40 px (at 96 DPI) above the window's visible bottom edge,
    # 75% across — right of the clip/emoji buttons and left of the mic button, in both
    # the two-pane and the narrow single-pane layouts.
    COMPOSER_BOTTOM_OFFSET = 40
    COMPOSER_X_FRACTION = 0.75
    # Share of sampled pixels that must change for the document preview to count as open.
    PREVIEW_CHANGED = 0.15
    # Below this the window is considered unchanged (cursor blink, clock, etc.).
    UNCHANGED = 0.01
    SAMPLE = 64

    def _use_physical_pixels(self):
        """Make this thread's window/cursor coordinates physical pixels, whatever the process DPI mode."""
        try:
            import ctypes
            ctypes.windll.user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(-4))  # PER_MONITOR_AWARE_V2
        except Exception:
            pass

    def window_on_screen_rect(self, hwnd):
        """Visible bounds (l, t, r, b) if hwnd is really shown on a monitor, else None."""
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            if not user32.IsWindow(hwnd) or not user32.IsWindowVisible(hwnd) or user32.IsIconic(hwnd):
                return None
            dwm = ctypes.windll.dwmapi
            cloaked = wintypes.DWORD()
            if dwm.DwmGetWindowAttribute(wintypes.HWND(hwnd), 14, ctypes.byref(cloaked), 4) == 0 and cloaked.value:
                return None  # DWMWA_CLOAKED: hidden by the shell (other virtual desktop, suspended app)
            r = wintypes.RECT()
            # 9 = DWMWA_EXTENDED_FRAME_BOUNDS: the visible frame, without the invisible resize border
            if dwm.DwmGetWindowAttribute(wintypes.HWND(hwnd), 9, ctypes.byref(r), ctypes.sizeof(r)) != 0:
                if not user32.GetWindowRect(hwnd, ctypes.byref(r)):
                    return None
            if r.right - r.left < 300 or r.bottom - r.top < 300:
                return None
            vx, vy = user32.GetSystemMetrics(76), user32.GetSystemMetrics(77)
            vw, vh = user32.GetSystemMetrics(78), user32.GetSystemMetrics(79)
            if r.right <= vx or r.bottom <= vy or r.left >= vx + vw or r.top >= vy + vh:
                return None  # e.g. the -32000 "minimized" parking position
            return (r.left, r.top, r.right, r.bottom)
        except Exception:
            return None

    def composer_point(self, hwnd, rect):
        """Screen point inside WhatsApp's message box for a window with visible bounds `rect`."""
        try:
            import ctypes
            dpi = ctypes.windll.user32.GetDpiForWindow(hwnd) or 96
        except Exception:
            dpi = 96
        left, top, right, bottom = rect
        x = left + int((right - left) * self.COMPOSER_X_FRACTION)
        y = bottom - int(self.COMPOSER_BOTTOM_OFFSET * dpi / 96)
        return x, y

    def window_owns_point(self, hwnd, x, y):
        """True if the top-most window at (x, y) belongs to hwnd (nothing is covering it)."""
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            user32.WindowFromPoint.argtypes = [wintypes.POINT]
            user32.WindowFromPoint.restype = wintypes.HWND
            user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
            user32.GetAncestor.restype = wintypes.HWND
            hit = user32.WindowFromPoint(wintypes.POINT(x, y))
            return bool(hit) and user32.GetAncestor(hit, 2) == hwnd  # GA_ROOT
        except Exception:
            return False

    def capture_signature(self, rect):
        """Downscaled BGRA snapshot of a screen rectangle (GDI only, no extra packages), or None."""
        try:
            import ctypes
            from ctypes import wintypes
            user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
            user32.GetDC.restype = wintypes.HANDLE
            user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HANDLE]
            gdi32.CreateCompatibleDC.argtypes = [wintypes.HANDLE]
            gdi32.CreateCompatibleDC.restype = wintypes.HANDLE
            gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_int]
            gdi32.CreateCompatibleBitmap.restype = wintypes.HANDLE
            gdi32.SelectObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
            gdi32.SelectObject.restype = wintypes.HANDLE
            gdi32.SetStretchBltMode.argtypes = [wintypes.HANDLE, ctypes.c_int]
            gdi32.DeleteObject.argtypes = [wintypes.HANDLE]
            gdi32.DeleteDC.argtypes = [wintypes.HANDLE]
            gdi32.StretchBlt.argtypes = ([wintypes.HANDLE] + [ctypes.c_int] * 4 + [wintypes.HANDLE]
                                         + [ctypes.c_int] * 4 + [wintypes.DWORD])
            gdi32.GetDIBits.argtypes = [wintypes.HANDLE, wintypes.HANDLE, wintypes.UINT, wintypes.UINT,
                                        ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT]

            class BITMAPINFOHEADER(ctypes.Structure):
                _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                            ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                            ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                            ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                            ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]

            left, top, right, bottom = rect
            n = self.SAMPLE
            screen_dc = user32.GetDC(None)
            if not screen_dc:
                return None
            mem_dc = gdi32.CreateCompatibleDC(screen_dc)
            bmp = gdi32.CreateCompatibleBitmap(screen_dc, n, n)
            old = gdi32.SelectObject(mem_dc, bmp)
            try:
                gdi32.SetStretchBltMode(mem_dc, 4)  # HALFTONE: average pixels instead of dropping them
                copied = gdi32.StretchBlt(mem_dc, 0, 0, n, n, screen_dc, left, top,
                                          right - left, bottom - top, 0x00CC0020)  # SRCCOPY
            finally:
                gdi32.SelectObject(mem_dc, old)  # a bitmap must be deselected before GetDIBits
            try:
                if not copied:
                    return None
                bih = BITMAPINFOHEADER(biSize=ctypes.sizeof(BITMAPINFOHEADER), biWidth=n, biHeight=-n,
                                       biPlanes=1, biBitCount=32)
                buf = ctypes.create_string_buffer(n * n * 4)
                if gdi32.GetDIBits(mem_dc, bmp, 0, n, buf, ctypes.byref(bih), 0) != n:
                    return None
                return buf.raw
            finally:
                gdi32.DeleteObject(bmp)
                gdi32.DeleteDC(mem_dc)
                user32.ReleaseDC(None, screen_dc)
        except Exception:
            return None

    @staticmethod
    def changed_fraction(before, after):
        """Share of pixels whose colour moved noticeably between two snapshots (0.0-1.0)."""
        if not before or not after or len(before) != len(after):
            return 0.0
        changed = 0
        for i in range(0, len(before), 4):
            if (abs(before[i] - after[i]) + abs(before[i + 1] - after[i + 1])
                    + abs(before[i + 2] - after[i + 2])) > 48:
                changed += 1
        return changed / (len(before) // 4)

    # ── Input helpers ────────────────────────────────────────────

    def _send_inputs(self, events):
        """SendInput wrapper. events: ('key', vk, is_up) or ('mouse', flags). True if all were accepted."""
        try:
            import ctypes
            from ctypes import wintypes

            ULONG_PTR = ctypes.c_size_t

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

            arr = (INPUT * len(events))()
            for i, ev in enumerate(events):
                if ev[0] == "key":
                    arr[i].type = 1  # INPUT_KEYBOARD
                    arr[i].u.ki = KEYBDINPUT(wVk=ev[1], dwFlags=0x0002 if ev[2] else 0)  # KEYEVENTF_KEYUP
                else:
                    arr[i].type = 0  # INPUT_MOUSE
                    arr[i].u.mi = MOUSEINPUT(dwFlags=ev[1])
            return ctypes.windll.user32.SendInput(len(events), arr, ctypes.sizeof(INPUT)) == len(events)
        except Exception:
            return False

    def send_ctrl_v(self):
        """
        Injects Ctrl+V with SendInput. Returns False if Windows rejected the input
        (e.g. UIPI blocks injection into an elevated window, or the desktop is locked).
        """
        if sys.platform != "win32":
            return False
        VK_CONTROL, VK_V = 0x11, 0x56
        return self._send_inputs([("key", VK_CONTROL, False), ("key", VK_V, False),
                                  ("key", VK_V, True), ("key", VK_CONTROL, True)])

    def click_at(self, x, y):
        """Real left click at (x, y), then puts the mouse cursor back where the user had it."""
        if sys.platform != "win32":
            return False
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            prev = wintypes.POINT()
            user32.GetCursorPos(ctypes.byref(prev))
            if not user32.SetCursorPos(x, y):
                return False
            time.sleep(0.05)
            ok = self._send_inputs([("mouse", 0x0002), ("mouse", 0x0004)])  # LEFTDOWN, LEFTUP
            time.sleep(0.1)
            user32.SetCursorPos(prev.x, prev.y)
            return ok
        except Exception:
            return False

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

    # ── Desktop attach flow ──────────────────────────────────────

    def wait_for_onscreen_window(self, timeout):
        """(hwnd, rect) of the WhatsApp window once it is really drawn on a monitor, else (None, None)."""
        start = time.time()
        reactivated = False
        while time.time() - start < timeout:
            wins = self.find_whatsapp_desktop_windows()
            shown = [(h, self.window_on_screen_rect(h)) for h, _title, _exe in wins]
            shown = [(h, r) for h, r in shown if r]
            if shown:
                return max(shown, key=lambda hr: (hr[1][2] - hr[1][0]) * (hr[1][3] - hr[1][1]))
            if wins and not reactivated and time.time() - start > 3:
                # Running but parked off-screen / in the tray: let WhatsApp itself show its window.
                self._open_url("whatsapp://")
                reactivated = True
            time.sleep(0.3)
        return None, None

    def wait_until_stable(self, rect, min_wait, max_wait):
        """Waits until the window stops changing (chat finished loading), between min_wait and max_wait s."""
        time.sleep(min_wait)
        deadline = time.time() + max(0.0, max_wait - min_wait)
        prev = self.capture_signature(rect)
        while time.time() < deadline:
            time.sleep(0.5)
            cur = self.capture_signature(rect)
            if prev is None or cur is None:
                return
            if self.changed_fraction(prev, cur) < self.UNCHANGED:
                return
            prev = cur

    def paste_and_verify(self, hwnd, rect, timeout=4.0):
        """
        Clicks into the message box, pastes, and watches the window for the document preview.
        Returns 'attached', 'unverified' (screen could not be read, or changed only a little),
        'nothing', or an error message.
        """
        x, y = self.composer_point(hwnd, rect)
        if not self.window_owns_point(hwnd, x, y):
            self.force_foreground(hwnd)
            time.sleep(0.3)
            if not self.window_owns_point(hwnd, x, y):
                return "Another window is covering WhatsApp."
        if not self.click_at(x, y):
            return "Windows blocked the mouse click into WhatsApp."
        time.sleep(0.4)
        if self.window_on_screen_rect(hwnd) is None:
            return "WhatsApp window was hidden."
        before = self.capture_signature(rect)
        if not self.send_ctrl_v():
            return "Windows blocked the paste keystroke (UIPI/locked desktop)."
        best = 0.0
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(0.4)
            best = max(best, self.changed_fraction(before, self.capture_signature(rect)))
            if best >= self.PREVIEW_CHANGED:
                return "attached"
        if before is None:
            return "unverified"
        return "nothing" if best < self.UNCHANGED else "unverified"

    def attach_via_desktop(self, abs_pdf, window_timeout=20.0, was_running=None):
        """
        Blocking (run on a background thread): wait for the WhatsApp window to be really on screen and
        idle, click into the message box, paste the PDF and confirm the preview appeared.
        Returns (ok, message).
        """
        if sys.platform != "win32":
            return False, "Native WhatsApp Desktop automation is only available on Windows."
        self._use_physical_pixels()

        hwnd, rect = self.wait_for_onscreen_window(window_timeout)
        if not hwnd:
            return False, "WhatsApp Desktop window did not appear on screen."

        # Opening the chat from a cold start takes longer than switching chats in a running app.
        if was_running:
            self.wait_until_stable(rect, min_wait=1.5, max_wait=8.0)
        else:
            self.wait_until_stable(rect, min_wait=4.0, max_wait=20.0)
        rect = self.window_on_screen_rect(hwnd) or rect

        for _attempt in range(2):
            # The user may have copied something else while we waited; re-assert the PDF.
            if not self.copy_pdf_to_clipboard(abs_pdf):
                return False, "Could not place the PDF on the clipboard."
            result = self.paste_and_verify(hwnd, rect)
            if result == "attached":
                return True, "PDF attached in WhatsApp Desktop — check the preview and press Send."
            if result == "unverified":
                return True, "PDF pasted into WhatsApp Desktop — check the chat and press Send."
            if result != "nothing":
                return False, result
            # Nothing happened at all (chat still loading or focus lost): give it a moment, retry once.
            time.sleep(1.5)
            rect = self.window_on_screen_rect(hwnd) or rect
        return False, "WhatsApp did not show the document preview."

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
        if phone and not clean_phone:
            # Opening WhatsApp without the number shows its "Send to" picker, where the PDF cannot be
            # attached automatically — refuse instead of silently degrading.
            return False, (f"'{phone}' is not a valid WhatsApp number for {agency_name}. "
                           f"Fix it in the WhatsApp Contacts tab."), False
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
