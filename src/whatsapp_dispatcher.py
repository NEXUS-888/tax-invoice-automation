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



class WhatsAppDispatcher:
    """
    Interactive "Share App" dispatcher.

    WhatsApp Desktop: the PDF file is copied to the clipboard and WhatsApp is opened with the invoice
    message (the chat directly when the number is known, otherwise WhatsApp's contact picker). The
    user presses Ctrl+V in the chat to attach the PDF and then Send. Nothing inside WhatsApp is
    automated, so it behaves the same on every laptop.
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
                                  target=None):
        """
        Copies the invoice PDF to the clipboard and opens WhatsApp with the invoice message:
          target 'app'/'desktop' (default) -> WhatsApp Desktop (falls back to WhatsApp Web if not installed)
          target 'web'                     -> WhatsApp Web in the default browser
          target 'auto'                    -> api.whatsapp.com link
        With a phone number the chat opens directly; without one WhatsApp shows its contact picker.
        The user presses Ctrl+V in the chat to attach the PDF. Returns (ok, message for the user).
        """
        abs_pdf = os.path.abspath(pdf_path) if pdf_path else ""
        has_pdf = bool(abs_pdf) and os.path.exists(abs_pdf)

        message = self.build_invoice_message(agency_name, invoice_no, month_desc)
        clean_phone = self.format_phone_number(phone) if phone else ""
        if phone and not clean_phone:
            # Opening WhatsApp without the number would silently fall back to the contact picker.
            return False, (f"'{phone}' is not a valid WhatsApp number for {agency_name}. "
                           f"Fix it in the WhatsApp Contacts tab.")
        target_mode = (target or self.mode or "app").lower()
        target_desc = f"{agency_name} ({clean_phone})" if clean_phone else agency_name

        copied = has_pdf and self.copy_pdf_to_clipboard(abs_pdf)
        opened = False
        if target_mode in ("app", "desktop"):
            opened = self._open_url(self.build_share_url("app", clean_phone, message))
            if not opened:
                target_mode = "web"  # WhatsApp Desktop not installed -> WhatsApp Web in the browser
        if not opened:
            opened = self._open_url(self.build_share_url(target_mode, clean_phone, message))
        dest_label = {"web": "WhatsApp Web", "auto": "WhatsApp"}.get(target_mode, "WhatsApp Desktop")

        if not opened:
            hint = self.run_fallback(abs_pdf) if has_pdf else ""
            return False, f"Could not open {dest_label}. {hint}".strip()
        if not has_pdf:
            return True, f"Opened {dest_label} for {target_desc}."
        if not copied:
            return True, f"Opened {dest_label} for {target_desc}. {self.run_fallback(abs_pdf, reason='Could not copy the PDF.')}"
        where = "in the chat" if clean_phone else "in the chat after picking the contact"
        return True, (f"Opened {dest_label} for {target_desc}. The PDF is copied — press Ctrl+V {where} "
                      f"to attach it, then Send.")

    # Backwards compatibility alias
    def send_agency_invoice(self, agency_name, phone, invoice_no, month_desc, pdf_path):
        return self.share_invoice_to_whatsapp(agency_name, phone, invoice_no, month_desc, pdf_path)
