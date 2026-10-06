import os
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


class WhatsAppDispatcher:
    """
    Direct WhatsApp App Share Dispatcher.
    Opens native WhatsApp Desktop app (or Web fallback) with pre-filled message.
    Automatically copies PDF file to Windows Clipboard (CF_HDROP) for instant Ctrl+V pasting.
    """

    def __init__(self, mode="app"):
        self.mode = mode

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

    def copy_pdf_to_clipboard(self, pdf_path):
        """
        Copies the PDF invoice file to Windows Clipboard (CF_HDROP).
        Pressing Ctrl+V in WhatsApp instantly attaches the PDF as a document.
        """
        abs_path = os.path.abspath(pdf_path)
        if not os.path.exists(abs_path):
            return False
        try:
            if QApplication:
                app = QApplication.instance()
                if app:
                    mime = QMimeData()
                    mime.setUrls([QUrl.fromLocalFile(abs_path)])
                    app.clipboard().setMimeData(mime)
                    return True
        except Exception:
            pass
        return False

    def build_invoice_message(self, agency_name, invoice_no, month_desc):
        """Builds standard invoice text message with payment reminder note."""
        return (
            f"Dear {agency_name},\n\n"
            f"Please find attached your invoice (Invoice No: {invoice_no}) for {month_desc}.\n\n"
            f"Note : please complete the payment before 10th of this month.\n\n"
            f"Thank you,\nANANYA ENTERPRISES"
        )

    def share_invoice_to_whatsapp(self, agency_name, phone=None, invoice_no="", month_desc="", pdf_path="", target=None, open_explorer=True):
        """
        Opens WhatsApp with pre-filled invoice message.
        - target: 'desktop' / 'app' (opens native WhatsApp Desktop App),
                  'web' (opens WhatsApp Web in default browser - 100% reliable),
                  'auto' (opens WhatsApp universal link api.whatsapp.com).
                  If None, falls back to self.mode (default: 'app').
        - If phone is given: opens chat directly for that number.
        - If phone is None/empty: opens WhatsApp with contact selector so user can pick any chat.
        - Automatically copies the PDF invoice to Windows Clipboard (CF_HDROP) for instant Ctrl+V attachment.
        - Highlights the PDF in Windows Explorer for drag & drop backup.
        """
        abs_pdf = os.path.abspath(pdf_path) if pdf_path else ""
        if abs_pdf and os.path.exists(abs_pdf):
            self.copy_pdf_to_clipboard(abs_pdf)

        message = self.build_invoice_message(agency_name, invoice_no, month_desc)
        encoded_msg = urllib.parse.quote(message)
        clean_phone = self.format_phone_number(phone) if phone else ""

        opened = False
        target_mode = (target or self.mode or "app").lower()

        if target_mode in ("app", "desktop"):
            if clean_phone:
                wa_url = f"whatsapp://send?phone={clean_phone}&text={encoded_msg}"
            else:
                wa_url = f"whatsapp://send?text={encoded_msg}"
            try:
                os.startfile(wa_url)
                opened = True
            except Exception:
                try:
                    webbrowser.open(wa_url)
                    opened = True
                except Exception:
                    pass
            # If native desktop protocol fails, fallback to web
            if not opened:
                wa_fallback_url = f"https://web.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}" if clean_phone else "https://web.whatsapp.com/"
                try:
                    webbrowser.open(wa_fallback_url)
                    opened = True
                except Exception:
                    pass
        elif target_mode == "auto":
            if clean_phone:
                wa_url = f"https://api.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}"
            else:
                wa_url = f"https://api.whatsapp.com/send?text={encoded_msg}"
            try:
                webbrowser.open(wa_url)
                opened = True
            except Exception:
                pass
        else: # "web" (WhatsApp Web in browser)
            if clean_phone:
                wa_url = f"https://web.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}"
            else:
                wa_url = "https://web.whatsapp.com/"
            try:
                webbrowser.open(wa_url)
                opened = True
            except Exception:
                pass

        # Highlight PDF in Explorer so user can also drag & drop
        if open_explorer and abs_pdf and os.path.exists(abs_pdf):
            try:
                subprocess.Popen(f'explorer /select,"{abs_pdf}"')
            except Exception:
                pass

        target_desc = f"{agency_name} ({clean_phone})" if clean_phone else agency_name
        dest_label = "WhatsApp Web" if target_mode == "web" else ("WhatsApp App" if target_mode in ("app", "desktop") else "WhatsApp")
        return True, f"Opened {dest_label} for {target_desc}. PDF copied to clipboard — press Ctrl+V to attach."

    # Backwards compatibility alias
    def send_agency_invoice(self, agency_name, phone, invoice_no, month_desc, pdf_path):
        return self.share_invoice_to_whatsapp(agency_name, phone, invoice_no, month_desc, pdf_path)
