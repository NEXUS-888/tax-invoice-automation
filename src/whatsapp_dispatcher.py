import os
import webbrowser
import urllib.parse
import subprocess

class WhatsAppDispatcher:
    def __init__(self, mode="web"):
        """
        WhatsApp Dispatcher
        mode: 'web' (WhatsApp Web automation / deep linking) or 'api'
        """
        self.mode = mode

    def format_phone_number(self, phone):
        """Clean phone number string to international standard e.g. +919876543210."""
        digits = ''.join(c for c in str(phone) if c.isdigit())
        if not digits:
            return ""
        if len(digits) == 10 and digits[0] in '6789':
            return f"91{digits}"
        if len(digits) == 12 and digits.startswith('91') and digits[2] in '6789':
            return digits
        return ""

    def generate_whatsapp_link(self, phone, message):
        """Generates a WhatsApp Web / App deep link URL."""
        clean_phone = self.format_phone_number(phone)
        encoded_msg = urllib.parse.quote(message)
        query = urllib.parse.urlencode({'phone': clean_phone, 'text': message})
        return f"https://api.whatsapp.com/send?{query}"

    def send_agency_invoice(self, agency_name, phone, invoice_no, month_desc, pdf_path):
        """
        Prepares and triggers WhatsApp dispatch for an agency.
        Opens WhatsApp Web with pre-filled message and copies the PDF path / opens the folder.
        """
        clean_phone = self.format_phone_number(phone)
        if not clean_phone:
            return False, "Invalid or missing phone number"

        msg = (
            f"Dear {agency_name},\n\n"
            f"Please find attached your invoice (Invoice No: {invoice_no}) for {month_desc}.\n\n"
            f"Thank you,\nANANYA ENTERPRISES"
        )

        wa_url = self.generate_whatsapp_link(clean_phone, msg)
        
        # Open WhatsApp Web in default browser
        webbrowser.open(wa_url)

        # Highlight PDF file in Windows Explorer so user can drag & drop it directly into WhatsApp
        if os.path.exists(pdf_path):
            try:
                subprocess.Popen(f'explorer /select,"{os.path.abspath(pdf_path)}"')
            except Exception:
                pass

        return True, f"Opened WhatsApp chat for {clean_phone}"
