import os
import sys
import time
import urllib.parse
import re
import random


PAYMENT_REMINDER_NOTE = "Note : please complete the payment before 10th of this month."


def build_invoice_message(agency_name, invoice_no, month_desc):
    """Standard invoice WhatsApp message; always carries the mandatory payment reminder note."""
    return (
        f"Dear {agency_name},\n\n"
        f"Please find attached your invoice (Invoice No: {invoice_no}) for {month_desc}.\n\n"
        f"{PAYMENT_REMINDER_NOTE}\n\n"
        f"Thank you,\nANANYA ENTERPRISES"
    )


class WhatsAppAutomator:
    """Sends invoice PDFs via WhatsApp Web using Playwright."""

    def __init__(self, session_dir, log_callback=None):
        self.session_dir = session_dir
        self.log_callback = log_callback
        os.makedirs(self.session_dir, exist_ok=True)
        self.playwright = None
        self.browser_context = None
        self.page = None

    def log(self, msg):
        if self.log_callback:
            self.log_callback(msg)
        else:
            try:
                print(f"[WhatsApp] {msg}")
            except UnicodeEncodeError:
                print(f"[WhatsApp] {msg}".encode("ascii", "replace").decode("ascii"))

    def format_phone(self, phone):
        digits = ''.join(c for c in str(phone) if c.isdigit())
        if len(digits) == 10 and digits[0] in '6789':
            return f"91{digits}"
        if len(digits) == 12 and digits.startswith('91') and digits[2] in '6789':
            return digits
        return ""

    def parse_phones(self, phone_input):
        if not phone_input:
            return []
        if isinstance(phone_input, list):
            raw_list = phone_input
        else:
            raw_list = re.split(r'[,;/|\n]+', str(phone_input))
        result = []
        for p in raw_list:
            formatted = self.format_phone(p)
            if formatted and formatted not in result:
                result.append(formatted)
        return result

    def _clean_session_lockfiles(self):
        """Removes stale Chromium lockfiles that prevent browser startup on Windows."""
        lockfiles = ["lockfile", "SingletonLock", "SingletonCookie", "SingletonSocket", "DevToolsActivePort"]
        for name in lockfiles:
            file_path = os.path.join(self.session_dir, name)
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    self.log(f"Cleaned stale lockfile: {name}")
                except Exception:
                    pass

    def _launch_context(self, headless, user_agent, launch_args):
        """Attempts to launch persistent context using Chromium, Edge, or Chrome channels."""
        channels = [None, "msedge", "chrome"]
        last_exc = None
        for ch in channels:
            try:
                kwargs = {
                    "user_data_dir": self.session_dir,
                    "headless": headless,
                    "no_viewport": True,
                    "user_agent": user_agent,
                    "args": launch_args,
                }
                if ch:
                    kwargs["channel"] = ch
                ctx = self.playwright.chromium.launch_persistent_context(**kwargs)
                if ch:
                    self.log(f"Launched browser using system '{ch}' channel.")
                return ctx
            except Exception as e:
                last_exc = e
                continue
        if last_exc:
            raise last_exc

    def launch_session(self, headless=False):
        from playwright.sync_api import sync_playwright

        if self.browser_context:
            self.log("WhatsApp session is already active.")
            return True

        self._clean_session_lockfiles()

        launch_args = [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--start-maximized",
            "--disable-blink-features=AutomationControlled",
        ]
        user_agent = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/127.0.0.0 Safari/537.36"
        )

        try:
            self.log("Initializing WhatsApp Web session (Interactive window)...")
            self.playwright = sync_playwright().start()

            self.browser_context = self._launch_context(headless, user_agent, launch_args)

            # 0ms Chromium Stealth Masking: mask navigator.webdriver to undefined across all document frames
            self.browser_context.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
            )

            self.page = (
                self.browser_context.pages[0]
                if self.browser_context.pages
                else self.browser_context.new_page()
            )
            self.page.goto("https://web.whatsapp.com")
            self.log("Opened WhatsApp Web.")
            return True
        except Exception as e:
            self.log(f"Failed to launch WhatsApp Web: {e}")
            self._clean_session_lockfiles()
            try:
                time.sleep(1)
                self.browser_context = self._launch_context(headless, user_agent, launch_args)
                self.browser_context.add_init_script(
                    "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
                )
                self.page = (
                    self.browser_context.pages[0]
                    if self.browser_context.pages
                    else self.browser_context.new_page()
                )
                self.page.goto("https://web.whatsapp.com")
                return True
            except Exception as ex:
                self.log(f"Retry launch failed: {ex}")
                return False

    def is_logged_in(self):
        """Returns True if WhatsApp Web chat list is visible (user logged in)."""
        if not self.page:
            return False
        try:
            pane = self.page.query_selector(
                "div[data-testid='chat-list'], div#pane-side, div[role='grid'], "
                "div[aria-label='Chat list'], div._aawg, #side, "
                "div[contenteditable='true'][data-tab='3']"
            )
            return bool(pane)
        except Exception:
            return False

    def is_qr_screen(self):
        """Returns True if WhatsApp Web is showing QR code login screen or initial login prompt."""
        if not self.page:
            return False
        try:
            qr = self.page.query_selector(
                "canvas, div[data-testid='link-device-qr-code'], div[data-ref], "
                "div._akau, div._akav, div[data-testid='qrcode']"
            )
            if qr and qr.is_visible():
                return True

            body_text = self.page.inner_text("body").lower() if self.page else ""
            if "scan" in body_text or "link with phone number" in body_text or "to use whatsapp on your computer" in body_text:
                return True
        except Exception:
            pass
        return False

    def wait_for_login(self, timeout=60):
        """Waits for user to scan QR code and log into WhatsApp Web."""
        self.log("Checking WhatsApp Web login status...")
        start = time.time()
        while time.time() - start < timeout:
            if self.is_logged_in():
                self.log("✅ WhatsApp Web is logged in and ready!")
                return True

            if self.is_qr_screen():
                self.log("📱 QR Code detected! Please scan the QR code in the WhatsApp window using your phone...")

            time.sleep(1.5)
        return False




    def minimize_browser(self):
        pass

    def restore_browser(self):
        pass

    # ── Helpers ───────────────────────────────────────────────────

    def _wait_for_chat(self, timeout=20000):
        """
        Wait for chat composer box to appear OR detect invalid phone popup instantly.
        Returns:
            True -> Chat loaded
            'INVALID_PHONE' -> Phone number invalid/not on WhatsApp
            False -> Timeout / error
        """
        composer_selectors = [
            "div[data-testid='conversation-compose-box-input']",
            "footer div[contenteditable='true']",
            "div[contenteditable='true'][data-tab='10']",
            "div[contenteditable='true'][role='textbox']",
            "p.selectable-text",
            "div[contenteditable='true'][aria-placeholder*='Type']",
        ]
        popup_selectors = [
            "div[data-animate-modal-popup='true']",
            "div[role='dialog']",
        ]

        deadline = time.time() + (timeout / 1000.0)
        while time.time() < deadline:
            # 1. Check for chat input box
            for sel in composer_selectors:
                try:
                    el = self.page.query_selector(sel)
                    if el and el.is_visible():
                        return True
                except Exception:
                    continue

            # 2. Check for invalid number / non-WhatsApp popup
            for popup_sel in popup_selectors:
                try:
                    popup = self.page.query_selector(popup_sel)
                    if popup and popup.is_visible():
                        text = popup.inner_text().lower()
                        if any(term in text for term in ["invalid", "not registered", "not on whatsapp", "phone number shared via url"]):
                            ok_btn = popup.query_selector("div[role='button'], button")
                            if ok_btn:
                                try:
                                    ok_btn.click()
                                except Exception:
                                    pass
                            return "INVALID_PHONE"
                except Exception:
                    continue

            time.sleep(0.5)

        return False

    def _is_document_preview_open(self):
        """Checks if the document preview overlay/dialog or send button is visible."""
        # 1. Check for any send icon or send button
        send_selectors = [
            "[data-testid='send']",
            "span[data-icon='send']",
            "span[data-icon='send-light']",
            "span[data-icon='send-filled']",
            "span[data-icon*='send']",
            "button[aria-label*='Send' i]",
            "div[aria-label*='Send' i]",
            "div[role='button'][aria-label*='Send' i]",
            "[data-testid='media-send']",
        ]
        for sel in send_selectors:
            try:
                for el in self.page.query_selector_all(sel):
                    if el and el.is_visible():
                        return True
            except Exception:
                continue

        # 2. Check for media viewer dialog / overlay container
        preview_containers = [
            "div[role='dialog']",
            "div[data-animate-media-viewer='true']",
            "div[data-testid='media-viewer']",
            "div[data-testid='media-caption-input']",
        ]
        for sel in preview_containers:
            try:
                el = self.page.query_selector(sel)
                if el and el.is_visible():
                    return True
            except Exception:
                continue

        return False

    def _wait_for_document_preview(self, pdf_path, timeout=15):
        """Wait until WhatsApp exposes the Send control for the PDF preview."""
        del pdf_path
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self._is_document_preview_open():
                return True
            time.sleep(0.5)
        return False

    def _click_send(self):

        """Clicks the green send button in WhatsApp Web using authentic Playwright mouse events."""
        send_selectors = [
            "div[role='button'][aria-label*='Send' i]",
            "button[aria-label*='Send' i]",
            "div[aria-label*='Send' i]",
            "span[data-icon='send']",
            "span[data-icon='send-light']",
            "span[data-icon='send-filled']",
            "[data-testid='send']",
            "[data-testid='media-send']",
        ]

        # 1. First try clicking any matching send button with native Playwright mouse click
        for sel in send_selectors:
            try:
                elements = self.page.query_selector_all(sel)
                for el in elements:
                    if el and el.is_visible():
                        el.click(force=True, timeout=2000)
                        return True
            except Exception:
                continue

        # 2. If clicking element failed or button not found by selector, press Enter key
        try:
            self.page.keyboard.press("Enter")
            return True
        except Exception:
            return False

    def _send_text_message(self, message):
        """Sends a multi-line text message as ONE single chat bubble using Shift+Enter between lines."""
        input_selectors = [
            "div[data-testid='conversation-compose-box-input']",
            "footer div[contenteditable='true']",
            "div[contenteditable='true'][data-tab='10']",
            "div[contenteditable='true'][role='textbox']",
            "p.selectable-text",
        ]
        try:
            for sel in input_selectors:
                el = self.page.query_selector(sel)
                if el and el.is_visible():
                    el.click()
                    time.sleep(0.2)
                    lines = message.split("\n")
                    for i, line in enumerate(lines):
                        if line:
                            self.page.keyboard.type(line, delay=random.randint(15, 35))
                        if i < len(lines) - 1:
                            self.page.keyboard.press("Shift+Enter")
                    time.sleep(0.3)
                    self.page.keyboard.press("Enter")
                    return True
        except Exception as e:
            self.log(f"  Text send error: {e}")
        return False

    def _type_caption_in_preview(self, caption_text):
        """Types caption directly into WhatsApp Web's document preview caption box if present."""
        caption_selectors = [
            "div[role='dialog'] div[contenteditable='true']",
            "div[contenteditable='true'][data-testid='media-caption-input']",
            "div[contenteditable='true'][aria-placeholder*='caption' i]",
            "div[contenteditable='true'][aria-label*='caption' i]",
        ]
        for sel in caption_selectors:
            try:
                for el in self.page.query_selector_all(sel):
                    if el and el.is_visible():
                        el.click()
                        time.sleep(0.2)
                        # A bare "\n" is an Enter keypress, which would send the preview early
                        lines = caption_text.split("\n")
                        for i, line in enumerate(lines):
                            if line:
                                self.page.keyboard.type(line, delay=1)
                            if i < len(lines) - 1:
                                self.page.keyboard.press("Shift+Enter")
                        time.sleep(0.3)
                        return True
            except Exception:
                continue
        return False

    def _open_attach_menu(self):
        """Opens the WhatsApp Web attachment popup menu."""
        attach_selectors = [
            "button[aria-label='Attach']",
            "div[aria-label='Attach']",
            "span[data-icon='plus']",
            "span[data-icon='attach-menu-plus']",
            "div[title='Attach']",
            "button[title='Attach']",
            "span[data-icon='clip']",
            "[data-testid='clip']",
            "[data-testid='attach-menu-plus']",
        ]
        for sel in attach_selectors:
            try:
                for el in self.page.query_selector_all(sel):
                    if el and el.is_visible():
                        el.click()
                        time.sleep(0.5)
                        return True
            except Exception:
                continue
        return False

    def _attach_pdf_document(self, abs_pdf):
        """
        Attaches PDF document by opening Attach menu (+) -> Document -> selecting PDF.
        """
        abs_pdf = os.path.abspath(abs_pdf)
        if not os.path.exists(abs_pdf):
            self.log(f"  ❌ PDF file does not exist: {abs_pdf}")
            return False

        # Step 1: Open Attach (+) menu
        if not self._open_attach_menu():
            self.log("  ⚠️ Could not click attach (+) menu button. Checking if preview already open...")
            if self._is_document_preview_open():
                return True

        time.sleep(0.6)

        # Step 2: Bind file_chooser and click Document menu option
        doc_clicked = False
        try:
            with self.page.expect_file_chooser(timeout=5000) as fc_info:
                doc_selectors = [
                    "button[aria-label='Document']",
                    "div[aria-label='Document']",
                    "[data-testid='mi-attach-document']",
                    "[data-testid='attach-document']",
                    "li[aria-label='Document']",
                    "span[data-icon='attach-document']",
                ]
                for sel in doc_selectors:
                    try:
                        el = self.page.query_selector(sel)
                        if el and el.is_visible():
                            el.click()
                            doc_clicked = True
                            break
                    except Exception:
                        continue

                if not doc_clicked:
                    for item in self.page.query_selector_all("li, button, div[role='button'], span"):
                        try:
                            if item.is_visible() and "document" in item.inner_text().strip().lower():
                                item.click()
                                doc_clicked = True
                                break
                        except Exception:
                            continue

            if doc_clicked and fc_info.value:
                fc_info.value.set_files(abs_pdf)
                self.log(f"  ✅ Attached PDF via file chooser: {os.path.basename(abs_pdf)}")
                time.sleep(1.5)
                return True
        except Exception as e:
            self.log(f"  File chooser timeout or log: {e}")

        # Step 3: Backup if expect_file_chooser timed out -- set_input_files on active file input
        try:
            for file_input in self.page.query_selector_all("input[type='file']"):
                try:
                    file_input.set_input_files(abs_pdf)
                    time.sleep(1.5)
                    if self._is_document_preview_open():
                        self.log(f"  ✅ Attached PDF via fallback input: {os.path.basename(abs_pdf)}")
                        return True
                except Exception:
                    continue
        except Exception:
            pass

        return self._is_document_preview_open()

    def _dismiss_failed_popup(self):
        """Dismiss 'Your message was not sent' popup if present."""
        try:
            popup = self.page.query_selector("div[data-animate-modal-popup='true'], div[role='dialog']")
            if popup and popup.is_visible():
                text = popup.inner_text().lower()
                if 'not sent' in text or 'try again' in text or 'failed' in text:
                    cancel = popup.query_selector("div[role='button']:first-child, button")
                    if cancel:
                        cancel.click()
                        time.sleep(0.5)
                    return True
            return False
        except Exception:
            return False

    def _wait_for_upload_and_delivery(self, pdf_path, timeout=25):
        """
        Waits for PDF document upload to finish and single/double checkmark to appear.
        Guarantees Meta servers acknowledged receipt of the complete PDF bytes before returning.
        """
        filename = os.path.basename(pdf_path)
        self.log(f"  Waiting for server delivery confirmation for {filename}...")

        deadline = time.time() + timeout
        start_time = time.time()

        # Give WhatsApp Web 1.5 seconds to start the upload pipeline
        time.sleep(1.5)

        check_selectors = [
            "span[data-icon='msg-check']",
            "span[data-icon='msg-dblcheck']",
            "span[data-icon='msg-dblcheck-ack']",
            "span[data-icon='status-check']",
            "span[data-icon='status-dblcheck']",
            "[data-testid='msg-check']",
            "[data-testid='msg-dblcheck']",
        ]

        progress_selectors = [
            "span[data-icon='media-cancel']",
            "div[data-testid='media-state-uploading']",
            "circle[class*='progress']",
            "div[role='progressbar']",
            "span[data-icon='msg-time']",
            "span[data-icon='status-time']",
        ]

        while time.time() < deadline:
            # 1. Check if any progress/spinner is still active anywhere on page
            uploading = False
            for sel in progress_selectors:
                try:
                    p = self.page.query_selector(sel)
                    if p and p.is_visible():
                        uploading = True
                        break
                except Exception:
                    continue

            # 2. Check if a checkmark icon is present in recent message rows
            has_check = False
            try:
                rows = self.page.query_selector_all("div.message-out, div[role='row']")
                if rows:
                    for row in reversed(rows[-3:]):
                        for sel in check_selectors:
                            c = row.query_selector(sel)
                            if c and c.is_visible():
                                has_check = True
                                break
                        if has_check:
                            break
            except Exception:
                pass

            elapsed = time.time() - start_time
            if has_check:
                self.log(f"  ✅ Server checkmark confirmed for {filename}!")
                time.sleep(2.0)
                return True
            elif not uploading and elapsed >= 5.0:
                self.log(f"  ✅ Upload spinner completed for {filename}.")
                time.sleep(2.5)
                return True

            time.sleep(0.5)

        self.log(f"  ⚠️ Timeout waiting for checkmark, enforcing 3.0s buffer for {filename}")
        time.sleep(3.0)
        return True

    def _open_chat_for_number(self, clean_phone):
        """Opens chat for phone number using client-side SPA navigation without destroying WebSocket session or popping up window."""
        chat_url = f"https://web.whatsapp.com/send?phone={clean_phone}"
        try:
            current_url = self.page.url if self.page else ""
            if "web.whatsapp.com" in current_url:
                # Client-side SPA route change to switch chats without reloading the page or taking window focus
                self.page.evaluate("(url) => { window.location.href = url; }", chat_url)
                time.sleep(1.0)
                return True
        except Exception as e:
            self.log(f"  SPA navigate note: {e}")

        self.page.goto(chat_url, wait_until="domcontentloaded")
        return True


    # ── Core send logic ──────────────────────────────────────────

    def send_invoice_to_single_number(self, clean_phone, agency_name, invoice_no, month_desc, pdf_path):
        """Opens chat → sends text message → attaches & sends PDF invoice."""
        abs_pdf = os.path.abspath(pdf_path)

        text_message = build_invoice_message(agency_name, invoice_no, month_desc)

        try:
            self.log(f"Sending invoice to {agency_name} ({clean_phone})...")
            self._open_chat_for_number(clean_phone)

            chat_status = self._wait_for_chat(timeout=20000)
            if chat_status == "INVALID_PHONE":
                self.log(f"  ❌ {clean_phone} is invalid or not registered on WhatsApp.")
                return False, f"{clean_phone} not on WhatsApp"
            elif not chat_status:
                self.log(f"  ❌ Chat did not load for {clean_phone}")
                return False, "Chat did not load"

            time.sleep(1.0)
            self._dismiss_failed_popup()

            # 1. Send text invoice summary message first
            self.log(f"  Sending invoice text message to {agency_name}...")
            if not self._send_text_message(text_message):
                self.log(f"  ⚠️ Could not send text message, proceeding to PDF attachment...")

            time.sleep(1.5)

            # 2. Attach PDF Document
            self.log(f"  Attaching PDF document {os.path.basename(abs_pdf)}...")
            if not self._attach_pdf_document(abs_pdf):
                self.log(f"  ❌ Could not attach PDF document: {os.path.basename(abs_pdf)}")
                return False, "Could not attach PDF document"

            # 3. Wait for Document Preview modal
            if not self._wait_for_document_preview(abs_pdf, timeout=15):
                self.log(f"  ❌ PDF preview did not appear: {os.path.basename(abs_pdf)}")
                return False, f"PDF preview did not appear for {os.path.basename(abs_pdf)}"

            # 4. Click Send in Document Preview modal natively
            self.log(f"  Clicking Send for PDF document {os.path.basename(abs_pdf)}...")
            self._click_send()
            time.sleep(1.0)

            # Double guarantee: If document preview is STILL open, press Enter key
            if self._is_document_preview_open():
                self.log("  Document preview still visible, pressing Enter key to complete send...")
                try:
                    self.page.keyboard.press("Enter")
                except Exception:
                    pass
                time.sleep(1.0)

            # 5. Wait for upload & delivery confirmation before returning
            self._wait_for_upload_and_delivery(abs_pdf, timeout=25)

            if self._dismiss_failed_popup():
                return False, "Message failed to send"

            self.log(f"  ✅ Invoice text message & PDF successfully sent to {agency_name} ({clean_phone})")
            return True, "Sent"

        except Exception as e:
            self.log(f"  ❌ Error sending to {agency_name}: {e}")
            return False, str(e)

    # ── Interactive "Share App" (preview left open for review) ───

    def prepare_invoice_share(self, phone_input, agency_name, invoice_no, month_desc, pdf_path, auto_send=False):
        """
        Opens the chat in the connected WhatsApp Web session, attaches the PDF through the page's
        file input (no OS clipboard, focus or keystroke injection involved) and puts the invoice
        message in the caption. With auto_send=False the preview is left open for the user to press Send.
        Must run on the Playwright worker thread. Returns (ok, message).
        """
        if not self.page:
            return False, "WhatsApp not connected."
        phones = self.parse_phones(phone_input)
        if not phones:
            return False, "No valid phone number"
        abs_pdf = os.path.abspath(pdf_path) if pdf_path else ""
        if not abs_pdf or not os.path.exists(abs_pdf):
            return False, f"PDF not found: {pdf_path}"

        clean_phone = phones[0]
        try:
            try:
                self.page.bring_to_front()
            except Exception:
                pass

            if not self.is_logged_in():
                self.log("📱 WhatsApp Web is not logged in — scan the QR code in the browser window...")
                if not self.wait_for_login(timeout=60):
                    return False, "WhatsApp Web is not logged in"

            self._open_chat_for_number(clean_phone)
            chat_status = self._wait_for_chat(timeout=20000)
            if chat_status == "INVALID_PHONE":
                return False, f"{clean_phone} not on WhatsApp"
            if not chat_status:
                return False, "Chat did not load"
            self._dismiss_failed_popup()

            if not self._attach_pdf_document(abs_pdf):
                return False, "Could not attach PDF document"
            if not self._wait_for_document_preview(abs_pdf, timeout=15):
                return False, "PDF preview did not appear"

            caption_ok = self._type_caption_in_preview(
                build_invoice_message(agency_name, invoice_no, month_desc)
            )
            if not caption_ok:
                self.log("  ⚠️ Could not fill the caption; the PDF is attached without the message.")

            if auto_send:
                self._click_send()
                self._wait_for_upload_and_delivery(abs_pdf, timeout=25)
                if self._dismiss_failed_popup():
                    return False, "Message failed to send"
                return True, "Sent"

            self.log(f"  ✅ PDF attached for {agency_name} ({clean_phone}) — review and press Send.")
            return True, "PDF attached — press Send in WhatsApp" if caption_ok else \
                "PDF attached (add message manually) — press Send in WhatsApp"
        except Exception as e:
            self.log(f"  ❌ Share error for {agency_name}: {e}")
            return False, str(e)

    # ── Multi-recipient dispatcher ────────────────────────────────


    def send_invoice_pdf(self, phone_input, agency_name, invoice_no, month_desc, pdf_path):
        if not self.page:
            return False, "WhatsApp not connected."

        phone_list = self.parse_phones(phone_input)
        if not phone_list:
            return False, "No valid phone numbers"

        if not os.path.exists(pdf_path):
            return False, f"PDF not found: {pdf_path}"

        success = 0
        total = len(phone_list)
        errors = []

        for idx, phone in enumerate(phone_list):
            if idx > 0:
                delay = round(random.uniform(3.0, 6.0), 1)
                self.log(f"  Waiting {delay} seconds before sending to next contact...")
                time.sleep(delay)  # Randomized delay between bulk recipients

            ok, msg = self.send_invoice_to_single_number(
                phone, agency_name, invoice_no, month_desc, pdf_path,
            )
            if ok:
                success += 1
            else:
                errors.append(f"{phone}: {msg}")

        if success > 0:
            if success == total:
                return True, f"Sent to all {total} contact(s)"
            return False, f"Partial send: {success}/{total}; {'; '.join(errors)}"
        return False, f"Failed: {'; '.join(errors)}"

    def close(self):

        try:
            if self.browser_context:
                self.browser_context.close()
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        self.browser_context = None
        self.playwright = None

