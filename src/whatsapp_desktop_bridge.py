"""
Drives the WhatsApp Desktop app for Windows from the inside.

Since late 2025 WhatsApp for Windows (WhatsApp.Root.exe) is web.whatsapp.com hosted in a WebView2
window. Outside automation (clipboard + simulated keys/clicks) is blind to it, and its whatsapp://
links open a "Send to" picker. Instead we use WebView2's documented remote-debugging switch and
attach Playwright to the page over CDP — the same mechanism Playwright documents for WebView2 apps.

The switch is passed through the WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS user environment variable
(no admin rights needed). It is set only while WhatsApp restarts and removed straight afterwards,
so no other WebView2 app picks it up. WhatsApp keeps the port until it is next closed.

Nothing here ever presses Send.
"""
import ctypes
import json
import os
import subprocess
import sys
import time
import urllib.parse
import urllib.request

DEBUG_PORT = 47823
ENV_VAR = "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"
WHATSAPP_EXE = "WhatsApp.Root.exe"
# Store app id of WhatsApp for Windows (publisher hash is fixed for the Store package).
WHATSAPP_AUMID = "5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App"


class WhatsAppDesktopBridge:
    def __init__(self, port=DEBUG_PORT, log=None):
        self.port = port
        self.log = log or (lambda msg: None)

    # ── Debug port ───────────────────────────────────────────────

    def _get_json(self, path, timeout=1.5):
        with urllib.request.urlopen(f"http://127.0.0.1:{self.port}{path}", timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))

    def is_available(self):
        """True if a WhatsApp Desktop page is reachable on the debug port."""
        try:
            return any("web.whatsapp.com" in t.get("url", "") and t.get("type") == "page"
                       for t in self._get_json("/json/list"))
        except Exception:
            return False

    def is_whatsapp_installed(self):
        if sys.platform != "win32":
            return False
        local = os.environ.get("LOCALAPPDATA", "")
        family = WHATSAPP_AUMID.split("!")[0]
        return os.path.isdir(os.path.join(local, "Packages", family))

    @staticmethod
    def _set_user_env(value):
        """Sets (or with None, removes) the WebView2 switch in the user environment and notifies Windows."""
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_SET_VALUE) as key:
            if value is None:
                try:
                    winreg.DeleteValue(key, ENV_VAR)
                except FileNotFoundError:
                    pass
            else:
                winreg.SetValueEx(key, ENV_VAR, 0, winreg.REG_SZ, value)
        result = ctypes.c_size_t()
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment",  # WM_SETTINGCHANGE
                                                 0x0002, 3000, ctypes.byref(result))  # SMTO_ABORTIFHUNG

    def enable(self, timeout=40.0):
        """
        Restarts WhatsApp with the debug port switched on. Returns (ok, message).
        Restarting only closes the window; chats and login are kept by WhatsApp.
        """
        if self.is_available():
            return True, "already connected"
        if not self.is_whatsapp_installed():
            return False, "WhatsApp Desktop (Microsoft Store version) is not installed."

        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        self.log("Connecting to WhatsApp Desktop (restarting it once)...")
        try:
            self._set_user_env(f"--remote-debugging-port={self.port}")
            subprocess.run(["taskkill", "/IM", WHATSAPP_EXE, "/F"], capture_output=True, creationflags=flags)
            for _ in range(20):  # wait for the old process to exit so the switch is picked up
                running = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {WHATSAPP_EXE}", "/NH"],
                                         capture_output=True, text=True, creationflags=flags).stdout
                if WHATSAPP_EXE.lower() not in running.lower():
                    break
                time.sleep(0.25)
            os.startfile(f"shell:AppsFolder\\{WHATSAPP_AUMID}")
            deadline = time.time() + timeout
            while time.time() < deadline:
                if self.is_available():
                    return True, "connected"
                time.sleep(0.5)
            return False, "WhatsApp restarted but did not accept the connection."
        except Exception as e:
            return False, f"Could not restart WhatsApp: {e}"
        finally:
            try:
                self._set_user_env(None)
            except Exception:
                pass

    # ── Attach ───────────────────────────────────────────────────

    @staticmethod
    def chat_url(current_url, phone):
        """web.whatsapp.com/send URL for `phone`, keeping the desktop wrapper's own query parameters."""
        query = urllib.parse.parse_qsl(urllib.parse.urlsplit(current_url).query)
        query = [(k, v) for k, v in query if k not in ("phone", "text", "app_absent")]
        return "https://web.whatsapp.com/send?" + urllib.parse.urlencode([("phone", phone)] + query)

    def bring_to_front(self):
        """Shows the WhatsApp window so the user can review the attached invoice."""
        try:
            from whatsapp_dispatcher import WhatsAppDispatcher
            d = WhatsAppDispatcher()
            wins = d.find_whatsapp_desktop_windows()
            if wins:
                ctypes.windll.user32.ShowWindow(wins[0][0], 9)  # SW_RESTORE
                d.force_foreground(wins[0][0])
            else:
                os.startfile(f"shell:AppsFolder\\{WHATSAPP_AUMID}")
        except Exception:
            pass

    OPEN_CHAT_TITLE = r"() => { const h = document.querySelector('#main header'); return h ? h.innerText.split('\n')[0] : null }"

    def _open_chat_title(self, page):
        try:
            return page.evaluate(self.OPEN_CHAT_TITLE)
        except Exception:
            return None

    def _wait_for_user_to_pick_chat(self, page, timeout):
        """
        Closes the open chat, shows WhatsApp's "New chat" contact list and waits until the user opens
        a chat (from that list, its search, or the chat list). Returns the chat title or None.
        """
        page.keyboard.press("Escape")  # closes the currently open chat, so any chat that opens is the choice
        time.sleep(0.8)
        baseline = self._open_chat_title(page)
        try:
            new_chat = page.query_selector("[role=button][aria-label='New chat'], button[aria-label='New chat']")
            if new_chat:
                new_chat.click()
        except Exception:
            pass  # the chat list and its search still work for picking
        self.log("Pick the contact in WhatsApp — the PDF and message will be added automatically.")
        deadline = time.time() + timeout
        while time.time() < deadline:
            title = self._open_chat_title(page)
            if title and title != baseline:
                return title
            time.sleep(0.5)
        return None

    def attach_invoice(self, phone, caption, pdf_path, auto_send=False, pick_timeout=300):
        """
        Gets the PDF into a chat inside WhatsApp Desktop with `caption` as its message:
          phone given -> opens that chat directly;
          phone None  -> shows WhatsApp's contact list and waits for the user to pick the chat.
        Attaches through WhatsApp's own file input and leaves the preview open for the user to press
        Send. Blocking — run on a background thread. Returns (ok, message).
        """
        from playwright.sync_api import sync_playwright
        from whatsapp_automator import WhatsAppAutomator

        abs_pdf = os.path.abspath(pdf_path)
        pw = sync_playwright().start()
        try:
            browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{self.port}", timeout=10000)
            pages = [p for ctx in browser.contexts for p in ctx.pages if "web.whatsapp.com" in p.url]
            if not pages:
                return False, "WhatsApp Desktop page not found."
            page = pages[0]

            wa = WhatsAppAutomator.__new__(WhatsAppAutomator)  # reuse its WhatsApp Web helpers only
            wa.page, wa.log_callback = page, self.log

            # WebView2 stops rendering while its window is hidden/minimized, which stalls the chat,
            # the attach menu and the preview — so show WhatsApp before driving it.
            self.bring_to_front()
            time.sleep(0.5)

            if not wa.is_logged_in() and not wa.wait_for_login(timeout=30):
                return False, "WhatsApp Desktop is not logged in."

            if phone:
                page.goto(self.chat_url(page.url, phone), wait_until="domcontentloaded", timeout=30000)
            elif not self._wait_for_user_to_pick_chat(page, pick_timeout):
                return False, "no chat was picked in WhatsApp"

            status = wa._wait_for_chat(timeout=30000)
            if status == "INVALID_PHONE":
                return False, f"{phone} is not on WhatsApp."
            if not status:
                return False, "The chat did not open in WhatsApp Desktop."
            wa._dismiss_failed_popup()

            if not wa._attach_pdf_document(abs_pdf):
                return False, "WhatsApp did not accept the PDF."
            if not wa._wait_for_document_preview(abs_pdf, timeout=15):
                return False, "The PDF preview did not appear."
            caption_ok = wa._type_caption_in_preview(caption)

            if auto_send:
                wa._click_send()
                wa._wait_for_upload_and_delivery(abs_pdf, timeout=25)
                return True, "Sent"

            if caption_ok:
                return True, "PDF attached with the invoice message — press Send in WhatsApp."
            return True, "PDF attached (type the message yourself) — press Send in WhatsApp."
        except Exception as e:
            return False, f"WhatsApp Desktop automation error: {e}"
        finally:
            try:
                pw.stop()  # drops the connection only; never closes WhatsApp
            except Exception:
                pass
