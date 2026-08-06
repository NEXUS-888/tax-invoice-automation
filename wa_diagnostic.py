"""Quick diagnostic to see what WhatsApp Web looks like to Playwright."""
import os, sys, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from playwright.sync_api import sync_playwright

SESSION_DIR = os.path.join(os.path.dirname(__file__), "data", "wa_session")
SCREENSHOT_PATH = os.path.join(os.path.dirname(__file__), "wa_debug.png")
TEST_PHONE = "919141844840"  # First agency number

pw = sync_playwright().start()
ctx = pw.chromium.launch_persistent_context(
    user_data_dir=SESSION_DIR,
    headless=False,
    viewport={"width": 1280, "height": 900},
    args=["--no-sandbox"],
)
page = ctx.pages[0] if ctx.pages else ctx.new_page()

print("1. Navigating to WhatsApp Web...")
page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")
print("   Waiting 10 seconds for WhatsApp to load...")
time.sleep(10)

# Take screenshot of main page
page.screenshot(path=SCREENSHOT_PATH.replace(".png", "_main.png"))
print(f"   Screenshot saved: {SCREENSHOT_PATH.replace('.png', '_main.png')}")

# Check what's on the page
print("\n2. Checking page elements...")
selectors_to_check = [
    ("Chat input (data-tab=10)", "div[contenteditable='true'][data-tab='10']"),
    ("Chat input (footer)", "footer div[contenteditable='true']"),
    ("Chat input (textbox)", "div[contenteditable='true'][role='textbox']"),
    ("Side pane", "div#pane-side"),
    ("Chat list", "div[data-testid='chat-list']"),
    ("QR code", "canvas[aria-label*='QR'], div[data-testid='qrcode']"),
    ("Intro/landing", "div._al_c, div[data-testid='intro-md-beta-message']"),
    ("Attach button (plus)", "span[data-icon='plus']"),
    ("Attach button (clip)", "span[data-icon='clip']"),
    ("Search", "div[data-testid='chat-list-search']"),
]

for label, sel in selectors_to_check:
    el = page.query_selector(sel)
    print(f"   {'✅' if el else '❌'} {label}: {'FOUND' if el else 'NOT FOUND'}")

print(f"\n3. Navigating to chat URL for test phone {TEST_PHONE}...")
page.goto(f"https://web.whatsapp.com/send?phone={TEST_PHONE}", wait_until="domcontentloaded")
print("   Waiting 15 seconds for chat to load...")
time.sleep(15)

# Take screenshot of chat page
page.screenshot(path=SCREENSHOT_PATH.replace(".png", "_chat.png"))
print(f"   Screenshot saved: {SCREENSHOT_PATH.replace('.png', '_chat.png')}")

# Check chat elements
print("\n4. Checking chat elements...")
chat_selectors = [
    ("Chat input (data-tab=10)", "div[contenteditable='true'][data-tab='10']"),
    ("Chat input (footer)", "footer div[contenteditable='true']"),
    ("Chat input (textbox)", "div[contenteditable='true'][role='textbox']"),
    ("Any contenteditable", "div[contenteditable='true']"),
    ("Popup/modal", "div[data-animate-modal-popup='true']"),
    ("Invalid phone popup", "div[role='dialog']"),
]

for label, sel in chat_selectors:
    el = page.query_selector(sel)
    if el:
        try:
            text = el.inner_text()[:100]
            print(f"   ✅ {label}: FOUND (text: {repr(text)})")
        except:
            print(f"   ✅ {label}: FOUND")
    else:
        print(f"   ❌ {label}: NOT FOUND")

# Dump all contenteditable elements
print("\n5. All contenteditable divs on page:")
for el in page.query_selector_all("div[contenteditable='true']"):
    try:
        attrs = page.evaluate("(el) => { let a = {}; for (let attr of el.attributes) a[attr.name] = attr.value; return a; }", el)
        print(f"   Found: {attrs}")
    except:
        print(f"   Found: (could not read attributes)")

print("\n6. Page title:", page.title())
print("   Page URL:", page.url)

input("\nPress Enter to close browser and exit...")
ctx.close()
pw.stop()
