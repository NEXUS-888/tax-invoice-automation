# Ananya Enterprises - Tax Invoice Automation & WhatsApp Dispatcher 🚀

[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-blue.svg)](#cross-platform-support)
[![Python](https://img.shields.io/badge/Python-3.10%2B-brightgreen.svg)](https://www.python.org/)
[![GUI](https://img.shields.io/badge/UI-PyQt6-darkblue.svg)](https://riverbankcomputing.com/software/pyqt/)
[![Automation](https://img.shields.io/badge/Engine-Playwright-orange.svg)](https://playwright.dev/)
[![Tests](https://img.shields.io/badge/Tests-39%20Passing-success.svg)](#running-tests)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

An enterprise-grade desktop automation solution designed for **Ananya Enterprises** to streamline monthly tax invoice generation, real-time load/rate adjustments, contact management, and direct WhatsApp invoice delivery.

---

## ✨ Features

- **📑 Automated Tax Invoice Generation:**
  - Instantly parses agency data from Excel workbooks (`.xlsm` / `.xlsx`).
  - Computes GST, totals, and Indian currency words automatically.
  - Generates pixel-perfect, printable ReportLab PDF tax invoices with standardized agency metadata.
  - Strictly embeds required payment terms (*"Note : please complete the payment before 10th of this month"*).

- **⚡ Real-Time Load & Rate Customization:**
  - Live editable tables in Tab 1 ("Agency & Load Manager") and Tab 3 ("Generate & Auto-Send").
  - Double-click any vehicle rate or load count to update totals instantaneously.
  - Editable invoice numbers with custom sequencing.

- **✏️ Agency Info & Contact Management:**
  - Dedicated **"✏️ Edit Agency Info"** dialog to modify agency names, vendor codes, addresses, PAN, GSTIN, and WhatsApp contact numbers.
  - Add new agencies and add custom vehicles on-the-fly.

- **💾 Persistent State Retention:**
  - Built-in `StateManager` saves all edits (loads, rates, agency metadata, custom invoice numbers) to `data/app_state.json`.
  - Your changes **never revert** to default workbook data on app restart or re-import.
  - Reset to workbook defaults anytime with the "🔄 Reset Defaults" button.

- **📱 Dual WhatsApp Sharing & Automation:**
  1. **Direct Share (`📤 Share App` / `🌐 Share Web`):**
     - Instant 1-click invoice sharing without browser overhead.
     - Automatically generates on-demand PDF if not already generated.
     - Automatically copies the PDF invoice to system clipboard (`CF_HDROP`) — simply press `Ctrl+V` to attach the file in WhatsApp.
     - Opens WhatsApp Desktop App or WhatsApp Web with pre-formatted invoice message and payment reminder.
     - Highlights the generated PDF in your system file explorer for quick drag-and-drop.
  2. **Automated WhatsApp Web Dispatcher (`📱 Auto-Send`):**
     - Interactive session login with QR code persistence in `data/wa_session`.
     - Multi-browser auto-fallback: Works with Playwright Chromium, Microsoft Edge, and Google Chrome out-of-the-box.
     - Bulk hands-free dispatch with real-time progress indicators, audit logs, and retry handling.

- **🧹 Clean File Management & Deduplication:**
  - Guaranteed single PDF per agency — eliminates duplicate stale copies.
  - Integrated output directory viewer and one-click PDF launcher.

---

## 🚀 Quick Start for Windows (Zero Setup Required!)

### 🌟 Option 1: 1-Click Desktop Installer (Recommended)
No extraction, no terminal, no technical setup:
1. Go to [**GitHub Releases**](https://github.com/NEXUS-888/tax-invoice-automation/releases).
2. Download **`Ananya_Invoice_Automation_Setup.exe`**.
3. Double-click the installer and click **"Install Now"**.
4. The setup wizard automatically:
   - Installs the app to your user programs folder (no admin permissions required).
   - Creates a **Desktop Shortcut** ("Ananya Invoice Automation").
   - Creates a **Start Menu Shortcut**.
   - Launches the app immediately!

### 📦 Option 2: Standalone Portable ZIP
If you prefer not to install anything and want a portable folder:
1. Download **`Ananya_Invoice_Automation_Windows.zip`** from [**Releases**](https://github.com/NEXUS-888/tax-invoice-automation/releases).
2. Extract the folder anywhere on your PC.
3. Double-click **`AnanyaInvoiceAutomation.exe`**.
4. The template `Ananya Bill.xlsm` is already included in the folder.

---

## 💻 Cross-Platform Setup from Source (Windows, macOS, Linux)

You can run and develop the application from source on any modern operating system:

### 1. Prerequisites
- **Python 3.10+** installed:
  - Windows: [python.org](https://www.python.org/) (ensure "Add Python to PATH" is checked)
  - macOS: `brew install python`
  - Linux: `sudo apt update && sudo apt install python3 python3-venv python3-pip`

### 2. Clone the Repository
```bash
git clone https://github.com/NEXUS-888/tax-invoice-automation.git
cd tax-invoice-automation
```

### 3. Create & Activate Virtual Environment
- **On Windows (PowerShell):**
  ```powershell
  python -m venv venv
  .\venv\Scripts\activate
  ```
- **On macOS / Linux:**
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  ```

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

*(Optional for automated Playwright WhatsApp bot):*
```bash
playwright install chromium
```
> **Note:** If you don't run `playwright install`, the app will automatically fall back to using your pre-installed Microsoft Edge or Google Chrome browser.

### 5. Launch the Application
```bash
python main.py
```

---

## 📦 Building Standalone Executable (.exe)

To bundle the application into a standalone Windows distribution:

```powershell
python build_exe.py
```

This creates:
- `dist/AnanyaInvoiceAutomation/`: Complete standalone folder containing `AnanyaInvoiceAutomation.exe`, bundled master template `Ananya Bill.xlsm`, `data/`, `output/`, and setup guide.
- `dist/Ananya_Invoice_Automation_Windows.zip`: Ready-to-distribute compressed archive.

---

## 📁 Repository Structure

```
tax-invoice-automation/
├── src/
│   ├── app_gui.py             # Main modern PyQt6 GUI (3-tab workflow, state, modals)
│   ├── excel_manager.py       # Master workbook loader, GST calculator, number-to-words
│   ├── pdf_generator.py       # Pixel-accurate ReportLab tax invoice renderer
│   ├── whatsapp_dispatcher.py # 1-click WhatsApp Desktop/Web share & clipboard copier
│   ├── whatsapp_automator.py  # Playwright automated WhatsApp Web session & delivery
│   ├── contacts_manager.py    # Agency telephone numbers & directory management
│   └── state_manager.py       # Persistent JSON state store for edits & overrides
├── tests/                     # Comprehensive test suite (39 tests)
│   ├── test_excel_manager.py
│   ├── test_pdf_generator.py
│   ├── test_gui_integration.py
│   ├── test_contacts_manager.py
│   ├── test_state_manager.py
│   ├── test_whatsapp_dispatcher.py
│   └── test_whatsapp_automator.py
├── data/                      # App data (contacts, state, WhatsApp persistent session)
├── output/                    # Generated PDF invoices and updated Excel sheets
├── Ananya Bill.xlsm           # Master Excel workbook template
├── build_exe.py               # Automated PyInstaller packaging script
├── main.py                    # Application entry point
├── requirements.txt           # Python package requirements
└── README.md                  # Project documentation
```

---

## 🧪 Running Tests

To verify system integrity, run the full automated test suite:

```bash
python -m unittest discover -s tests -v
```

All 39 unit and integration tests validate:
- Accurate GST calculations and Indian Rupee number-to-words conversion.
- PDF generation fidelity and payment deadline notes.
- On-demand PDF generation when sharing via WhatsApp.
- Rate, load, and invoice number persistence across sessions.
- Edit Agency Info dialog validation (handling integer/string vendor codes and blank fields).
- Cross-platform file manager highlighting and URL dispatching.
- WhatsApp Web session locking, resilience, and retry logic.

---

## 📖 User Workflow Guide

1. **Tab 1: Agency & Load Manager**
   - Review imported agencies and vehicles.
   - Double-click any rate or load to make adjustments.
   - Click **✏️ Edit Agency Info** to update GST, address, or WhatsApp numbers.
   - All changes save automatically in the background.

2. **Tab 2: Master Contact Directory**
   - Maintain phone numbers and notification preferences for each agency.
   - Syncs dynamically with workbook entries.

3. **Tab 3: Generate & Auto-Send WhatsApp Invoices**
   - Click **1. Generate All PDFs & Excel** to generate the full batch.
   - Use **📤 Share App** next to any agency to immediately open WhatsApp with the PDF copied to clipboard.
   - Or click **📱 Connect WhatsApp Account** followed by **2. Automated Send ALL Invoices** for bulk delivery.

---

## 🛡️ License

This project is licensed under the [MIT License](LICENSE).
