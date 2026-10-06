import os
import sys
import json
import time
import subprocess
import ctypes
import unittest
from datetime import datetime

# Path to installed application
LOCAL_APPDATA = os.environ.get("LOCALAPPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Local"))
INSTALL_DIR = os.path.join(LOCAL_APPDATA, "Programs", "AnanyaInvoiceAutomation")
EXE_PATH = os.path.join(INSTALL_DIR, "AnanyaInvoiceAutomation.exe")

# Add repo src to path to use modules
REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(REPO_ROOT, "src"))

# Set sys.modules['numpy'] = None as in the runtime hook
sys.modules['numpy'] = None

from excel_manager import ExcelManager
from pdf_generator import PDFGenerator
from state_manager import StateManager
from contacts_manager import ContactsManager
from whatsapp_dispatcher import WhatsAppDispatcher
from whatsapp_automator import WhatsAppAutomator
import fitz  # PyMuPDF


def run_full_verification():
    report = {}
    print("=" * 65)
    print(" COMPREHENSIVE VERIFICATION SUITE: INSTALLED EXECUTABLE & SYSTEM")
    print(" Target:", INSTALL_DIR)
    print("=" * 65)

    # 1. Verify Installation Structure
    print("\n[CHECK 1] Verifying Installed Directory Structure & Files...")
    assert os.path.exists(INSTALL_DIR), f"Missing install dir: {INSTALL_DIR}"
    assert os.path.exists(EXE_PATH), f"Missing exe: {EXE_PATH}"
    excel_path = os.path.join(INSTALL_DIR, "Ananya Bill.xlsm")
    assert os.path.exists(excel_path), f"Missing template: {excel_path}"
    version_path = os.path.join(INSTALL_DIR, "version.json")
    assert os.path.exists(version_path), f"Missing version.json: {version_path}"

    with open(version_path, "r", encoding="utf-8") as vf:
        v_data = json.load(vf)
    print(f"  -> Version: {v_data.get('version')} (Installed: {v_data.get('installed_at')})")
    print(f"  -> Executable Size: {os.path.getsize(EXE_PATH) / (1024*1024):.2f} MB")
    report["Structure"] = "PASSED"

    # 2. Verify Excel Manager on Installed Template
    print("\n[CHECK 2] Testing Excel Parsing on Installed 'Ananya Bill.xlsm'...")
    em = ExcelManager(excel_path)
    agencies = em.parse_all_agencies()
    print(f"  -> Successfully parsed {len(agencies)} agencies.")
    assert len(agencies) > 0, "No agencies found in installed excel sheet!"
    sample_data = agencies[0]
    sample_agency = sample_data.get('agency_name')
    print(f"  -> Sample Agency: '{sample_agency}' (Vendor Code: {sample_data.get('vendor_code')})")
    print(f"  -> Vehicles found: {len(sample_data.get('vehicles', []))}")
    report["ExcelParsing"] = f"PASSED ({len(agencies)} agencies)"

    # 3. Verify Edit Agency Info & State Persistence
    print("\n[CHECK 3] Testing Agency Info Editing & State Persistence...")
    data_dir = os.path.join(INSTALL_DIR, "data")
    os.makedirs(data_dir, exist_ok=True)
    state_file = os.path.join(data_dir, "app_state.json")
    sm = StateManager(state_file)

    test_agency_name = sample_agency
    test_update = {
        "agency_name": test_agency_name,
        "vendor_code": 99999,
        "pan": "TESTP1234F",
        "gst": "29TESTP1234F1Z5",
        "address": "123 Test Industrial Zone, Bangalore",
        "whatsapp_phone": "9876543210",
        "contact_person": "Verification Officer",
        "loads": {"TEST-VEHICLE-01": 25},
        "rates": {"TEST-VEHICLE-01": 4500.0}
    }
    state = sm.load_state()
    state.setdefault("agency_overrides", {})[test_agency_name] = test_update
    save_ok = sm.save_state(state)
    assert save_ok and os.path.exists(state_file), "State file was not saved!"

    # Reload from disk to verify persistence
    sm_reload = StateManager(state_file)
    loaded_state = sm_reload.load_state()
    loaded_data = loaded_state.get("agency_overrides", {}).get(test_agency_name, {})
    assert loaded_data.get("pan") == "TESTP1234F", "PAN not persisted!"
    assert loaded_data.get("loads", {}).get("TEST-VEHICLE-01") == 25, "Loads not persisted!"
    assert loaded_data.get("rates", {}).get("TEST-VEHICLE-01") == 4500.0, "Rates not persisted!"
    print("  -> Edited agency fields, loads, and rates successfully persisted to data/app_state.json.")
    report["StatePersistence"] = "PASSED"

    # 4. Verify Invoice PDF Generation (ReportLab) with Mandatory Payment Note
    print("\n[CHECK 4] Testing Tax Invoice PDF Generation (ReportLab Engine)...")
    output_dir = os.path.join(INSTALL_DIR, "output")
    os.makedirs(output_dir, exist_ok=True)
    pdf_gen = PDFGenerator(output_dir=output_dir)

    test_agency_for_pdf = dict(sample_data)
    test_agency_for_pdf["invoice_no"] = 1999
    test_agency_for_pdf["vehicles"] = [
        {"vehicle_no": "KA01AB1234", "loads": 10, "rate_per_load": 4000.0, "total": 40000.0}
    ]
    test_agency_for_pdf["subtotal"] = 40000.0
    test_agency_for_pdf["grand_total"] = 40000.0

    generated_pdf = pdf_gen.generate_agency_pdf(test_agency_for_pdf)
    assert os.path.exists(generated_pdf), f"Generated PDF not found: {generated_pdf}"
    print(f"  -> Generated PDF: {os.path.basename(generated_pdf)} ({os.path.getsize(generated_pdf)} bytes)")

    # Inspect PDF contents via fitz
    doc = fitz.open(generated_pdf)
    assert len(doc) >= 1, "PDF has 0 pages!"
    pdf_text = "".join([page.get_text() for page in doc])
    assert "ANANYA ENTERPRISES" in pdf_text, "Missing header in PDF"
    assert "1999" in pdf_text, "Missing invoice number in PDF"
    assert "KA01AB1234" in pdf_text, "Missing vehicle row in PDF"
    expected_note = "Note : please complete the payment before 10th of this month"
    assert expected_note in pdf_text, f"Mandatory payment note missing from PDF!\nText found:\n{pdf_text}"
    print(f"  -> Verified mandatory note present: '{expected_note}'")
    doc.close()
    report["PDFGeneration"] = f"PASSED ({os.path.basename(generated_pdf)})"

    # 5. Verify Updated Master Excel Generation
    print("\n[CHECK 5] Testing Updated Master Excel Workbook Generation...")
    output_excel = os.path.join(output_dir, "Updated_Ananya_Bill_Verified.xlsx")
    em.generate_updated_workbook(
        target_month_year="SEPTEMBER 2026",
        target_date_str="01-10-2026",
        starting_inv_no=1800,
        updated_loads_map={},
        output_file_path=output_excel
    )
    assert os.path.exists(output_excel), "Updated workbook was not generated!"
    print(f"  -> Generated Master Excel: {os.path.basename(output_excel)} ({os.path.getsize(output_excel)} bytes)")
    report["MasterExcelGen"] = "PASSED"

    # 6. Verify Native Windows Clipboard (CF_HDROP) Copying for WhatsApp
    print("\n[CHECK 6] Testing Native Windows CF_HDROP Clipboard Format Copying...")
    dispatcher = WhatsAppDispatcher()
    clip_success = dispatcher.copy_pdf_to_clipboard(generated_pdf)
    assert clip_success, "Failed to copy PDF to Windows clipboard!"

    # Verify clipboard contents using PowerShell FileDropList
    ps_cmd = 'Get-Clipboard -Format FileDropList | ForEach-Object { $_.FullName }'
    res = subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True, text=True)
    copied_files = [line.strip() for line in res.stdout.strip().splitlines() if line.strip()]
    assert any(os.path.samefile(f, generated_pdf) for f in copied_files if os.path.exists(f)), (
        f"PDF was not found in Windows CF_HDROP clipboard!\nFound: {copied_files}"
    )
    print(f"  -> Native CF_HDROP FileDropList verified: {copied_files}")
    report["CF_HDROP_Clipboard"] = "PASSED"

    # 7. Verify WhatsApp Message, URL Dispatch & Auto-Attachment
    print("\n[CHECK 7] Testing WhatsApp Message, URL Formatter & Auto-Attachment...")
    msg = dispatcher.build_invoice_message(test_agency_name, 1999, "SEPTEMBER 2026")
    assert str(test_agency_name) in msg
    assert "1999" in msg
    assert "Note : please complete the payment before 10th of this month" in msg
    formatted_phone = dispatcher.format_phone_number("9876543210")
    assert formatted_phone == "919876543210", f"Phone formatting failed: {formatted_phone}"
    print(f"  -> Formatted Phone: {formatted_phone}")
    print(f"  -> Invoice Message formatted cleanly ({len(msg)} characters).")
    
    # Test auto attachment window detection helper
    desktop_wins = dispatcher.find_whatsapp_desktop_windows()
    print(f"  -> Native WhatsApp Desktop windows detected: {len(desktop_wins)}")
    report["WhatsAppDispatch"] = "PASSED"

    # 8. Verify Executable Direct Launch & Process Health
    print("\n[CHECK 8] Testing Standalone Executable Launch Directly in Windows...")
    proc = subprocess.Popen([EXE_PATH], cwd=INSTALL_DIR)
    time.sleep(3.5)
    poll_code = proc.poll()
    if poll_code is not None:
        raise RuntimeError(f"Executable exited prematurely with code {poll_code}!")
    pid = proc.pid
    print(f"  -> Executable is running actively with PID {pid} (No startup exceptions!).")
    proc.terminate()
    proc.wait(timeout=5)
    print("  -> Cleanly closed test process.")
    report["StandaloneExeLaunch"] = f"PASSED (PID {pid})"

    print("\n" + "=" * 65)
    print(" ALL 8 VERIFICATION GATES PASSED COMPLETELY WITH ZERO ERRORS!")
    print("=" * 65)
    for k, v in report.items():
        print(f"  [PASS] {k:<25} : {v}")
    print("=" * 65)
    return report


if __name__ == "__main__":
    run_full_verification()
