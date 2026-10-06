#!/usr/bin/env python3
"""
Build Script for Ananya Enterprises - Invoice Automation & WhatsApp Dispatcher
Compiles the application into a standalone Windows executable using PyInstaller
and packages it into a distribution zip for GitHub Releases.
"""

import os
import sys
import shutil
import subprocess
import zipfile

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DIST_DIR = os.path.join(BASE_DIR, "dist")
BUILD_DIR = os.path.join(BASE_DIR, "build")
APP_DIST_DIR = os.path.join(DIST_DIR, "AnanyaInvoiceAutomation")
ZIP_NAME = "Ananya_Invoice_Automation_Windows.zip"
ZIP_PATH = os.path.join(DIST_DIR, ZIP_NAME)

README_CONTENT = """===============================================================
ANANYA ENTERPRISES - INVOICE AUTOMATION & WHATSAPP DISPATCHER
===============================================================

Standalone Windows Distribution
Version: 1.0.9

--- QUICK START GUIDE ---

1. HOW TO LAUNCH:
   - Double-click 'AnanyaInvoiceAutomation.exe'.
   - No Python or technical installation is required!

2. MASTER EXCEL FILE:
   - 'Ananya Bill.xlsm' is included in this folder.
   - You can replace or browse any updated master Excel sheet inside the app.

3. EDITING LOADS, RATES & AGENCIES:
   - Tab 1 ("Agency & Load Manager"):
     * Double-click any vehicle rate (Column 2) or load count (Column 4) to edit.
     * Click '✏️ Edit Agency Info' to customize agency name, PAN, GST, address, and WhatsApp phone numbers.
     * Click '➕ Add New Agency' or '➕ Add Vehicle' as needed.
     * All changes are saved permanently in the 'data' folder and will NEVER revert on restart!

4. WHATSAPP WEB AUTOMATION:
   - Click '📱 Connect WhatsApp Account' in the top bar.
   - An interactive browser window will open.
   - Scan the QR code once with your phone's WhatsApp.
   - Your session remains saved securely in 'data/wa_session' on your PC!

5. GENERATING & SENDING INVOICES:
   - Tab 3 ("Generate & Auto-Send WhatsApp Invoices"):
     * Click '1. Generate All PDFs & Excel' to batch-produce all bills.
     * Click '2. Automated Send ALL Invoices via WhatsApp' for hands-free bulk delivery.
     * Or use '📤 Share App' / '📱 Auto-Send' next to any individual agency.
   - All generated PDF bills are saved in the 'output' folder.

For assistance or updates, check the GitHub repository:
https://github.com/NEXUS-888/tax-invoice-automation
===============================================================
"""


def clean_previous_builds():
    print("[1/4] Cleaning previous build artifacts...")
    for path in [DIST_DIR, BUILD_DIR]:
        if os.path.exists(path):
            try:
                shutil.rmtree(path)
                print(f"  Removed: {path}")
            except Exception as e:
                print(f"  Warning removing {path}: {e}")


def run_pyinstaller():
    print("\n[2/4] Running PyInstaller compiler...")
    spec_file = os.path.join(BASE_DIR, "AnanyaInvoiceAutomation.spec")
    if not os.path.exists(spec_file):
        raise FileNotFoundError(f"Spec file not found: {spec_file}")

    cmd = [sys.executable, "-m", "PyInstaller", "--clean", "-y", spec_file]
    print(f"  Command: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode != 0:
        raise RuntimeError(f"PyInstaller build failed with exit code {result.returncode}")
    print("  PyInstaller build completed successfully.")


def setup_distribution_folder():
    print("\n[3/4] Preparing distribution folder...")
    if not os.path.exists(APP_DIST_DIR):
        raise FileNotFoundError(f"Expected build folder not found: {APP_DIST_DIR}")

    # Ensure master Excel is present
    master_src = os.path.join(BASE_DIR, "Ananya Bill.xlsm")
    master_dest = os.path.join(APP_DIST_DIR, "Ananya Bill.xlsm")
    if os.path.exists(master_src) and not os.path.exists(master_dest):
        shutil.copy2(master_src, master_dest)
        print("  Copied Ananya Bill.xlsm template")

    # Create empty data and output directories
    os.makedirs(os.path.join(APP_DIST_DIR, "data"), exist_ok=True)
    os.makedirs(os.path.join(APP_DIST_DIR, "output"), exist_ok=True)

    # Write README
    readme_path = os.path.join(APP_DIST_DIR, "README_SETUP.txt")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(README_CONTENT)
    print("  Created README_SETUP.txt")


INSTALLER_EXE_NAME = "Ananya_Invoice_Automation_Setup.exe"
INSTALLER_EXE_PATH = os.path.join(DIST_DIR, INSTALLER_EXE_NAME)


def package_zip():
    print(f"\n[4/5] Creating distribution ZIP package: {ZIP_NAME}...")
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(APP_DIST_DIR):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, DIST_DIR)
                zf.write(full_path, rel_path)
    zip_size_mb = os.path.getsize(ZIP_PATH) / (1024 * 1024)
    print(f"  Success! Package created at:\n  {ZIP_PATH} ({zip_size_mb:.2f} MB)")


def build_installer_exe():
    print(f"\n[5/5] Building One-Click Setup Installer: {INSTALLER_EXE_NAME}...")
    if not os.path.exists(ZIP_PATH):
        raise FileNotFoundError(f"Zip archive not found for installer bundling: {ZIP_PATH}")

    installer_script = os.path.join(BASE_DIR, "installer_app.py")
    sep = ";" if os.name == "nt" else ":"
    data_arg = f"{ZIP_PATH}{sep}."

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--clean",
        "--noconsole",
        "--onefile",
        "--name",
        "Ananya_Invoice_Automation_Setup",
        "--add-data",
        data_arg,
        "--distpath",
        DIST_DIR,
        installer_script
    ]
    print(f"  Command: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode != 0:
        raise RuntimeError(f"Installer PyInstaller build failed with exit code {result.returncode}")

    if os.path.exists(INSTALLER_EXE_PATH):
        exe_size_mb = os.path.getsize(INSTALLER_EXE_PATH) / (1024 * 1024)
        print(f"  Success! One-Click Installer created at:\n  {INSTALLER_EXE_PATH} ({exe_size_mb:.2f} MB)")


def main():
    print("======================================================")
    print(" Building Standalone Ananya Invoice Automation (.exe)")
    print("======================================================")
    clean_previous_builds()
    run_pyinstaller()
    setup_distribution_folder()
    package_zip()
    build_installer_exe()
    print("\n[DONE] Build and packaging complete!")
    print(f"1-Click Setup Installer: {INSTALLER_EXE_PATH}")
    print(f"Standalone ZIP Package:  {ZIP_PATH}")
    print(f"Extracted App Folder:    {APP_DIST_DIR}")


if __name__ == "__main__":
    main()

