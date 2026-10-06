#!/usr/bin/env python3
"""
Ananya Enterprises - One-Click Windows Setup Installer & Updater
Extracts application files, creates Desktop & Start Menu shortcuts, and launches the app.
Automatically detects existing installations, offers seamless 1-click update,
and safeguards existing user data, rates, agencies, and WhatsApp session.
"""

import os
import sys
import json
import time
import shutil
import zipfile
import tempfile
import subprocess
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_NAME = "Ananya Invoice Automation"
APP_VERSION = "1.0.7"
PUBLISHER = "Ananya Enterprises"
EXE_NAME = "AnanyaInvoiceAutomation.exe"
ZIP_FILENAME = "Ananya_Invoice_Automation_Windows.zip"


def get_default_install_dir():
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        local_app_data = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(local_app_data, "Programs", "AnanyaInvoiceAutomation")


def find_zip_package():
    # 1. PyInstaller onefile temp folder
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        candidate = os.path.join(bundle_dir, ZIP_FILENAME)
        if os.path.exists(candidate):
            return candidate

    # 2. Alongside the installer executable / script
    base_dir = os.path.abspath(os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__))
    candidate = os.path.join(base_dir, ZIP_FILENAME)
    if os.path.exists(candidate):
        return candidate

    # 3. In dist/ folder relative to repo root (for dev testing)
    candidate = os.path.join(base_dir, "dist", ZIP_FILENAME)
    if os.path.exists(candidate):
        return candidate

    return None


def is_app_running(target_dir=None, exe_name=EXE_NAME):
    """Checks if the application executable or any process in target_dir is currently running."""
    if os.name != "nt":
        return False

    # 1. Check using CSV format to prevent table column truncation of long names (>25 chars)
    try:
        output_csv = subprocess.check_output(
            f'tasklist /FI "IMAGENAME eq {exe_name}" /FO CSV /NH',
            shell=True, stderr=subprocess.DEVNULL
        ).decode(errors="ignore")
        if exe_name.lower() in output_csv.lower():
            return True
    except Exception:
        pass

    # 2. Check prefix in standard tasklist as fallback
    try:
        output_raw = subprocess.check_output(
            'tasklist /NH', shell=True, stderr=subprocess.DEVNULL
        ).decode(errors="ignore")
        prefix = exe_name[:20].lower()
        if prefix in output_raw.lower():
            return True
    except Exception:
        pass

    # 3. If target_dir is provided, check if ANY process is running from that directory
    if target_dir and os.path.exists(target_dir):
        try:
            norm_dir = os.path.normpath(target_dir).replace("'", "''")
            ps_cmd = f"(Get-CimInstance Win32_Process | Where-Object {{ $_.ExecutablePath -like '{norm_dir}*' }}).Count"
            res = subprocess.check_output(["powershell", "-NoProfile", "-Command", ps_cmd], text=True, stderr=subprocess.DEVNULL).strip()
            if res.isdigit() and int(res) > 0:
                return True
        except Exception:
            pass

    return False


def close_running_app(target_dir=None, exe_name=EXE_NAME):
    """Terminates the running application process and any child processes in target_dir cleanly."""
    if os.name != "nt":
        return True

    # 1. Kill by image name pattern (including truncated name)
    patterns = [exe_name, "AnanyaInvoiceAutomation*"]
    for pat in patterns:
        try:
            subprocess.run(
                f'taskkill /F /IM "{pat}" /T',
                shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        except Exception:
            pass

    # 2. Kill any processes running from target_dir via PowerShell
    if target_dir and os.path.exists(target_dir):
        try:
            norm_dir = os.path.normpath(target_dir).replace("'", "''")
            ps_cmd = f"Get-CimInstance Win32_Process | Where-Object {{ $_.ExecutablePath -like '{norm_dir}*' }} | ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force }}"
            subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

    # 3. Wait for process handles to release (up to 3 seconds)
    for _ in range(10):
        if not is_app_running(target_dir, exe_name):
            return True
        time.sleep(0.3)

    return not is_app_running(target_dir, exe_name)


def safe_write_file(src_stream, dest_path, max_retries=4):
    """
    Safely writes data from src_stream to dest_path with retry, attribute reset,
    and locked DLL renaming fallback (the standard Windows installer strategy).
    """
    import stat
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)

    # If file exists, ensure write permissions (clear read-only flag)
    if os.path.exists(dest_path):
        try:
            os.chmod(dest_path, stat.S_IWRITE | stat.S_IREAD)
        except Exception:
            pass

    # Read data from source stream
    data = src_stream.read()

    last_err = None
    for attempt in range(max_retries):
        try:
            with open(dest_path, "wb") as dst:
                dst.write(data)
            return True
        except PermissionError as e:
            last_err = e
            time.sleep(0.4)
            # Try clearing attributes again
            try:
                os.chmod(dest_path, stat.S_IWRITE)
            except Exception:
                pass

    # If direct write still failed due to file lock:
    # Use Windows rename trick: rename locked file to .old.<timestamp>, then write new file!
    try:
        old_temp_path = f"{dest_path}.old.{int(time.time()*1000)}"
        if os.path.exists(dest_path):
            os.replace(dest_path, old_temp_path)
        with open(dest_path, "wb") as dst:
            dst.write(data)
        return True
    except Exception as e2:
        raise PermissionError(f"Failed to write '{dest_path}' after retries and rename fallback: {last_err or e2}")


def check_existing_installation(target_dir):
    """
    Checks if an existing installation is present in target_dir.
    Returns (is_installed: bool, version_str: str or None).
    """
    exe_file = os.path.join(target_dir, EXE_NAME)
    if not os.path.exists(exe_file):
        return False, None

    version_file = os.path.join(target_dir, "version.json")
    if os.path.exists(version_file):
        try:
            with open(version_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return True, data.get("version", "1.0.0")
        except Exception:
            return True, "1.0.0"

    return True, "1.0.0"


def create_windows_shortcut(target_exe, shortcut_name, working_dir, folder_type="Desktop", desc=""):
    """
    Creates a Windows .lnk shortcut using WScript.Shell via a temporary VBScript.
    folder_type can be 'Desktop' or 'Programs'.
    """
    vbs_content = (
        'Set oWS = WScript.CreateObject("WScript.Shell")\r\n'
        f'targetFolder = oWS.SpecialFolders("{folder_type}")\r\n'
        f'Set oLink = oWS.CreateShortcut(targetFolder & "\\{shortcut_name}")\r\n'
        f'oLink.TargetPath = "{target_exe}"\r\n'
        f'oLink.WorkingDirectory = "{working_dir}"\r\n'
        f'oLink.Description = "{desc}"\r\n'
        'oLink.Save\r\n'
    )
    with tempfile.NamedTemporaryFile("w", suffix=".vbs", delete=False) as f:
        f.write(vbs_content)
        vbs_path = f.name

    try:
        flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        subprocess.run(["cscript", "//nologo", vbs_path], check=True, creationflags=flags)
        return True
    except Exception as e:
        print(f"Error creating shortcut: {e}")
        return False
    finally:
        if os.path.exists(vbs_path):
            try:
                os.remove(vbs_path)
            except Exception:
                pass


def perform_installation(zip_path, target_dir, create_desktop=True, create_start=True, progress_callback=None):
    """
    Extracts files from zip_path into target_dir.
    Safely preserves existing user data, rates, agencies, and WhatsApp session.
    """
    os.makedirs(target_dir, exist_ok=True)
    os.makedirs(os.path.join(target_dir, "data"), exist_ok=True)
    os.makedirs(os.path.join(target_dir, "output"), exist_ok=True)

    # Ensure any lingering processes in target_dir are cleanly closed before overwriting binaries
    close_running_app(target_dir)

    with zipfile.ZipFile(zip_path, "r") as zf:
        members = zf.infolist()
        total_files = len(members)

        for idx, member in enumerate(members):
            filename = member.filename
            # Strip top-level directory if present
            if filename.startswith("AnanyaInvoiceAutomation/") or filename.startswith("AnanyaInvoiceAutomation\\"):
                rel_parts = filename.split("/", 1) if "/" in filename else filename.split("\\", 1)
                rel_path = rel_parts[1] if len(rel_parts) > 1 else ""
            else:
                rel_path = filename

            if not rel_path or rel_path.endswith("/") or rel_path.endswith("\\"):
                continue

            dest_path = os.path.join(target_dir, rel_path)

            # Preserve existing user state, WhatsApp session, and modified Excel file on update
            is_user_data = (
                rel_path.startswith("data/") or
                rel_path.startswith("data\\") or
                rel_path.startswith("output/") or
                rel_path.startswith("output\\") or
                rel_path.lower() == "ananya bill.xlsm"
            )
            if is_user_data and os.path.exists(dest_path):
                # Preserve existing data file
                continue

            os.makedirs(os.path.dirname(dest_path), exist_ok=True)

            with zf.open(member) as src:
                safe_write_file(src, dest_path)

            if progress_callback:
                progress = int(((idx + 1) / total_files) * 85)
                progress_callback(progress, f"Updating {os.path.basename(rel_path)}...")

    # Clean up any temporary .old.* files left over from in-use renames
    try:
        for root, dirs, files in os.walk(target_dir):
            for f in files:
                if ".old." in f:
                    try:
                        os.remove(os.path.join(root, f))
                    except Exception:
                        pass
    except Exception:
        pass

    # Write or update version metadata
    version_info = {
        "app_name": APP_NAME,
        "version": APP_VERSION,
        "installed_at": datetime.now().isoformat()
    }
    try:
        with open(os.path.join(target_dir, "version.json"), "w", encoding="utf-8") as vf:
            json.dump(version_info, vf, indent=2)
    except Exception:
        pass

    exe_path = os.path.join(target_dir, EXE_NAME)

    if create_desktop:
        if progress_callback:
            progress_callback(90, "Creating Desktop shortcut...")
        create_windows_shortcut(
            target_exe=exe_path,
            shortcut_name=f"{APP_NAME}.lnk",
            working_dir=target_dir,
            folder_type="Desktop",
            desc="Ananya Enterprises Invoice Automation & WhatsApp Dispatcher"
        )

    if create_start:
        if progress_callback:
            progress_callback(95, "Creating Start Menu shortcut...")
        create_windows_shortcut(
            target_exe=exe_path,
            shortcut_name=f"{APP_NAME}.lnk",
            working_dir=target_dir,
            folder_type="Programs",
            desc="Ananya Enterprises Invoice Automation & WhatsApp Dispatcher"
        )

    if progress_callback:
        progress_callback(100, "Complete!")

    return exe_path


class SetupWizard(tk.Tk):
    def __init__(self, zip_path):
        super().__init__()
        self.zip_path = zip_path
        self.is_update = False
        self.title(f"{APP_NAME} - Setup Wizard")
        self.geometry("560x450")
        self.resizable(False, False)

        # Style
        self.configure(bg="#F8FAFC")

        # Top Header Banner
        self.header_frame = tk.Frame(self, bg="#1E3A8A", height=78)
        self.header_frame.pack(fill=tk.X, side=tk.TOP)
        self.header_frame.pack_propagate(False)

        self.title_lbl = tk.Label(
            self.header_frame,
            text=APP_NAME,
            font=("Segoe UI", 15, "bold"),
            fg="#FFFFFF",
            bg="#1E3A8A"
        )
        self.title_lbl.pack(anchor="w", padx=20, pady=(10, 2))

        self.sub_lbl = tk.Label(
            self.header_frame,
            text=f"One-Click Installation Wizard (Version {APP_VERSION})",
            font=("Segoe UI", 9),
            fg="#93C5FD",
            bg="#1E3A8A"
        )
        self.sub_lbl.pack(anchor="w", padx=20)

        # Main Body Frame
        self.body_frame = tk.Frame(self, bg="#F8FAFC", padx=24, pady=16)
        self.body_frame.pack(fill=tk.BOTH, expand=True)

        self.info_lbl = tk.Label(
            self.body_frame,
            text="The wizard will install Ananya Invoice Automation onto your computer.",
            font=("Segoe UI", 9),
            fg="#334155",
            bg="#F8FAFC",
            justify="left",
            wraplength=500
        )
        self.info_lbl.pack(anchor="w", pady=(0, 12))

        # Target directory group
        dir_lbl = tk.Label(
            self.body_frame,
            text="Installation Folder:",
            font=("Segoe UI", 9, "bold"),
            fg="#1E293B",
            bg="#F8FAFC"
        )
        dir_lbl.pack(anchor="w", pady=(0, 4))

        dir_input_frame = tk.Frame(self.body_frame, bg="#F8FAFC")
        dir_input_frame.pack(fill=tk.X, pady=(0, 14))

        self.dir_var = tk.StringVar(value=get_default_install_dir())
        self.dir_entry = ttk.Entry(dir_input_frame, textvariable=self.dir_var, font=("Segoe UI", 9))
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        self.dir_var.trace_add("write", lambda *args: self.refresh_install_mode())

        browse_btn = ttk.Button(dir_input_frame, text="Browse...", command=self.on_browse)
        browse_btn.pack(side=tk.RIGHT, padx=(8, 0))

        # Options
        self.desktop_var = tk.BooleanVar(value=True)
        self.start_var = tk.BooleanVar(value=True)
        self.launch_var = tk.BooleanVar(value=True)

        chk_desktop = ttk.Checkbutton(
            self.body_frame,
            text="Create / Update Desktop Shortcut (Recommended)",
            variable=self.desktop_var
        )
        chk_desktop.pack(anchor="w", pady=3)

        chk_start = ttk.Checkbutton(
            self.body_frame,
            text="Create / Update Start Menu Shortcut",
            variable=self.start_var
        )
        chk_start.pack(anchor="w", pady=3)

        self.chk_launch = ttk.Checkbutton(
            self.body_frame,
            text="Launch Ananya Invoice Automation immediately after install",
            variable=self.launch_var
        )
        self.chk_launch.pack(anchor="w", pady=(3, 14))

        # Progress bar
        self.status_var = tk.StringVar(value="Ready.")
        self.status_lbl = tk.Label(
            self.body_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 8),
            fg="#64748B",
            bg="#F8FAFC"
        )
        self.status_lbl.pack(anchor="w", pady=(0, 4))

        self.progress_bar = ttk.Progressbar(self.body_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 16))

        # Bottom Button Bar
        btn_frame = tk.Frame(self, bg="#E2E8F0", height=50)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        btn_frame.pack_propagate(False)

        self.cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self.destroy)
        self.cancel_btn.pack(side=tk.RIGHT, padx=(0, 16), pady=10)

        self.install_btn = ttk.Button(btn_frame, text="Install Now ➔", command=self.start_install)
        self.install_btn.pack(side=tk.RIGHT, padx=(0, 8), pady=10)

        # Initial check
        self.refresh_install_mode()

    def refresh_install_mode(self):
        target_dir = self.dir_var.get().strip()
        is_installed, old_ver = check_existing_installation(target_dir)
        self.is_update = is_installed

        if is_installed:
            old_label = f" (v{old_ver})" if old_ver else ""
            self.sub_lbl.config(
                text=f"Update / Upgrade Detected • Existing Installation{old_label} ➔ v{APP_VERSION}",
                fg="#FDE047"
            )
            self.info_lbl.config(
                text=(
                    f"Existing installation detected!\n"
                    f"Click 'Update Now' to upgrade to Version {APP_VERSION}. All your saved rates, "
                    f"agencies, WhatsApp login, and settings in 'data/' will be safely preserved."
                ),
                fg="#1E3A8A"
            )
            self.install_btn.config(text="Update Now ➔")
            self.chk_launch.config(text="Launch Ananya Invoice Automation immediately after update")
            self.status_var.set("Ready to update.")
            self.title(f"{APP_NAME} - Update to v{APP_VERSION}")
        else:
            self.sub_lbl.config(
                text=f"One-Click Installation Wizard (Version {APP_VERSION})",
                fg="#93C5FD"
            )
            self.info_lbl.config(
                text="The wizard will install Ananya Invoice Automation onto your computer.",
                fg="#334155"
            )
            self.install_btn.config(text="Install Now ➔")
            self.chk_launch.config(text="Launch Ananya Invoice Automation immediately after install")
            self.status_var.set("Ready to install.")
            self.title(f"{APP_NAME} - Setup Wizard")

    def on_browse(self):
        chosen = filedialog.askdirectory(initialdir=self.dir_var.get())
        if chosen:
            self.dir_var.set(os.path.normpath(chosen))

    def update_progress(self, percent, msg):
        self.progress_bar["value"] = percent
        self.status_var.set(msg)
        self.update_idletasks()

    def start_install(self):
        target_dir = self.dir_var.get().strip()
        if not target_dir:
            messagebox.showwarning("Warning", "Please specify a valid installation directory.")
            return

        # Check if the app is currently running and needs to be closed
        if is_app_running(target_dir):
            ans = messagebox.askyesno(
                "Application is Running",
                f"{APP_NAME} is currently running.\n\n"
                "The installer needs to close the running application to update files.\n\n"
                "Would you like to close it automatically and proceed?",
                icon="warning"
            )
            if not ans:
                return

            closed = close_running_app(target_dir)
            if not closed and is_app_running(target_dir):
                messagebox.showerror(
                    "Cannot Close Application",
                    f"Please close {APP_NAME} manually from your taskbar or Task Manager, then click '{self.install_btn.cget('text')}' again."
                )
                return

        self.install_btn.config(state="disabled")
        self.dir_entry.config(state="disabled")

        threading.Thread(target=self._run_install_worker, args=(target_dir,), daemon=True).start()

    def _run_install_worker(self, target_dir):
        try:
            exe_path = perform_installation(
                zip_path=self.zip_path,
                target_dir=target_dir,
                create_desktop=self.desktop_var.get(),
                create_start=self.start_var.get(),
                progress_callback=self.update_progress
            )

            self.after(0, self._install_success, exe_path)
        except Exception as e:
            self.after(0, self._install_error, str(e))

    def _install_success(self, exe_path):
        action_word = "Update" if self.is_update else "Installation"
        self.status_var.set(f"{action_word} Completed Successfully!")
        self.progress_bar["value"] = 100
        self.install_btn.config(text="Finished", state="normal", command=self.destroy)
        self.cancel_btn.pack_forget()

        if self.launch_var.get() and os.path.exists(exe_path):
            try:
                subprocess.Popen([exe_path], cwd=os.path.dirname(exe_path))
            except Exception:
                pass

        if self.is_update:
            messagebox.showinfo(
                "Update Successful",
                f"Ananya Invoice Automation was updated to version {APP_VERSION} successfully!\n\n"
                "All your existing agencies, rates, loads, and settings were preserved."
            )
        else:
            messagebox.showinfo(
                "Installation Successful",
                f"Ananya Invoice Automation was installed successfully!\n\n"
                "You can launch it anytime using the Desktop shortcut."
            )
        self.destroy()

    def _install_error(self, err_msg):
        self.status_var.set(f"Error: {err_msg}")
        self.install_btn.config(state="normal")
        self.dir_entry.config(state="normal")
        messagebox.showerror("Setup Failed", f"An error occurred during setup:\n\n{err_msg}")


def main():
    zip_path = find_zip_package()
    if not zip_path or not os.path.exists(zip_path):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing Installation Package",
            f"Could not locate the application distribution archive: {ZIP_FILENAME}\n\n"
            "Please ensure the installer was downloaded completely."
        )
        sys.exit(1)

    # Check for silent install flag
    if "/S" in sys.argv or "--silent" in sys.argv:
        target = get_default_install_dir()
        if is_app_running():
            close_running_app()
        exe = perform_installation(zip_path, target, create_desktop=True, create_start=True)
        if ("--launch" in sys.argv or "/LAUNCH" in sys.argv) and os.path.exists(exe):
            subprocess.Popen([exe], cwd=target)
        sys.exit(0)

    app = SetupWizard(zip_path)
    app.mainloop()


if __name__ == "__main__":
    main()
