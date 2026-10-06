import os
import sys
import json
import zipfile
import tempfile
import unittest

# Ensure repo root is in python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from installer_app import (
    check_existing_installation,
    perform_installation,
    is_app_running,
    APP_VERSION,
    EXE_NAME
)


class TestInstallerUpdater(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.TemporaryDirectory()
        self.target_dir = os.path.join(self.test_dir.name, "InstalledApp")
        self.zip_file = os.path.join(self.test_dir.name, "test_pkg.zip")

        # Create a mock zip package with new version files
        with zipfile.ZipFile(self.zip_file, "w") as zf:
            zf.writestr(f"AnanyaInvoiceAutomation/{EXE_NAME}", "fake-binary-content-v1.0.1")
            zf.writestr("AnanyaInvoiceAutomation/data/app_state.json", '{"default_state": true}')
            zf.writestr("AnanyaInvoiceAutomation/Ananya Bill.xlsm", "default-excel-content")
            zf.writestr("AnanyaInvoiceAutomation/new_asset.dll", "new-dll-content")

    def tearDown(self):
        self.test_dir.cleanup()

    def test_fresh_installation_creates_files_and_version_json(self):
        # Target dir is currently empty
        is_inst, ver = check_existing_installation(self.target_dir)
        self.assertFalse(is_inst)
        self.assertIsNone(ver)

        # Perform install
        exe_path = perform_installation(
            zip_path=self.zip_file,
            target_dir=self.target_dir,
            create_desktop=False,
            create_start=False
        )

        self.assertTrue(os.path.exists(exe_path))
        self.assertTrue(os.path.exists(os.path.join(self.target_dir, "new_asset.dll")))

        # Check version.json was created
        version_file = os.path.join(self.target_dir, "version.json")
        self.assertTrue(os.path.exists(version_file))
        with open(version_file, "r", encoding="utf-8") as f:
            v_data = json.load(f)
            self.assertEqual(v_data.get("version"), APP_VERSION)

    def test_update_preserves_existing_user_data_and_template(self):
        # Simulate pre-existing installation with custom data
        os.makedirs(os.path.join(self.target_dir, "data"), exist_ok=True)
        custom_state_file = os.path.join(self.target_dir, "data", "app_state.json")
        with open(custom_state_file, "w", encoding="utf-8") as f:
            f.write('{"user_custom_agency": "ABC Logistics", "custom_rate": 5000}')

        custom_excel = os.path.join(self.target_dir, "Ananya Bill.xlsm")
        with open(custom_excel, "w", encoding="utf-8") as f:
            f.write("user-modified-excel-content")

        old_exe = os.path.join(self.target_dir, EXE_NAME)
        with open(old_exe, "w", encoding="utf-8") as f:
            f.write("old-binary-content-v1.0.0")

        # Verify detection
        is_inst, ver = check_existing_installation(self.target_dir)
        self.assertTrue(is_inst)

        # Run perform_installation (Update)
        exe_path = perform_installation(
            zip_path=self.zip_file,
            target_dir=self.target_dir,
            create_desktop=False,
            create_start=False
        )

        # 1. Binary must be updated to new version
        with open(exe_path, "r", encoding="utf-8") as f:
            self.assertEqual(f.read(), "fake-binary-content-v1.0.1")

        # 2. User data must NOT be overwritten!
        with open(custom_state_file, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("ABC Logistics", content)
            self.assertIn("5000", content)
            self.assertNotIn("default_state", content)

        # 3. User customized excel template must NOT be overwritten!
        with open(custom_excel, "r", encoding="utf-8") as f:
            self.assertEqual(f.read(), "user-modified-excel-content")

        # 4. New asset was added
        self.assertTrue(os.path.exists(os.path.join(self.target_dir, "new_asset.dll")))

    def test_update_overwrites_readonly_dlls(self):
        import stat
        os.makedirs(self.target_dir, exist_ok=True)
        readonly_dll = os.path.join(self.target_dir, "new_asset.dll")
        with open(readonly_dll, "w") as f:
            f.write("old-readonly-dll")
        # Mark as read-only (which normally causes PermissionError on Windows)
        os.chmod(readonly_dll, stat.S_IREAD)

        # Installation should safely overcome read-only attribute and overwrite
        exe_path = perform_installation(
            zip_path=self.zip_file,
            target_dir=self.target_dir,
            create_desktop=False,
            create_start=False
        )

        with open(readonly_dll, "r") as f:
            self.assertEqual(f.read(), "new-dll-content")


if __name__ == "__main__":
    unittest.main()

