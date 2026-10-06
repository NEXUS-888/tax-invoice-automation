import os
import sys
import unittest
import tempfile
import shutil

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from contacts_manager import ContactsManager


class TestContactsManager(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.contacts_path = os.path.join(self.test_dir, "contacts.json")
        self.mgr = ContactsManager(self.contacts_path)

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_parse_phone_list(self):
        # Clean 10-digit
        self.assertEqual(self.mgr.parse_phone_list("9141844840"), ["919141844840"])
        # Clean 12-digit with 91
        self.assertEqual(self.mgr.parse_phone_list("919141844840"), ["919141844840"])
        # Multiple separated by comma and slash
        phones = self.mgr.parse_phone_list("9141844840, 9880183840 / 7353306462")
        self.assertEqual(phones, ["919141844840", "919880183840", "917353306462"])
        # Duplicate removal
        phones_dup = self.mgr.parse_phone_list("9141844840, 9141844840")
        self.assertEqual(phones_dup, ["919141844840"])
        # Invalid numbers
        self.assertEqual(self.mgr.parse_phone_list("12345, abc, 0000000000"), [])
        # Leading trunk 0 / 0091 used to be silently dropped, leaving the agency with no number
        self.assertEqual(self.mgr.parse_phone_list("09141844840; 0091 98801 83840"),
                         ["919141844840", "919880183840"])

    def test_whitespace_insensitive_lookup(self):
        # Update with trailing whitespace in sheet name
        self.mgr.update_phone("VINAYAKA BHARATGAS ", "9141844840", "VINAYAKA BHARAT GAS")

        # Retrieve with exact
        self.assertEqual(self.mgr.get_phones("VINAYAKA BHARATGAS "), ["919141844840"])
        # Retrieve without trailing space
        self.assertEqual(self.mgr.get_phones("VINAYAKA BHARATGAS"), ["919141844840"])
        # Retrieve case-insensitive
        self.assertEqual(self.mgr.get_phones("vinayaka bharatgas"), ["919141844840"])

    def test_sync_agencies(self):
        sample_agencies = [
            {'sheet_name': 'AGENCY A', 'agency_name': 'Agency Alpha'},
            {'sheet_name': 'AGENCY B ', 'agency_name': 'Agency Beta'},
        ]
        self.mgr.sync_agencies(sample_agencies)
        self.assertIn('AGENCY A', self.mgr.contacts)
        # Check lookup works for both
        self.assertEqual(self.mgr.get_phones('AGENCY A'), [])
        self.assertEqual(self.mgr.get_phones('AGENCY B'), [])


if __name__ == "__main__":
    unittest.main()
