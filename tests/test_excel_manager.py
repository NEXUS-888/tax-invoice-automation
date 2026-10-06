import os
import sys
import unittest
import tempfile
import shutil
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from excel_manager import ExcelManager, num_to_words_indian


class TestExcelManager(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.master_file = os.path.join(os.path.dirname(__file__), "..", "Ananya Bill.xlsm")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_num_to_words_indian(self):
        self.assertEqual(num_to_words_indian(0), "Rupees Zero Only")
        self.assertEqual(num_to_words_indian(1), "Rupees One Only")
        self.assertEqual(num_to_words_indian(15), "Rupees Fifteen Only")
        self.assertEqual(num_to_words_indian(100), "Rupees One Hundred Only")
        self.assertEqual(num_to_words_indian(1000), "Rupees One Thousand Only")
        self.assertEqual(num_to_words_indian(100000), "Rupees One Lakh Only")
        self.assertEqual(num_to_words_indian(10000000), "Rupees One Crore Only")
        self.assertEqual(num_to_words_indian(14868), "Rupees Fourteen Thousand Eight Hundred Sixty Eight Only")
        self.assertEqual(num_to_words_indian(12637.80), "Rupees Twelve Thousand Six Hundred Thirty Seven Paise Eighty Only")

    def test_parse_all_agencies(self):
        mgr = ExcelManager(self.master_file)
        agencies = mgr.parse_all_agencies()
        self.assertEqual(len(agencies), 46)

        # Check structure of first agency
        first = agencies[0]
        self.assertIn('sheet_name', first)
        self.assertIn('agency_name', first)
        self.assertIn('vehicles', first)
        self.assertIn('grand_total', first)
        self.assertGreater(len(first['vehicles']), 0)

    def test_generate_updated_workbook_deduplication(self):
        mgr = ExcelManager(self.master_file)
        out_excel = os.path.join(self.test_dir, "test_out.xlsx")

        # Pass loads with both stripped and unstripped keys to test deduplication
        loads_map = {
            'VINAYAKA BHARATGAS': [{'vehicle_no': 'KA 27 A 9934', 'loads': 25.0, 'rate': 630.0}],
            'ADVAITA': [{'vehicle_no': 'KA 19 D 5917', 'loads': 15.0, 'rate': 630.0}]
        }

        updated = mgr.generate_updated_workbook(
            target_month_year="SEPTEMBER 2026",
            target_date_str="05-09-2026",
            starting_inv_no=2100,
            updated_loads_map=loads_map,
            output_file_path=out_excel
        )

        # Ensure exact agency count: 46 (no duplicate sheets created!)
        self.assertEqual(len(updated), 46)

        wb = openpyxl.load_workbook(out_excel, data_only=True)
        # Check sheet count excluding Sheet1
        agency_sheets = [s for s in wb.sheetnames if s != 'Sheet1']
        self.assertEqual(len(agency_sheets), 46)

        # Check that Vinayaka got the updated load
        vinayaka = next(a for a in updated if 'VINAYAKA' in a['sheet_name'])
        self.assertEqual(vinayaka['vehicles'][0]['loads'], 25.0)

        # Check that Note is written in cell A35
        for s in agency_sheets[:5]:
            ws = wb[s]
            self.assertIn("Note : please complete the payment before 10th of this month", str(ws['A35'].value))

    def test_build_agency_invoice_data(self):
        mgr = ExcelManager(self.master_file)
        data = mgr.build_agency_invoice_data(
            sheet_name="MAHALASA",
            target_month_year="OCTOBER 2026",
            target_date_str="05-10-2026",
            invoice_no=1850,
            vehicles_list=[{'vehicle_no': 'KA 19 D 5917', 'loads': 20.0, 'rate': 550.0}]
        )

        self.assertEqual(data['sheet_name'], "MAHALASA")
        self.assertEqual(data['invoice_no'], 1850)
        self.assertEqual(data['date'], "05-10-2026")
        self.assertIn("OCTOBER 2026", data['month_desc'])
        self.assertEqual(data['total'], 11000.0)
        self.assertEqual(data['sgst'], 990.0)
        self.assertEqual(data['cgst'], 990.0)
        self.assertEqual(data['grand_total'], 12980.0)
        self.assertIn("Twelve Thousand", data['in_words'])


if __name__ == "__main__":
    unittest.main()
