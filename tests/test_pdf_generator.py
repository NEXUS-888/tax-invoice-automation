import os
import sys
import unittest
import tempfile
import shutil
import fitz

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from excel_manager import ExcelManager
from pdf_generator import PDFGenerator


class TestPDFGenerator(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.master_file = os.path.join(os.path.dirname(__file__), "..", "Ananya Bill.xlsm")
        self.em = ExcelManager(self.master_file)
        self.agencies = self.em.parse_all_agencies()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_single_pdf_generation_content_and_note(self):
        pdf_gen = PDFGenerator(self.test_dir)
        first_agency = self.agencies[0]
        pdf_path = pdf_gen.generate_agency_pdf(first_agency)

        self.assertTrue(os.path.exists(pdf_path))

        # Inspect PDF with PyMuPDF
        doc = fitz.open(pdf_path)
        self.assertEqual(len(doc), 1, "PDF must fit on exactly 1 page!")

        text = doc[0].get_text()
        self.assertIn("ANANYA ENTERPRISES", text)
        self.assertIn("BANK ACCOUNT DETAILS", text)
        self.assertIn("Note : please complete the payment before 10th of this month", text)

    def test_batch_generate_and_duplicate_cleanup(self):
        pdf_gen = PDFGenerator(self.test_dir)
        
        # Test first batch run
        paths = pdf_gen.batch_generate(self.agencies[:5], clean_output_dir=True)
        self.assertEqual(len(paths), 5)
        files_on_disk = [f for f in os.listdir(self.test_dir) if f.lower().endswith('.pdf')]
        self.assertEqual(len(files_on_disk), 5)

        # Test second batch run with changed invoice numbers
        for a in self.agencies[:5]:
            a['invoice_no'] = int(a['invoice_no'] or 100) + 500

        paths2 = pdf_gen.batch_generate(self.agencies[:5], clean_output_dir=True)
        self.assertEqual(len(paths2), 5)
        files_after_rerun = [f for f in os.listdir(self.test_dir) if f.lower().endswith('.pdf')]
        self.assertEqual(len(files_after_rerun), 5, "Old invoices must be cleaned; no duplicates allowed!")

    def test_date_formatting(self):
        pdf_gen = PDFGenerator(self.test_dir)
        agency_data = dict(self.agencies[0])
        # Test string YYYY-MM-DD
        agency_data['date'] = '2026-09-05'
        path = pdf_gen.generate_agency_pdf(agency_data)
        doc = fitz.open(path)
        self.assertIn("05-09-2026", doc[0].get_text())


if __name__ == "__main__":
    unittest.main()
