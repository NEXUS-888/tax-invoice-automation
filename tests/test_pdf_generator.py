import os
import sys
import unittest
import tempfile
import shutil
try:
    import fitz
except ImportError:
    fitz = None

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from excel_manager import ExcelManager
from unittest.mock import patch
import pdf_generator
from pdf_generator import PDFGenerator, default_signature_path


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

        # Inspect PDF with PyMuPDF if available
        if fitz:
            doc = fitz.open(pdf_path)
            self.assertEqual(len(doc), 1, "PDF must fit on exactly 1 page!")

            text = doc[0].get_text()
            self.assertIn("ANANYA ENTERPRISES", text)
            self.assertIn("BANK ACCOUNT DETAILS", text)
            self.assertIn("Note : please complete the payment before 10th of this month", text)
        else:
            self.assertGreater(os.path.getsize(pdf_path), 500)

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
        if fitz:
            doc = fitz.open(path)
            self.assertIn("05-09-2026", doc[0].get_text())
        else:
            self.assertTrue(os.path.exists(path))

    def test_signature_is_drawn_in_pdf(self):
        pdf_gen = PDFGenerator(self.test_dir)
        self.assertTrue(pdf_gen.has_signature, f"signature image missing: {pdf_gen.signature_path}")
        path = pdf_gen.generate_agency_pdf(dict(self.agencies[0]))
        if fitz:
            self.assertEqual(len(fitz.open(path)[0].get_images()), 1, "signature image must be on the page")

    def test_missing_signature_is_reported(self):
        pdf_gen = PDFGenerator(self.test_dir, signature_path=os.path.join(self.test_dir, "nope.png"))
        self.assertFalse(pdf_gen.has_signature)

    def test_frozen_app_finds_bundled_signature(self):
        """Installed app: the signature is bundled in _internal/data, not next to the exe."""
        install = tempfile.mkdtemp(dir=self.test_dir)
        bundle = os.path.join(install, "_internal")
        os.makedirs(os.path.join(bundle, "data"))
        bundled = os.path.join(bundle, "data", "signature_final.png")
        open(bundled, "wb").close()
        with patch.object(pdf_generator.sys, "frozen", True, create=True),                 patch.object(pdf_generator.sys, "_MEIPASS", bundle, create=True),                 patch.object(pdf_generator.sys, "executable", os.path.join(install, "AnanyaInvoiceAutomation.exe")):
            self.assertEqual(default_signature_path(), bundled)
            # a replacement placed in data/ next to the exe wins
            os.makedirs(os.path.join(install, "data"))
            override = os.path.join(install, "data", "signature_final.png")
            open(override, "wb").close()
            self.assertEqual(default_signature_path(), override)

    def test_regenerating_one_agency_keeps_other_agencies_pdfs(self):
        """Old PDFs are matched by full name: 'BG' must not delete 'KCN_BG' (substring match did)."""
        pdf_gen = PDFGenerator(self.test_dir)
        other = os.path.join(self.test_dir, "Invoice_1800_KCN_BG.pdf")
        own_old = os.path.join(self.test_dir, "Invoice_1700_BG.pdf")
        for f in (other, own_old):
            open(f, "wb").close()
        agency = dict(self.agencies[0], sheet_name="BG", invoice_no=1799)
        pdf_gen.generate_agency_pdf(agency)
        self.assertTrue(os.path.exists(other))
        self.assertFalse(os.path.exists(own_old))
        self.assertTrue(os.path.exists(os.path.join(self.test_dir, "Invoice_1799_BG.pdf")))


if __name__ == "__main__":
    unittest.main()
