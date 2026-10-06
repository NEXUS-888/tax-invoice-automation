import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock

import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from PyQt6.QtWidgets import QApplication, QPushButton, QTableWidget, QDialog, QMessageBox
from PyQt6.QtCore import Qt
from app_gui import InvoiceAutomationApp, EditAgencyDialog
from state_manager import StateManager

app = QApplication.instance()
if not app:
    app = QApplication([])


class TestGUIIntegration(unittest.TestCase):
    @patch("app_gui.InvoiceAutomationApp.load_master_file")
    def setUp(self, mock_load):
        self.gui = InvoiceAutomationApp()
        self.temp_dir = tempfile.TemporaryDirectory()
        self.temp_state = os.path.join(self.temp_dir.name, "test_state.json")
        self.gui.state_path = self.temp_state
        self.gui.state_mgr = StateManager(self.temp_state)

    def tearDown(self):
        self.gui.close()
        self.temp_dir.cleanup()

    def test_dispatch_table_columns(self):
        # Column count must be 8
        self.assertEqual(self.gui.dispatch_table.columnCount(), 8)
        headers = [self.gui.dispatch_table.horizontalHeaderItem(i).text() for i in range(8)]
        self.assertIn("Auto-Send", headers[6])
        self.assertIn("Share App", headers[7])

    def test_agency_banner_buttons(self):
        # Ensure btn_share_current_agency exists and is a QPushButton
        self.assertIsInstance(self.gui.btn_share_current_agency, QPushButton)
        self.assertEqual(self.gui.btn_share_current_agency.text(), "📤 Share App")

    def test_populate_dispatch_table_creates_share_buttons(self):
        sample_agencies = [
            {
                'sheet_name': 'AGENCY_ONE',
                'agency_name': 'AGENCY ONE GAS',
                'invoice_no': '101',
                'grand_total': 15000.0,
                'month_desc': 'SEPTEMBER 2026'
            }
        ]
        sample_pdfs = [
            {
                'sheet_name': 'AGENCY_ONE',
                'pdf_path': 'output/AGENCY_ONE.pdf'
            }
        ]
        self.gui.contacts_mgr.get_phones = MagicMock(return_value=['919876543210'])
        self.gui.populate_dispatch_table(sample_agencies, sample_pdfs)

        self.assertEqual(self.gui.dispatch_table.rowCount(), 1)
        send_btn = self.gui.dispatch_table.cellWidget(0, 6)
        share_btn = self.gui.dispatch_table.cellWidget(0, 7)

        self.assertIsInstance(send_btn, QPushButton)
        self.assertIsInstance(share_btn, QPushButton)
        self.assertEqual(send_btn.text(), "📱 Auto-Send")
        self.assertEqual(share_btn.text(), "📤 Share App")

    def test_get_or_generate_agency_pdf_on_demand(self):
        # Even when batch list is completely empty:
        self.gui.generated_agency_list = []
        self.gui.generated_pdf_list = []

        self.gui.parsed_agencies = [{
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA GAS AGENCY',
            'pan_no': 'ADGPH7087R',
            'vendor_code': '380893',
            'gst_no': '29ADGPH7087R2ZJ',
            'gstin': 'GSTIN123',
            'address': ['Line 1', 'Line 2', 'City'],
            'vehicles': [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        }]
        self.gui.agency_vehicles_data['MAHALASA'] = [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]

        agency_data, pdf_path = self.gui.get_or_generate_agency_pdf('MAHALASA')
        self.assertIsNotNone(agency_data)
        self.assertIsNotNone(pdf_path)
        self.assertTrue(os.path.exists(pdf_path))
        self.assertEqual(agency_data['sheet_name'], 'MAHALASA')
        self.assertEqual(agency_data['total'], 5500.0)

        # Clean up generated test PDF
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

    @patch("app_gui.InvoiceAutomationApp.on_share_button_clicked")
    def test_share_current_agency_whatsapp_on_demand(self, mock_share_clicked):
        self.gui.current_selected_sheet = 'MAHALASA'
        self.gui.parsed_agencies = [{
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA GAS AGENCY',
            'pan_no': 'ADGPH7087R',
            'vendor_code': '380893',
            'gst_no': '29ADGPH7087R2ZJ',
            'gstin': 'GSTIN123',
            'address': ['Line 1', 'Line 2', 'City'],
            'vehicles': [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        }]
        self.gui.agency_vehicles_data['MAHALASA'] = [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]

        # Without running batch generation:
        self.gui.generated_agency_list = []
        self.gui.generated_pdf_list = []

        # Calling share_current_agency_whatsapp should automatically generate that agency's PDF and call on_share_button_clicked
        self.gui.share_current_agency_whatsapp()

        self.assertTrue(mock_share_clicked.called)
        call_item = mock_share_clicked.call_args[0][0]
        self.assertEqual(call_item['sheet_name'], 'MAHALASA')
        self.assertTrue(os.path.exists(call_item['pdf_path']))

        # Clean up
        if os.path.exists(call_item['pdf_path']):
            os.remove(call_item['pdf_path'])

    def test_manual_invoice_number_editing_in_tab1(self):
        self.gui.current_selected_sheet = 'MAHALASA'
        self.gui.parsed_agencies = [{
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA GAS AGENCY',
            'pan_no': 'ADGPH7087R',
            'vendor_code': '380893',
            'gst_no': '29ADGPH7087R2ZJ',
            'gstin': 'GSTIN123',
            'address': ['Line 1', 'Line 2', 'City'],
            'vehicles': [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        }]
        self.gui.agency_vehicles_data['MAHALASA'] = [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        self.gui.agency_inv_map['MAHALASA'] = 1721

        # User edits the spinbox in Tab 1 banner
        self.gui.agency_inv_spin.setValue(9999)
        self.assertEqual(self.gui.agency_inv_map['MAHALASA'], 9999)

        # Generating on-demand PDF must use the manual invoice number
        agency_data, pdf_path = self.gui.get_or_generate_agency_pdf('MAHALASA')
        self.assertEqual(agency_data['invoice_no'], 9999)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

    def test_manual_invoice_number_editing_in_tab3(self):
        self.gui.current_selected_sheet = 'MAHALASA'
        self.gui.parsed_agencies = [{
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA GAS AGENCY',
            'pan_no': 'ADGPH7087R',
            'vendor_code': '380893',
            'gst_no': '29ADGPH7087R2ZJ',
            'gstin': 'GSTIN123',
            'address': ['Line 1', 'Line 2', 'City'],
            'vehicles': [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        }]
        self.gui.agency_vehicles_data['MAHALASA'] = [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        self.gui.agency_inv_map['MAHALASA'] = 1721

        sample_agencies = [{
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA GAS AGENCY',
            'invoice_no': 1721,
            'grand_total': 6490.0,
            'month_desc': 'SEPTEMBER 2026'
        }]
        self.gui.populate_dispatch_table(sample_agencies, [])

        # User edits Cell (0, 1) in dispatch table
        inv_item = self.gui.dispatch_table.item(0, 1)
        self.assertIsNotNone(inv_item)
        inv_item.setText("8888")

        self.assertEqual(self.gui.agency_inv_map['MAHALASA'], 8888)
        self.assertEqual(self.gui.agency_inv_spin.value(), 8888)

    def test_agency_banner_edit_button(self):
        self.assertIsInstance(self.gui.btn_edit_agency, QPushButton)
        self.assertEqual(self.gui.btn_edit_agency.text(), "✏️ Edit Agency Info")
        self.assertIsInstance(self.gui.btn_reset_defaults, QPushButton)
        self.assertEqual(self.gui.btn_reset_defaults.text(), "🔄 Reset Defaults")

    @patch("PyQt6.QtWidgets.QMessageBox.information")
    @patch.object(EditAgencyDialog, "exec", return_value=QDialog.DialogCode.Accepted)
    @patch.object(EditAgencyDialog, "get_data")
    def test_edit_agency_dialog_updates_metadata_and_state(self, mock_get_data, mock_exec, mock_info):
        self.gui.current_selected_sheet = 'MAHALASA'
        self.gui.parsed_agencies = [{
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA GAS AGENCY',
            'pan_no': 'ADGPH7087R',
            'vendor_code': '380893',
            'gst_no': '29ADGPH7087R2ZJ',
            'gstin': 'GSTIN123',
            'address': ['Line 1', 'Line 2', 'City'],
            'vehicles': [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        }]
        self.gui.agency_vehicles_data['MAHALASA'] = [{'vehicle_no': 'KA 19 D 5917', 'loads': 10.0, 'rate': 550.0}]
        self.gui.agency_meta_map['MAHALASA'] = {
            'agency_name': 'MAHALASA GAS AGENCY',
            'pan_no': 'ADGPH7087R',
            'vendor_code': '380893',
            'gst_no': '29ADGPH7087R2ZJ',
            'gstin': 'GSTIN123',
            'address': ['Line 1', 'Line 2', 'City']
        }
        self.gui.populate_agency_list()

        mock_get_data.return_value = {
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA SUPER GAS',
            'pan_no': 'NEWPAN1234F',
            'vendor_code': '999888',
            'gst_no': '29NEWPAN1234F1Z0',
            'gstin': 'NEWGSTIN',
            'address': ['New Line 1', 'New Line 2', 'New City'],
            'phone': '919999888877'
        }

        self.gui.open_edit_agency_dialog()

        # Check in-memory maps updated
        self.assertEqual(self.gui.agency_meta_map['MAHALASA']['agency_name'], 'MAHALASA SUPER GAS')
        self.assertEqual(self.gui.agency_meta_map['MAHALASA']['pan_no'], 'NEWPAN1234F')
        self.assertEqual(self.gui.parsed_agencies[0]['agency_name'], 'MAHALASA SUPER GAS')
        self.assertIn('MAHALASA SUPER GAS', self.gui.lbl_agency_title.text())
        self.assertEqual(self.gui.contacts_mgr.get_phones('MAHALASA'), ['919999888877'])

        # Check saved state on disk
        saved = self.gui.state_mgr.load_state()
        self.assertIsNotNone(saved)
        self.assertIn('agency_meta_overrides', saved)
        self.assertEqual(saved['agency_meta_overrides']['MAHALASA']['agency_name'], 'MAHALASA SUPER GAS')

    def test_edit_agency_dialog_handles_integer_vendor_code_and_none_values(self):
        meta = {
            'sheet_name': 'TEST_AGENCY',
            'agency_name': None,
            'pan_no': 'ADGPH7087R',
            'vendor_code': 380893,  # Integer as parsed from Excel
            'gst_no': None,
            'gstin': None,
            'address': ['Line 1']
        }
        dialog = EditAgencyDialog(meta, None, self.gui)
        self.assertEqual(dialog.vendor_edit.text(), "380893")
        self.assertEqual(dialog.agency_name_edit.text(), "")
        self.assertEqual(dialog.phone_edit.text(), "")
        data = dialog.get_data()
        self.assertEqual(data['vendor_code'], "380893")

    def test_rate_per_load_editing_in_table(self):
        self.gui.current_selected_sheet = 'MAHALASA'
        self.gui.agency_vehicles_data['MAHALASA'] = [
            {'sl': 1, 'vehicle_no': 'KA 19 D 5917', 'rate': 550.0, 'loads': 10.0, 'total_amount': 5500.0}
        ]
        self.gui.display_agency_vehicles('MAHALASA')

        # Check initial rate item in column 2
        item_rate = self.gui.agency_vehicle_table.item(0, 2)
        self.assertIsNotNone(item_rate)
        self.assertEqual(item_rate.text(), "550.00")

        # Edit rate to 650.00
        item_rate.setText("650.00")

        self.assertEqual(self.gui.agency_vehicles_data['MAHALASA'][0]['rate'], 650.0)
        self.assertEqual(self.gui.agency_vehicles_data['MAHALASA'][0]['total_amount'], 6500.0)
        self.assertIn("6500.00", self.gui.lbl_agency_subtotal.text())

        # Verify auto-saved state
        saved = self.gui.state_mgr.load_state()
        self.assertEqual(saved['agency_loads']['MAHALASA'][0]['rate'], 650.0)

    @patch("app_gui.ExcelManager")
    def test_persistent_state_retains_loads_and_metadata_on_reload(self, mock_excel_mgr_cls):
        # Setup mock excel parser returning default baseline
        mock_mgr_instance = MagicMock()
        mock_excel_mgr_cls.return_value = mock_mgr_instance
        mock_mgr_instance.parse_all_agencies.return_value = [{
            'sheet_name': 'MAHALASA',
            'agency_name': 'MAHALASA GAS AGENCY',
            'pan_no': 'DEFAULT_PAN',
            'vendor_code': '111',
            'gst_no': 'DEFAULT_GST',
            'gstin': 'DEFAULT_GSTIN',
            'address': ['Default 1', 'Default 2', 'Default 3'],
            'invoice_no': 100,
            'vehicles': [{'vehicle_no': 'KA 19 D 5917', 'rate': 550.0, 'loads': 10.0, 'total_amount': 5500.0}],
            'grand_total': 6490.0
        }]

        # Pre-populate saved state simulating previous session edits
        saved_payload = {
            'target_month': 'OCTOBER 2026',
            'invoice_date': '2026-10-05',
            'starting_inv_no': 3000,
            'agency_inv_map': {'MAHALASA': 7777},
            'agency_meta_overrides': {
                'MAHALASA': {
                    'agency_name': 'MAHALASA EDITED CORP',
                    'pan_no': 'EDITED_PAN',
                    'vendor_code': '999',
                    'gst_no': 'EDITED_GST',
                    'gstin': 'EDITED_GSTIN',
                    'address': ['Ed 1', 'Ed 2', 'Ed 3']
                }
            },
            'agency_loads': {
                'MAHALASA': [{'sl': 1, 'vehicle_no': 'KA 19 D 5917', 'rate': 650.0, 'loads': 30.0, 'total_amount': 19500.0}]
            },
            'custom_agencies': [{
                'sheet_name': 'CUSTOM_AGENCY',
                'agency_name': 'CUSTOM AGENCY PVT LTD',
                'pan_no': 'CUST_PAN',
                'vendor_code': '222',
                'gst_no': 'CUST_GST',
                'gstin': 'CUST_GSTIN',
                'address': ['C1', 'C2', 'C3'],
                'date': '05-10-2026',
                'invoice_no': 3001,
                'month_desc': 'TEST MONTH',
                'vehicles': [{'sl': 1, 'vehicle_no': 'KA 01 AA 1234', 'rate': 700.0, 'loads': 15.0, 'total_amount': 10500.0}],
                'total': 0, 'sgst': 0, 'cgst': 0, 'grand_total': 0, 'in_words': ''
            }]
        }
        self.gui.state_mgr.save_state(saved_payload)

        # Call load_master_file (unmocked method on self.gui)
        InvoiceAutomationApp.load_master_file(self.gui)

        # Verify that state overlay was applied, NOT reverted to Excel defaults:
        self.assertEqual(self.gui.month_edit.text(), "OCTOBER 2026")
        self.assertEqual(self.gui.starting_inv_spin.value(), 3000)
        self.assertEqual(self.gui.agency_meta_map['MAHALASA']['agency_name'], 'MAHALASA EDITED CORP')
        self.assertEqual(self.gui.agency_meta_map['MAHALASA']['pan_no'], 'EDITED_PAN')
        self.assertEqual(self.gui.agency_vehicles_data['MAHALASA'][0]['loads'], 30.0)
        self.assertEqual(self.gui.agency_vehicles_data['MAHALASA'][0]['rate'], 650.0)
        self.assertEqual(self.gui.agency_inv_map['MAHALASA'], 7777)

        # Verify custom agency was restored
        self.assertTrue(any(a['sheet_name'] == 'CUSTOM_AGENCY' for a in self.gui.parsed_agencies))
        self.assertIn('CUSTOM_AGENCY', self.gui.agency_vehicles_data)

    @patch("PyQt6.QtWidgets.QMessageBox.information")
    @patch("PyQt6.QtWidgets.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes)
    @patch.object(InvoiceAutomationApp, "load_master_file")
    def test_reset_defaults_clears_state(self, mock_load, mock_question, mock_info):
        # Save dummy state
        self.gui.state_mgr.save_state({'target_month': 'TEST_MONTH'})
        self.assertTrue(os.path.exists(self.temp_state))

        self.gui.reset_to_excel_defaults()

        self.assertEqual(self.gui.state_mgr.load_state(), {})
        self.assertEqual(self.gui.custom_agencies_list, [])
        self.assertTrue(mock_load.called)


if __name__ == "__main__":
    unittest.main()
