import sys
# Force openpyxl to use pure-Python numeric types (int, float, Decimal)
# and completely prevent dummy/incomplete numpy in frozen executables
sys.modules['numpy'] = None

import os
import sys
import subprocess
import time
import random
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QTabWidget, QFileDialog, QMessageBox, QTextEdit, QHeaderView,
    QGroupBox, QSpinBox, QDateEdit, QSplitter, QListWidget, QListWidgetItem,
    QDialog, QFormLayout, QDoubleSpinBox, QProgressBar, QMenu, QToolTip, QInputDialog
)
from PyQt6.QtCore import Qt, QDate, QThread, QObject, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QFont, QColor, QIcon, QCursor

from excel_manager import ExcelManager
from pdf_generator import PDFGenerator
from contacts_manager import ContactsManager, normalize_phone
from whatsapp_dispatcher import WhatsAppDispatcher
from whatsapp_automator import WhatsAppAutomator
from state_manager import StateManager


class AddNewAgencyDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Agency Profile")
        self.setFixedSize(480, 520)

        layout = QFormLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        self.sheet_name_edit = QLineEdit()
        self.sheet_name_edit.setPlaceholderText("Short Name e.g. ROYAL BHARAT")
        layout.addRow("Agency Sheet Name*:", self.sheet_name_edit)

        self.agency_name_edit = QLineEdit()
        self.agency_name_edit.setPlaceholderText("Full Name e.g. ROYAL BHARAT GAS AGENCY")
        layout.addRow("Full Agency Name*:", self.agency_name_edit)

        self.pan_edit = QLineEdit("ADGPH7087R")
        layout.addRow("PAN No:", self.pan_edit)

        self.vendor_edit = QLineEdit("380893")
        layout.addRow("Vendor Code:", self.vendor_edit)

        self.gst_no_edit = QLineEdit("29ADGPH7087R2ZJ")
        layout.addRow("GST No:", self.gst_no_edit)

        self.gstin_edit = QLineEdit()
        self.gstin_edit.setPlaceholderText("e.g. GSTIN – 29ABDFM3371Q2ZL")
        layout.addRow("GSTIN:", self.gstin_edit)

        self.addr1_edit = QLineEdit()
        self.addr1_edit.setPlaceholderText("Address Line 1")
        layout.addRow("Address Line 1:", self.addr1_edit)

        self.addr2_edit = QLineEdit()
        self.addr2_edit.setPlaceholderText("Address Line 2")
        layout.addRow("Address Line 2:", self.addr2_edit)

        self.addr3_edit = QLineEdit()
        self.addr3_edit.setPlaceholderText("City / Pincode e.g. BANGALORE-560010")
        layout.addRow("Address Line 3:", self.addr3_edit)

        self.vehicle_edit = QLineEdit()
        self.vehicle_edit.setPlaceholderText("First Vehicle e.g. KA-02-AG-1234")
        layout.addRow("Primary Vehicle No:", self.vehicle_edit)

        self.rate_spin = QDoubleSpinBox()
        self.rate_spin.setRange(0, 100000)
        self.rate_spin.setValue(550.0)
        layout.addRow("Rate per Load (₹):", self.rate_spin)

        self.phone_edit = QLineEdit()
        self.phone_edit.setPlaceholderText("WhatsApp Number(s) e.g. 919876543210, 919876543211")
        layout.addRow("WhatsApp Phone(s):", self.phone_edit)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("Save New Agency")
        add_btn.setObjectName("actionBtn")
        add_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(cancel_btn)

        layout.addRow(btn_layout)

    def get_data(self):
        s_name = self.sheet_name_edit.text().strip().upper()
        a_name = self.agency_name_edit.text().strip().upper() or s_name
        v_no = self.vehicle_edit.text().strip().upper()

        return {
            'sheet_name': s_name,
            'agency_name': a_name,
            'pan_no': self.pan_edit.text().strip().upper(),
            'vendor_code': self.vendor_edit.text().strip(),
            'gst_no': self.gst_no_edit.text().strip().upper(),
            'gstin': self.gstin_edit.text().strip().upper(),
            'address': [
                self.addr1_edit.text().strip(),
                self.addr2_edit.text().strip(),
                self.addr3_edit.text().strip()
            ],
            'vehicle_no': v_no,
            'rate': self.rate_spin.value(),
            'phone': self.phone_edit.text().strip()
        }


class EditAgencyDialog(QDialog):
    def __init__(self, agency_meta, phone_str="", parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Edit Agency Profile: {agency_meta.get('sheet_name', '')}")
        self.setFixedSize(500, 530)

        layout = QFormLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        self.sheet_name = str(agency_meta.get('sheet_name') or '')
        sheet_lbl = QLabel(self.sheet_name)
        sheet_lbl.setStyleSheet("font-weight: bold; font-size: 14px; color: #38bdf8;")
        layout.addRow("Agency Sheet Name:", sheet_lbl)

        self.agency_name_edit = QLineEdit(str(agency_meta.get('agency_name') or ''))
        self.agency_name_edit.setPlaceholderText("Full Agency Name")
        layout.addRow("Full Agency Name*:", self.agency_name_edit)

        self.pan_edit = QLineEdit(str(agency_meta.get('pan_no') or ''))
        layout.addRow("PAN No:", self.pan_edit)

        self.vendor_edit = QLineEdit(str(agency_meta.get('vendor_code') or ''))
        layout.addRow("Vendor Code:", self.vendor_edit)

        self.gst_no_edit = QLineEdit(str(agency_meta.get('gst_no') or ''))
        layout.addRow("GST No:", self.gst_no_edit)

        self.gstin_edit = QLineEdit(str(agency_meta.get('gstin') or ''))
        layout.addRow("GSTIN:", self.gstin_edit)

        addr = agency_meta.get('address', ['', '', ''])
        if not isinstance(addr, list):
            addr = [str(addr)]
        if len(addr) < 3:
            addr = addr + [''] * (3 - len(addr))

        self.addr1_edit = QLineEdit(str(addr[0] or ''))
        layout.addRow("Address Line 1:", self.addr1_edit)

        self.addr2_edit = QLineEdit(str(addr[1] or ''))
        layout.addRow("Address Line 2:", self.addr2_edit)

        self.addr3_edit = QLineEdit(str(addr[2] or ''))
        layout.addRow("Address Line 3:", self.addr3_edit)

        self.phone_edit = QLineEdit(str(phone_str or ''))
        self.phone_edit.setPlaceholderText("WhatsApp Number(s) e.g. 919876543210, 919876543211")
        layout.addRow("WhatsApp Phone(s):", self.phone_edit)

        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save Changes")
        save_btn.setObjectName("actionBtn")
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)

        layout.addRow(btn_layout)

    def get_data(self):
        a_name = self.agency_name_edit.text().strip().upper() or self.sheet_name
        return {
            'sheet_name': self.sheet_name,
            'agency_name': a_name,
            'pan_no': self.pan_edit.text().strip().upper(),
            'vendor_code': self.vendor_edit.text().strip(),
            'gst_no': self.gst_no_edit.text().strip().upper(),
            'gstin': self.gstin_edit.text().strip().upper(),
            'address': [
                self.addr1_edit.text().strip(),
                self.addr2_edit.text().strip(),
                self.addr3_edit.text().strip()
            ],
            'phone': self.phone_edit.text().strip()
        }


class AddVehicleDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Vehicle")
        self.setFixedSize(360, 220)

        layout = QFormLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        self.vno_edit = QLineEdit()
        self.vno_edit.setPlaceholderText("e.g. KA-02-AG-9999")
        layout.addRow("Vehicle No:", self.vno_edit)

        self.rate_spin = QDoubleSpinBox()
        self.rate_spin.setRange(0, 100000)
        self.rate_spin.setValue(550.0)
        self.rate_spin.setDecimals(2)
        layout.addRow("Rate per Load (₹):", self.rate_spin)

        self.load_spin = QDoubleSpinBox()
        self.load_spin.setRange(0, 10000)
        self.load_spin.setValue(10.0)
        self.load_spin.setDecimals(0)
        layout.addRow("Initial Loads:", self.load_spin)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("Add Vehicle")
        add_btn.setObjectName("addBtn")
        add_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(cancel_btn)

        layout.addRow(btn_layout)

    def get_data(self):
        return {
            'vehicle_no': self.vno_edit.text().strip().upper(),
            'rate': self.rate_spin.value(),
            'loads': self.load_spin.value()
        }


class WhatsAppWorker(QObject):
    """
    Runs ALL Playwright operations on a single dedicated Python thread.
    Uses a queue instead of QThread signals to avoid asyncio event loop conflicts.
    """

    log_message = pyqtSignal(str)
    connected = pyqtSignal(bool, str)
    single_result = pyqtSignal(object, bool, str)
    share_result = pyqtSignal(object, bool, str)
    bulk_result = pyqtSignal(int, bool, str)
    bulk_finished = pyqtSignal(int, int)

    def __init__(self, session_dir):
        super().__init__()
        self.session_dir = session_dir
        self.automator = None
        self._thread = None
        self._queue = None

    def start_in_thread(self):
        """Launch the Playwright session on a dedicated background thread."""
        import threading
        import queue
        self._queue = queue.Queue()
        self._thread = threading.Thread(target=self._worker_loop, daemon=True)
        self._thread.start()

    def _worker_loop(self):
        """Runs on the dedicated thread. Processes all tasks sequentially."""
        import queue

        # Step 1: Launch Playwright session (Interactive visible browser window)
        try:
            self.automator = WhatsAppAutomator(self.session_dir, log_callback=self.log_message.emit)
            ok = self.automator.launch_session(headless=False)
            self.connected.emit(ok, "WhatsApp Web session started." if ok else "Failed to start WhatsApp Web.")
        except Exception as e:
            self.connected.emit(False, f"Failed: {e}")
            return






        # Step 2: Process queued tasks forever
        while True:
            try:
                task = self._queue.get(timeout=1)
            except queue.Empty:
                continue

            if task is None:
                # Shutdown signal
                break

            action = task.get('action')
            try:
                if action == 'send_single':
                    self._do_send_single(task['item'])
                elif action == 'share_preview':
                    self._do_share_preview(task['item'])
                elif action == 'send_bulk':
                    self._do_send_bulk(task['jobs'])
                elif action == 'close':
                    break
            except Exception as e:
                self.log_message.emit(f"Worker error: {e}")

        # Cleanup
        try:
            if self.automator:
                self.automator.close()
        except Exception:
            pass

    def _do_send_single(self, item):
        if not self.automator:
            self.single_result.emit(item, False, "WhatsApp is not connected.")
            return
        try:
            ok, msg = self.automator.send_invoice_pdf(
                phone_input=item['phone'],
                agency_name=item['agency_name'],
                invoice_no=item['inv_no'],
                month_desc=item['month_desc'],
                pdf_path=item['pdf_path'],
            )
            self.single_result.emit(item, ok, msg)
        except Exception as e:
            self.single_result.emit(item, False, f"Error: {e}")

    def _do_share_preview(self, item):
        if not self.automator:
            self.share_result.emit(item, False, "WhatsApp is not connected.")
            return
        try:
            ok, msg = self.automator.prepare_invoice_share(
                phone_input=item['phone'],
                agency_name=item['agency_name'],
                invoice_no=item['inv_no'],
                month_desc=item['month_desc'],
                pdf_path=item['pdf_path'],
                auto_send=False,
            )
            self.share_result.emit(item, ok, msg)
        except Exception as e:
            self.share_result.emit(item, False, f"Error: {e}")

    def queue_share_preview(self, item):
        """Queue an interactive share: attach PDF + caption, leave Send to the user (called from main thread)."""
        if self._queue:
            self._queue.put({'action': 'share_preview', 'item': item})

    def _do_send_bulk(self, jobs):
        success_count = 0
        total = len(jobs)
        if not self.automator:
            self.bulk_finished.emit(0, total)
            return

        # Ensure WhatsApp Web is logged in (prompts to scan QR if needed)
        if not self.automator.is_logged_in():
            self.log_message.emit("📱 Checking WhatsApp Web session...")
            if not self.automator.wait_for_login(timeout=60):
                self.log_message.emit("❌ WhatsApp Web is not logged in! Please scan the QR code in the browser window and try again.")
                self.bulk_finished.emit(0, total)
                return

        self.log_message.emit("🟢 WhatsApp Web logged in and ready! Starting dispatch...")

        try:
            for idx, job in enumerate(jobs):
                if idx > 0:
                    delay = round(random.uniform(5.0, 12.0), 1)
                    self.log_message.emit(f"⏳ Waiting {delay}s anti-ban pause before sending next agency invoice...")
                    time.sleep(delay)

                item = job['item']
                if not item['phone']:
                    self.bulk_result.emit(job['row_idx'], False, "No Phone Number")
                    continue
                try:
                    ok, msg = self.automator.send_invoice_pdf(
                        phone_input=item['phone'],
                        agency_name=item['agency_name'],
                        invoice_no=item['inv_no'],
                        month_desc=item['month_desc'],
                        pdf_path=item['pdf_path'],
                    )
                    if ok:
                        success_count += 1
                    self.bulk_result.emit(job['row_idx'], ok, msg)
                except Exception as e:
                    self.bulk_result.emit(job['row_idx'], False, f"Error: {e}")
        except Exception as e:
            self.log_message.emit(f"Bulk send error: {e}")
        finally:
            self.bulk_finished.emit(success_count, total)

    def queue_single_send(self, item):
        """Queue a single agency send (called from main thread)."""
        if self._queue:
            self._queue.put({'action': 'send_single', 'item': item})

    def queue_bulk_send(self, jobs):
        """Queue a bulk send (called from main thread)."""
        if self._queue:
            self._queue.put({'action': 'send_bulk', 'jobs': jobs})

    def close(self):
        """Signal the worker thread to shut down."""
        if self._queue:
            self._queue.put(None)
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5)


class InvoiceAutomationApp(QMainWindow):
    request_single_send = pyqtSignal(object)
    request_bulk_send = pyqtSignal(object)
    request_close_whatsapp = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ananya Enterprises - Invoice Automation & WhatsApp Dispatcher")
        self.resize(1220, 850)
        self.setMinimumSize(1000, 700)

        # Base paths
        if getattr(sys, 'frozen', False):
            self.base_dir = os.path.abspath(os.path.dirname(sys.executable))
            bundle_dir = getattr(sys, '_MEIPASS', self.base_dir)
        else:
            self.base_dir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
            bundle_dir = self.base_dir

        self.master_excel_path = os.path.join(self.base_dir, "Ananya Bill.xlsm")
        if not os.path.exists(self.master_excel_path):
            bundled_excel = os.path.join(bundle_dir, "Ananya Bill.xlsm")
            if os.path.exists(bundled_excel):
                try:
                    import shutil
                    shutil.copy2(bundled_excel, self.master_excel_path)
                except Exception:
                    self.master_excel_path = bundled_excel

        self.contacts_path = os.path.join(self.base_dir, "data", "contacts.json")
        self.state_path = os.path.join(self.base_dir, "data", "app_state.json")
        self.wa_session_dir = os.path.join(self.base_dir, "data", "wa_session")
        self.output_dir = os.path.join(self.base_dir, "output")
        os.makedirs(os.path.join(self.base_dir, "data"), exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)

        # Managers
        self.contacts_mgr = ContactsManager(self.contacts_path)
        self.state_mgr = StateManager(self.state_path)
        self.wa_dispatcher = WhatsAppDispatcher()
        self.wa_automator = None
        self.wa_thread = None
        self.wa_worker = None
        self.whatsapp_connected = False
        self.parsed_agencies = []
        self.custom_agencies_list = []
        self.agency_vehicles_data = {}
        self.agency_meta_map = {}
        self.agency_inv_map = {}
        self.current_selected_sheet = None
        self.generated_agency_list = []
        self.generated_pdf_list = []
        self.current_pdf_dir = self.output_dir

        # UI Setup
        self.init_ui()
        self.load_master_file()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(12)

        # Stylesheet
        self.setStyleSheet("""
            QMainWindow, QDialog { background-color: #121418; color: #f8fafc; }
            
            /* Container Cards */
            QGroupBox {
                font-weight: bold;
                font-size: 13px;
                border: 1px solid #2a2f3a;
                border-radius: 8px;
                margin-top: 8px;
                padding: 14px;
                background-color: #1a1e24;
                color: #38bdf8;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 6px;
                color: #38bdf8;
            }

            /* Typography */
            QLabel {
                color: #f1f5f9;
                font-size: 13px;
                font-weight: 500;
            }

            /* Input Controls */
            QLineEdit, QSpinBox, QDoubleSpinBox, QDateEdit {
                background-color: #0f1216;
                color: #f8fafc;
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 6px 10px;
                min-height: 30px;
                font-size: 13px;
            }
            QLineEdit:focus, QSpinBox:focus, QDateEdit:focus {
                border: 1px solid #38bdf8;
                background-color: #161b22;
            }

            /* Buttons */
            QPushButton {
                background-color: #2563eb;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 6px 14px;
                font-weight: bold;
                font-size: 13px;
                min-height: 30px;
            }
            QPushButton:hover {
                background-color: #1d4ed8;
            }
            QPushButton#actionBtn {
                background-color: #16a34a;
                font-size: 14px;
                padding: 8px 18px;
                min-height: 36px;
            }
            QPushButton#actionBtn:hover {
                background-color: #15803d;
            }
            QPushButton#waConnectBtn {
                background-color: #0d9488;
            }
            QPushButton#waConnectBtn:hover {
                background-color: #0f766e;
            }
            QPushButton#bulkSendBtn {
                background-color: #2563eb;
                font-size: 14px;
                min-height: 36px;
            }
            QPushButton#singleSendBtn {
                background-color: #16a34a;
                font-size: 12px;
                padding: 4px 8px;
                min-height: 26px;
            }
            QPushButton#singleSendBtn:hover {
                background-color: #15803d;
            }
            QPushButton#shareBtn {
                background-color: #059669;
                font-size: 12px;
                padding: 4px 8px;
                min-height: 26px;
            }
            QPushButton#shareBtn:hover {
                background-color: #10b981;
            }
            QPushButton#viewPdfBtn {
                background-color: #0284c7;
                font-size: 12px;
                padding: 4px 8px;
                min-height: 26px;
            }
            QPushButton#viewPdfBtn:hover {
                background-color: #0369a1;
            }
            QPushButton#folderBtn {
                background-color: #475569;
                font-size: 13px;
                min-height: 36px;
            }
            QPushButton#folderBtn:hover {
                background-color: #334155;
            }
            QPushButton#addBtn {
                background-color: #ea580c;
            }
            QPushButton#addBtn:hover {
                background-color: #c2410c;
            }
            QPushButton#newAgencyBtn {
                background-color: #8b5cf6;
            }
            QPushButton#newAgencyBtn:hover {
                background-color: #7c3aed;
            }
            QPushButton#editAgencyBtn {
                background-color: #0284c7;
                font-size: 12px;
                padding: 4px 10px;
                min-height: 26px;
            }
            QPushButton#editAgencyBtn:hover {
                background-color: #0369a1;
            }
            QPushButton#resetBtn {
                background-color: #334155;
                font-size: 12px;
                padding: 6px 12px;
                min-height: 30px;
            }
            QPushButton#resetBtn:hover {
                background-color: #dc2626;
            }

            /* Tab Navigation Bar */
            QTabWidget::pane {
                border: 1px solid #2a2f3a;
                background-color: #1a1e24;
                border-radius: 8px;
            }
            QTabBar::tab {
                background-color: #16191f;
                color: #94a3b8;
                border: 1px solid #2a2f3a;
                padding: 10px 22px;
                font-size: 13px;
                font-weight: bold;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                margin-right: 3px;
            }
            QTabBar::tab:selected {
                background-color: #2563eb;
                color: #ffffff;
                border-bottom: 2px solid #38bdf8;
            }
            QTabBar::tab:hover:!selected {
                background-color: #222730;
                color: #e2e8f0;
            }

            /* Agency Left List Widget */
            QListWidget {
                background-color: #0f1216;
                border: 1px solid #2a2f3a;
                border-radius: 6px;
                color: #f8fafc;
                font-size: 13px;
                padding: 4px;
            }
            QListWidget::item {
                padding: 10px 12px;
                border-bottom: 1px solid #1e232d;
                border-radius: 4px;
                color: #e2e8f0;
            }
            QListWidget::item:selected {
                background-color: #2563eb;
                color: #ffffff;
                font-weight: bold;
            }
            QListWidget::item:hover:!selected {
                background-color: #1e232d;
            }

            /* Tables */
            QTableWidget {
                background-color: #0f1216;
                border: 1px solid #2a2f3a;
                border-radius: 6px;
                gridline-color: #2a2f3a;
                color: #f8fafc;
                font-size: 13px;
            }
            QTableWidget::item {
                color: #f8fafc;
                padding: 6px;
            }
            QHeaderView::section {
                background-color: #16191f;
                color: #38bdf8;
                font-weight: bold;
                padding: 8px;
                border: 1px solid #2a2f3a;
            }
            QTableWidget::item:selected {
                background-color: #1e293b;
                color: #38bdf8;
            }

            /* Progress Bar */
            QProgressBar {
                border: 1px solid #334155;
                border-radius: 6px;
                text-align: center;
                background-color: #0f1216;
                color: #f8fafc;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #2563eb;
                border-radius: 5px;
            }

            /* Splitter */
            QSplitter::handle {
                background-color: #2a2f3a;
                width: 3px;
            }
        """)

        # --- TOP HEADER & CONTROLS ---
        top_group = QGroupBox("Invoice Batch Setup & WhatsApp Connection")
        top_grid = QGridLayout(top_group)
        top_grid.setHorizontalSpacing(15)
        top_grid.setVerticalSpacing(10)

        # Row 0: Master File
        top_grid.addWidget(QLabel("Master Excel File:"), 0, 0)
        self.file_path_edit = QLineEdit(self.master_excel_path)
        self.file_path_edit.setReadOnly(True)
        top_grid.addWidget(self.file_path_edit, 0, 1, 1, 3)

        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_master_file)
        top_grid.addWidget(browse_btn, 0, 4)

        # WhatsApp Connect Button on Top
        self.btn_wa_connect = QPushButton("📱 Connect WhatsApp Account")
        self.btn_wa_connect.setObjectName("waConnectBtn")
        self.btn_wa_connect.clicked.connect(self.connect_whatsapp_account)
        top_grid.addWidget(self.btn_wa_connect, 0, 5)

        self.btn_reset_defaults = QPushButton("🔄 Reset Defaults")
        self.btn_reset_defaults.setObjectName("resetBtn")
        self.btn_reset_defaults.setToolTip("Clear saved overrides and reload default data from master Excel")
        self.btn_reset_defaults.clicked.connect(self.reset_to_excel_defaults)
        top_grid.addWidget(self.btn_reset_defaults, 0, 6)

        # Row 1: Target Month, Date, Starting Inv
        top_grid.addWidget(QLabel("Target Month:"), 1, 0)
        self.month_edit = QLineEdit("AUGUST 2026")
        self.month_edit.setPlaceholderText("e.g. AUGUST 2026")
        self.month_edit.textChanged.connect(lambda: self.save_app_state())
        top_grid.addWidget(self.month_edit, 1, 1)

        top_grid.addWidget(QLabel("Invoice Date:"), 1, 2)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.dateChanged.connect(lambda: self.save_app_state())
        top_grid.addWidget(self.date_edit, 1, 3)

        top_grid.addWidget(QLabel("Starting Inv #:"), 1, 4)
        self.starting_inv_spin = QSpinBox()
        self.starting_inv_spin.setRange(1, 99999)
        self.starting_inv_spin.setValue(1721)
        self.starting_inv_spin.valueChanged.connect(self.on_starting_inv_changed)
        top_grid.addWidget(self.starting_inv_spin, 1, 5)

        main_layout.addWidget(top_group)

        # --- TABS WORKSPACE ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Tab 1: Agency & Load Manager
        self.tab_loads = QWidget()
        self.setup_agency_wise_loads_tab()
        self.tabs.addTab(self.tab_loads, "🏢 Agency & Load Manager")

        # Tab 2: WhatsApp Contacts Directory
        self.tab_contacts = QWidget()
        self.setup_contacts_tab()
        self.tabs.addTab(self.tab_contacts, "📱 WhatsApp Contacts")

        # Tab 3: PDF Generation & Automated Dispatch
        self.tab_dispatch = QWidget()
        self.setup_dispatch_tab()
        self.tabs.addTab(self.tab_dispatch, "📄 Generate & Auto-Send WhatsApp Invoices")

    def setup_agency_wise_loads_tab(self):
        layout = QVBoxLayout(self.tab_loads)
        layout.setContentsMargins(10, 10, 10, 10)

        # Main Splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- LEFT PANE: AGENCY LIST ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_header_layout = QHBoxLayout()
        left_header = QLabel("Select Agency:")
        left_header.setStyleSheet("font-weight: bold; font-size: 13px; color: #38bdf8;")
        left_header_layout.addWidget(left_header)
        left_header_layout.addStretch()

        self.btn_new_agency = QPushButton("➕ Add New Agency")
        self.btn_new_agency.setObjectName("newAgencyBtn")
        self.btn_new_agency.clicked.connect(self.open_add_new_agency_dialog)
        left_header_layout.addWidget(self.btn_new_agency)

        left_layout.addLayout(left_header_layout)

        self.agency_search_edit = QLineEdit()
        self.agency_search_edit.setPlaceholderText("Search agencies...")
        self.agency_search_edit.textChanged.connect(self.filter_agency_list)
        left_layout.addWidget(self.agency_search_edit)

        self.agency_list_widget = QListWidget()
        self.agency_list_widget.currentItemChanged.connect(self.on_agency_selected)
        left_layout.addWidget(self.agency_list_widget)

        splitter.addWidget(left_widget)

        # --- RIGHT PANE: SELECTED AGENCY VEHICLE DETAILS ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 0, 0, 0)

        # Agency Banner Card
        self.agency_banner = QGroupBox("Agency Details")
        banner_layout = QHBoxLayout(self.agency_banner)
        
        self.lbl_agency_title = QLabel("Select an agency from the left panel")
        self.lbl_agency_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #38bdf8;")
        banner_layout.addWidget(self.lbl_agency_title)
        
        banner_layout.addStretch()

        lbl_inv = QLabel("Invoice No:")
        lbl_inv.setStyleSheet("font-size: 13px; font-weight: bold; color: #f8fafc;")
        banner_layout.addWidget(lbl_inv)

        self.agency_inv_spin = QSpinBox()
        self.agency_inv_spin.setRange(1, 9999999)
        self.agency_inv_spin.setFixedWidth(110)
        self.agency_inv_spin.setToolTip("Manually enter or override invoice number for this agency")
        self.agency_inv_spin.valueChanged.connect(self.on_agency_inv_no_changed)
        banner_layout.addWidget(self.agency_inv_spin)

        banner_layout.addSpacing(10)

        self.btn_view_current_pdf = QPushButton("👁️ View PDF Invoice")
        self.btn_view_current_pdf.setObjectName("viewPdfBtn")
        self.btn_view_current_pdf.clicked.connect(self.view_current_agency_pdf)
        banner_layout.addWidget(self.btn_view_current_pdf)

        self.btn_send_current_agency = QPushButton("📱 Auto-Send")
        self.btn_send_current_agency.setObjectName("singleSendBtn")
        self.btn_send_current_agency.setToolTip("Automated background send via WhatsApp Web")
        self.btn_send_current_agency.clicked.connect(self.send_current_agency_whatsapp)
        banner_layout.addWidget(self.btn_send_current_agency)

        self.btn_share_current_agency = QPushButton("📤 Share App")
        self.btn_share_current_agency.setObjectName("shareBtn")
        self.btn_share_current_agency.setToolTip("Open in native WhatsApp App (Copies PDF to clipboard for Ctrl+V)")
        self.btn_share_current_agency.clicked.connect(self.share_current_agency_whatsapp)
        banner_layout.addWidget(self.btn_share_current_agency)

        self.btn_edit_agency = QPushButton("✏️ Edit Agency Info")
        self.btn_edit_agency.setObjectName("editAgencyBtn")
        self.btn_edit_agency.setToolTip("Edit agency name, PAN, GST, address, vendor code")
        self.btn_edit_agency.clicked.connect(self.open_edit_agency_dialog)
        banner_layout.addWidget(self.btn_edit_agency)

        self.btn_add_vehicle = QPushButton("➕ Add Vehicle")
        self.btn_add_vehicle.setObjectName("addBtn")
        self.btn_add_vehicle.clicked.connect(self.open_add_vehicle_dialog)
        banner_layout.addWidget(self.btn_add_vehicle)

        right_layout.addWidget(self.agency_banner)

        # Vehicle Table for Selected Agency
        self.agency_vehicle_table = QTableWidget()
        self.agency_vehicle_table.setColumnCount(5)
        self.agency_vehicle_table.setHorizontalHeaderLabels([
            "Sl No", "Vehicle Reg Number", "Rate per Load (₹) (Edit)", "Current Loads", "Next Month Loads (Edit Here)"
        ])
        self.agency_vehicle_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.agency_vehicle_table.cellChanged.connect(self.on_load_cell_changed)
        right_layout.addWidget(self.agency_vehicle_table)

        # Summary Bar at Bottom of Agency View
        summary_box = QHBoxLayout()
        self.lbl_agency_subtotal = QLabel("Total Subtotal: ₹ 0.00")
        self.lbl_agency_subtotal.setStyleSheet("font-size: 13px; font-weight: bold; color: #cbd5e1;")
        self.lbl_agency_grand = QLabel("Grand Total (incl 18% GST): ₹ 0.00")
        self.lbl_agency_grand.setStyleSheet("font-size: 14px; font-weight: bold; color: #4ade80;")
        
        summary_box.addWidget(self.lbl_agency_subtotal)
        summary_box.addSpacing(25)
        summary_box.addWidget(self.lbl_agency_grand)
        summary_box.addStretch()

        right_layout.addLayout(summary_box)

        splitter.addWidget(right_widget)

        # Set ratio: 30% left list, 70% right detail view
        splitter.setSizes([340, 800])
        layout.addWidget(splitter)

    def setup_contacts_tab(self):
        layout = QVBoxLayout(self.tab_contacts)
        layout.setContentsMargins(10, 10, 10, 10)

        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("Enter agency WhatsApp numbers (comma separated for multiple e.g. 919876543210, 919876543211):"))
        
        save_contacts_btn = QPushButton("Save WhatsApp Numbers")
        save_contacts_btn.clicked.connect(self.save_contacts_from_table)
        header_layout.addWidget(save_contacts_btn)

        layout.addLayout(header_layout)

        # Contacts Table
        self.contacts_table = QTableWidget()
        self.contacts_table.setColumnCount(3)
        self.contacts_table.setHorizontalHeaderLabels(["Agency Sheet Name", "Full Agency Name", "WhatsApp Numbers (Multiple Allowed)"])
        self.contacts_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.contacts_table.cellChanged.connect(self.on_contacts_cell_changed)
        layout.addWidget(self.contacts_table)

    def on_contacts_cell_changed(self, row, col):
        if col != 2:
            return
        s_item = self.contacts_table.item(row, 0)
        a_item = self.contacts_table.item(row, 1)
        p_item = self.contacts_table.item(row, 2)
        if s_item and p_item:
            s_name = s_item.text()
            a_name = a_item.text() if a_item else ""
            phone_input = p_item.text().strip()
            self.contacts_mgr.update_phone(s_name, phone_input, a_name)

    def setup_dispatch_tab(self):
        layout = QVBoxLayout(self.tab_dispatch)
        layout.setContentsMargins(10, 10, 10, 10)

        # Action Buttons
        btn_layout = QHBoxLayout()

        self.btn_gen_all = QPushButton("1. Generate All PDFs & Excel")
        self.btn_gen_all.setObjectName("actionBtn")
        self.btn_gen_all.clicked.connect(self.run_batch_generation)
        btn_layout.addWidget(self.btn_gen_all)

        self.btn_open_folder = QPushButton("📂 Open Invoices Folder")
        self.btn_open_folder.setObjectName("folderBtn")
        self.btn_open_folder.clicked.connect(self.open_output_folder)
        btn_layout.addWidget(self.btn_open_folder)

        self.btn_send_all_wa = QPushButton("2. 🚀 Automated Send ALL Invoices via WhatsApp")
        self.btn_send_all_wa.setObjectName("bulkSendBtn")
        self.btn_send_all_wa.clicked.connect(self.run_bulk_whatsapp_send)
        btn_layout.addWidget(self.btn_send_all_wa)

        layout.addLayout(btn_layout)

        # Progress Bar
        self.send_progress_bar = QProgressBar()
        self.send_progress_bar.setValue(0)
        self.send_progress_bar.setFixedHeight(22)
        layout.addWidget(self.send_progress_bar)

        # Log & Dispatch Table Splitter
        splitter = QSplitter(Qt.Orientation.Vertical)

        # Summary Table
        self.dispatch_table = QTableWidget()
        self.dispatch_table.setColumnCount(8)
        self.dispatch_table.setHorizontalHeaderLabels([
            "Agency", "Inv No", "Grand Total (₹)", "WhatsApp Numbers", "PDF Invoice", "WhatsApp Status", "Auto-Send", "Share App"
        ])
        self.dispatch_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.dispatch_table.cellChanged.connect(self.on_dispatch_cell_changed)
        splitter.addWidget(self.dispatch_table)

        # Log Console
        log_group = QGroupBox("Execution Log Console")
        log_layout = QVBoxLayout(log_group)
        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)
        self.log_console.setStyleSheet("font-family: Consolas; font-size: 12px; background-color: #0f1216; color: #4ade80;")
        log_layout.addWidget(self.log_console)

        splitter.addWidget(log_group)
        layout.addWidget(splitter)

    def connect_whatsapp_account(self):
        """Launches Playwright WhatsApp Web session window for login/QR code scan."""
        if self.wa_worker and self.wa_worker._thread and self.wa_worker._thread.is_alive():
            self.log("WhatsApp session is already starting or active.")
            return

        self.log("Opening WhatsApp Web connection window...")
        self.wa_worker = WhatsAppWorker(self.wa_session_dir)
        self.wa_worker.log_message.connect(self.log)
        self.wa_worker.connected.connect(self.on_whatsapp_connected)
        self.wa_worker.single_result.connect(self.on_single_send_result)
        self.wa_worker.share_result.connect(self.on_share_result)
        self.wa_worker.bulk_result.connect(self.on_bulk_send_result)
        self.wa_worker.bulk_finished.connect(self.on_bulk_send_finished)
        self.wa_worker.start_in_thread()
        self.btn_wa_connect.setEnabled(False)
        self.btn_wa_connect.setText("⏳ Starting WhatsApp...")

    @pyqtSlot(bool, str)
    def on_whatsapp_connected(self, success, message):
        self.whatsapp_connected = success
        self.btn_wa_connect.setEnabled(True)
        if success:
            self.wa_automator = self.wa_worker.automator
            self.btn_wa_connect.setText("🟢 WhatsApp Connected")
            self.btn_wa_connect.setStyleSheet("background-color: #059669; font-weight: bold;")
            QMessageBox.information(
                self, "WhatsApp Connection",
                "WhatsApp Web browser launched!\n\nIf you see a QR code on screen, scan it with your phone's WhatsApp.\n\nOnce logged in, your session remains saved permanently on your computer!"
            )
        else:
            self.btn_wa_connect.setText("📱 Connect WhatsApp Account")
            QMessageBox.critical(self, "Connection Error", message)

    def browse_master_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Master Excel File", "", "Excel Files (*.xlsm *.xlsx)")
        if file_path:
            self.master_excel_path = file_path
            self.file_path_edit.setText(file_path)
            self.load_master_file()

    def load_master_file(self):
        if not os.path.exists(self.master_excel_path):
            self.log(f"Master file not found at: {self.master_excel_path}")
            return

        try:
            # Generated files belong to the previously loaded workbook.
            self.current_selected_sheet = None
            self.generated_agency_list = []
            self.generated_pdf_list = []
            self.dispatch_table.setRowCount(0)
            self.log(f"Loading master workbook: {self.master_excel_path}...")
            mgr = ExcelManager(self.master_excel_path)
            self.parsed_agencies = mgr.parse_all_agencies()
            self.contacts_mgr.sync_agencies(self.parsed_agencies)

            # Build in-memory vehicles map for agency-wise view
            self.agency_vehicles_data = {}
            self.agency_meta_map = {}
            max_inv = 0
            for a in self.parsed_agencies:
                s_name = a['sheet_name']
                self.agency_vehicles_data[s_name] = [dict(v) for v in a['vehicles']]
                self.agency_meta_map[s_name] = {
                    'agency_name': a['agency_name'],
                    'pan_no': a['pan_no'],
                    'vendor_code': a['vendor_code'],
                    'gst_no': a['gst_no'],
                    'gstin': a['gstin'],
                    'address': a['address']
                }
                if a['invoice_no'] and str(a['invoice_no']).isdigit():
                    max_inv = max(max_inv, int(a['invoice_no']))

            if max_inv > 0:
                self.starting_inv_spin.blockSignals(True)
                self.starting_inv_spin.setValue(max_inv + 1)
                self.starting_inv_spin.blockSignals(False)

            # Pre-populate default agency_inv_map
            base_inv = self.starting_inv_spin.value()
            self.agency_inv_map = {}
            for idx, a in enumerate(self.parsed_agencies):
                s_name = a['sheet_name']
                self.agency_inv_map[s_name] = base_inv + idx

            # Restore saved state if exists
            saved_state = self.state_mgr.load_state()
            if saved_state:
                # Top controls
                if saved_state.get('target_month') and hasattr(self, 'month_edit'):
                    self.month_edit.blockSignals(True)
                    self.month_edit.setText(saved_state['target_month'])
                    self.month_edit.blockSignals(False)
                if saved_state.get('invoice_date') and hasattr(self, 'date_edit'):
                    q_date = QDate.fromString(saved_state['invoice_date'], "yyyy-MM-dd")
                    if q_date.isValid():
                        self.date_edit.blockSignals(True)
                        self.date_edit.setDate(q_date)
                        self.date_edit.blockSignals(False)
                if saved_state.get('starting_inv_no') and hasattr(self, 'starting_inv_spin'):
                    self.starting_inv_spin.blockSignals(True)
                    self.starting_inv_spin.setValue(int(saved_state['starting_inv_no']))
                    self.starting_inv_spin.blockSignals(False)

                # Custom agencies
                self.custom_agencies_list = saved_state.get('custom_agencies', [])
                for ca in self.custom_agencies_list:
                    s_name = ca['sheet_name']
                    if not any(a['sheet_name'] == s_name for a in self.parsed_agencies):
                        self.parsed_agencies.append(dict(ca))
                    self.agency_meta_map[s_name] = {
                        'agency_name': ca.get('agency_name', s_name),
                        'pan_no': ca.get('pan_no', ''),
                        'vendor_code': ca.get('vendor_code', ''),
                        'gst_no': ca.get('gst_no', ''),
                        'gstin': ca.get('gstin', ''),
                        'address': ca.get('address', ['', '', ''])
                    }
                    if s_name not in self.agency_vehicles_data:
                        self.agency_vehicles_data[s_name] = [dict(v) for v in ca.get('vehicles', [])]

                # Agency metadata overrides
                meta_overrides = saved_state.get('agency_meta_overrides', {})
                for s_name, meta in meta_overrides.items():
                    if s_name in self.agency_meta_map:
                        self.agency_meta_map[s_name].update(meta)
                    else:
                        self.agency_meta_map[s_name] = dict(meta)
                    for a in self.parsed_agencies:
                        if a['sheet_name'] == s_name:
                            a.update(meta)
                            break

                # Agency loads and rates overrides
                loads_data = saved_state.get('agency_loads', {})
                for s_name, v_list in loads_data.items():
                    self.agency_vehicles_data[s_name] = [dict(v) for v in v_list]
                    for a in self.parsed_agencies:
                        if a['sheet_name'] == s_name:
                            a['vehicles'] = [dict(v) for v in v_list]
                            break

                # Agency invoice numbers overrides
                saved_inv_map = saved_state.get('agency_inv_map', {})
                for s_name, inv in saved_inv_map.items():
                    self.agency_inv_map[s_name] = int(inv)

            self.contacts_mgr.sync_agencies(self.parsed_agencies)
            self.populate_agency_list()
            self.populate_contacts_table()

            # Pre-populate dispatch table with loaded agencies so user can share immediately
            target_month = self.month_edit.text().strip().upper() if hasattr(self, 'month_edit') else ""
            if not target_month:
                target_month = datetime.now().strftime("%B %Y").upper()
            month_desc = f"LPG CYLINDER LOADING AND UNLOADING CHARGES FOR THE MONTH {target_month}"
            initial_agencies = []
            curr_base = self.starting_inv_spin.value()
            for idx, a in enumerate(self.parsed_agencies):
                s_name = a['sheet_name']
                v_list = self.agency_vehicles_data.get(s_name, a.get('vehicles', []))
                subtotal = sum(float(v.get('rate', 0.0)) * float(v.get('loads', 0.0)) for v in v_list)
                g_total = round(subtotal * 1.18, 2)
                initial_agencies.append({
                    'sheet_name': s_name,
                    'agency_name': self.agency_meta_map.get(s_name, {}).get('agency_name', a['agency_name']),
                    'invoice_no': self.agency_inv_map.get(s_name, curr_base + idx),
                    'grand_total': g_total,
                    'month_desc': month_desc
                })
            self.populate_dispatch_table(initial_agencies, [])
            self.log(f"Loaded {len(self.parsed_agencies)} agency invoice sheets successfully.")

        except Exception as e:
            self.log(f"Error loading master file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to parse Excel file:\n{str(e)}")

    def open_add_new_agency_dialog(self):
        """Opens dialog to register a brand new agency profile."""
        dialog = AddNewAgencyDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            s_name = data['sheet_name']
            if not s_name:
                QMessageBox.warning(self, "Warning", "Agency sheet name cannot be empty.")
                return

            if s_name in self.agency_vehicles_data:
                QMessageBox.warning(self, "Warning", f"Agency '{s_name}' already exists.")
                return

            # Register in app memory
            v_list = []
            if data['vehicle_no']:
                v_list.append({
                    'sl': 1,
                    'hsn': "",
                    'vehicle_no': data['vehicle_no'],
                    'rate': data['rate'],
                    'loads': 10.0,
                    'total_amount': data['rate'] * 10.0
                })

            self.agency_vehicles_data[s_name] = v_list
            self.agency_meta_map[s_name] = {
                'agency_name': data['agency_name'],
                'pan_no': data['pan_no'],
                'vendor_code': data['vendor_code'],
                'gst_no': data['gst_no'],
                'gstin': data['gstin'],
                'address': data['address']
            }

            new_record = {
                'sheet_name': s_name,
                'agency_name': data['agency_name'],
                'pan_no': data['pan_no'],
                'vendor_code': data['vendor_code'],
                'gst_no': data['gst_no'],
                'gstin': data['gstin'],
                'address': data['address'],
                'date': self.date_edit.date().toString("dd-MM-yyyy"),
                'invoice_no': 0,
                'month_desc': f"LPG CYLINDER LOADING AND UNLOADING CHARGES FOR THE MONTH {self.month_edit.text().strip().upper()}",
                'vehicles': v_list,
                'total': 0, 'sgst': 0, 'cgst': 0, 'grand_total': 0, 'in_words': ''
            }
            self.parsed_agencies.append(new_record)
            self.custom_agencies_list.append(new_record)

            # Register phone in contacts
            if data['phone']:
                self.contacts_mgr.update_phone(s_name, data['phone'], data['agency_name'])

            self.populate_agency_list()
            self.populate_contacts_table()

            # Select newly added agency in list
            for i in range(self.agency_list_widget.count()):
                item = self.agency_list_widget.item(i)
                if item.data(Qt.ItemDataRole.UserRole) == s_name:
                    self.agency_list_widget.setCurrentItem(item)
                    break

            self.save_app_state()
            self.log(f"Successfully created new agency profile: {data['agency_name']} ({s_name}).")
            QMessageBox.information(self, "Success", f"New Agency '{data['agency_name']}' created successfully!")

    def populate_agency_list(self):
        self.agency_list_widget.clear()
        for a in self.parsed_agencies:
            s_name = a['sheet_name']
            v_count = len(self.agency_vehicles_data.get(s_name, []))
            item = QListWidgetItem(f"{a['agency_name']} ({v_count} vehicles)")
            item.setData(Qt.ItemDataRole.UserRole, s_name)
            self.agency_list_widget.addItem(item)

        if self.agency_list_widget.count() > 0 and not self.current_selected_sheet:
            self.agency_list_widget.setCurrentRow(0)

    def filter_agency_list(self, query):
        query = query.lower().strip()
        for i in range(self.agency_list_widget.count()):
            item = self.agency_list_widget.item(i)
            item.setHidden(query not in item.text().lower())

    def on_agency_selected(self, current_item, previous_item):
        if not current_item:
            return
        sheet_name = current_item.data(Qt.ItemDataRole.UserRole)
        self.current_selected_sheet = sheet_name
        
        agency_info = next((a for a in self.parsed_agencies if a['sheet_name'] == sheet_name), None)
        if agency_info:
            self.lbl_agency_title.setText(f"{agency_info['agency_name']}  (Sheet: {sheet_name})")

        if hasattr(self, 'agency_inv_spin'):
            inv_no = self.agency_inv_map.get(sheet_name)
            if inv_no is None:
                base_inv = self.starting_inv_spin.value() if hasattr(self, 'starting_inv_spin') else 101
                idx = 0
                for i, a in enumerate(self.parsed_agencies):
                    if a['sheet_name'] == sheet_name:
                        idx = i
                        break
                inv_no = base_inv + idx
                self.agency_inv_map[sheet_name] = inv_no
            self.agency_inv_spin.blockSignals(True)
            self.agency_inv_spin.setValue(int(inv_no))
            self.agency_inv_spin.blockSignals(False)

        self.display_agency_vehicles(sheet_name)

    def display_agency_vehicles(self, sheet_name):
        self.agency_vehicle_table.blockSignals(True)
        self.agency_vehicle_table.setRowCount(0)

        vehicles = self.agency_vehicles_data.get(sheet_name, [])
        subtotal = 0.0

        for r_idx, v in enumerate(vehicles):
            self.agency_vehicle_table.insertRow(r_idx)

            # Sl No
            item_sl = QTableWidgetItem(str(v.get('sl', r_idx+1)))
            item_sl.setFlags(item_sl.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_sl.setForeground(QColor("#f8fafc"))
            self.agency_vehicle_table.setItem(r_idx, 0, item_sl)

            # Vehicle No
            item_vno = QTableWidgetItem(v['vehicle_no'])
            item_vno.setFlags(item_vno.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_vno.setForeground(QColor("#f8fafc"))
            self.agency_vehicle_table.setItem(r_idx, 1, item_vno)

            # Rate (EDITABLE CELL WITH HIGH-CONTRAST CYAN TEXT ON DARK BG)
            item_rate = QTableWidgetItem(f"{float(v['rate']):.2f}")
            item_rate.setBackground(QColor("#0f172a"))
            item_rate.setForeground(QColor("#38bdf8"))
            font_rate = QFont()
            font_rate.setBold(True)
            font_rate.setPointSize(10)
            item_rate.setFont(font_rate)
            item_rate.setToolTip("Double click to edit vehicle rate per load")
            self.agency_vehicle_table.setItem(r_idx, 2, item_rate)

            # Current Loads
            item_cur = QTableWidgetItem(f"{float(v['loads']):.0f}")
            item_cur.setFlags(item_cur.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_cur.setForeground(QColor("#94a3b8"))
            self.agency_vehicle_table.setItem(r_idx, 3, item_cur)

            # Next Month Loads (EDITABLE CELL WITH HIGH-CONTRAST CYAN TEXT ON DARK BG)
            item_next = QTableWidgetItem(f"{float(v['loads']):.0f}")
            item_next.setBackground(QColor("#0f172a"))
            item_next.setForeground(QColor("#38bdf8"))
            font = QFont()
            font.setBold(True)
            font.setPointSize(10)
            item_next.setFont(font)
            item_next.setToolTip("Double click to edit next month load count")
            self.agency_vehicle_table.setItem(r_idx, 4, item_next)

            subtotal += float(v['loads']) * float(v['rate'])

        grand_total = subtotal * 1.18
        self.lbl_agency_subtotal.setText(f"Total Subtotal: ₹ {subtotal:.2f}")
        self.lbl_agency_grand.setText(f"Grand Total (incl 18% GST): ₹ {grand_total:.2f}")

        self.agency_vehicle_table.blockSignals(False)

    def on_load_cell_changed(self, row, col):
        if col not in (2, 4) or not self.current_selected_sheet:
            return

        item = self.agency_vehicle_table.item(row, col)
        if not item:
            return

        v_list = self.agency_vehicles_data.get(self.current_selected_sheet, [])
        if row >= len(v_list):
            return

        try:
            val = float(item.text().strip())
        except ValueError:
            val = 0.0

        if col == 2:
            v_list[row]['rate'] = val
        elif col == 4:
            v_list[row]['loads'] = val

        v_list[row]['total_amount'] = v_list[row]['rate'] * v_list[row]['loads']

        # Update parsed_agencies vehicles copy if present
        for a in self.parsed_agencies:
            if a['sheet_name'] == self.current_selected_sheet:
                a['vehicles'] = [dict(v) for v in v_list]
                break

        self.display_agency_vehicles(self.current_selected_sheet)

        # Update grand total in dispatch table row if present
        if hasattr(self, 'dispatch_table'):
            subtotal = sum(float(v.get('rate', 0.0)) * float(v.get('loads', 0.0)) for v in v_list)
            g_total = round(subtotal * 1.18, 2)
            self.dispatch_table.blockSignals(True)
            for r in range(self.dispatch_table.rowCount()):
                s_item = self.dispatch_table.item(r, 0)
                agency_meta = self.agency_meta_map.get(self.current_selected_sheet, {})
                a_name = agency_meta.get('agency_name', self.current_selected_sheet)
                if s_item and (s_item.data(Qt.ItemDataRole.UserRole) == self.current_selected_sheet or s_item.text() in (self.current_selected_sheet, a_name)):
                    tot_item = self.dispatch_table.item(r, 2)
                    if tot_item:
                        tot_item.setText(f"{g_total:.2f}")
                    break
            self.dispatch_table.blockSignals(False)

        self.save_app_state()

    def on_agency_inv_no_changed(self, new_val):
        """Called when user edits the invoice number for the currently selected agency in Tab 1."""
        if not self.current_selected_sheet:
            return
        self.agency_inv_map[self.current_selected_sheet] = new_val

        # Update Dispatch Table Column 1 for this agency
        if hasattr(self, 'dispatch_table'):
            self.dispatch_table.blockSignals(True)
            for r in range(self.dispatch_table.rowCount()):
                s_item = self.dispatch_table.item(r, 0)
                agency_meta = self.agency_meta_map.get(self.current_selected_sheet, {})
                a_name = agency_meta.get('agency_name', self.current_selected_sheet)
                if s_item and s_item.text() in (self.current_selected_sheet, a_name):
                    inv_item = self.dispatch_table.item(r, 1)
                    if inv_item and inv_item.text() != str(new_val):
                        inv_item.setText(str(new_val))
                    break
            self.dispatch_table.blockSignals(False)

        # Update generated_agency_list if present
        for a in self.generated_agency_list:
            if a['sheet_name'] == self.current_selected_sheet:
                a['invoice_no'] = new_val
                break

        self.save_app_state()

    def on_starting_inv_changed(self, start_val):
        """Called when user changes Starting Inv # in top bar."""
        for idx, a in enumerate(self.parsed_agencies):
            s_name = a['sheet_name']
            self.agency_inv_map[s_name] = start_val + idx

        if self.current_selected_sheet and self.current_selected_sheet in self.agency_inv_map:
            if hasattr(self, 'agency_inv_spin'):
                self.agency_inv_spin.blockSignals(True)
                self.agency_inv_spin.setValue(self.agency_inv_map[self.current_selected_sheet])
                self.agency_inv_spin.blockSignals(False)

        if hasattr(self, 'dispatch_table'):
            self.dispatch_table.blockSignals(True)
            for r in range(self.dispatch_table.rowCount()):
                s_item = self.dispatch_table.item(r, 0)
                if s_item:
                    for a in self.parsed_agencies:
                        if a['agency_name'] == s_item.text() or a['sheet_name'] == s_item.text():
                            inv_item = self.dispatch_table.item(r, 1)
                            if inv_item:
                                inv_item.setText(str(self.agency_inv_map[a['sheet_name']]))
                            break
            self.dispatch_table.blockSignals(False)

        self.save_app_state()

    def on_dispatch_cell_changed(self, row, col):
        """Called when user edits a cell in the Dispatch Table (specifically Column 1: Inv No)."""
        if col != 1:
            return
        s_item = self.dispatch_table.item(row, 0)
        inv_item = self.dispatch_table.item(row, 1)
        if not s_item or not inv_item:
            return

        try:
            new_inv = int(inv_item.text().strip())
        except ValueError:
            return

        sheet_name = None
        table_text = s_item.text()
        for a in self.parsed_agencies:
            if a['agency_name'] == table_text or a['sheet_name'] == table_text:
                sheet_name = a['sheet_name']
                break
        if not sheet_name:
            sheet_name = table_text

        self.agency_inv_map[sheet_name] = new_inv

        # If currently selected in Tab 1, sync the spinbox
        if self.current_selected_sheet == sheet_name and hasattr(self, 'agency_inv_spin'):
            self.agency_inv_spin.blockSignals(True)
            self.agency_inv_spin.setValue(new_inv)
            self.agency_inv_spin.blockSignals(False)

        # Update in generated_agency_list if present
        for a in self.generated_agency_list:
            if a['sheet_name'] == sheet_name:
                a['invoice_no'] = new_inv
                break

        self.save_app_state()

    def open_add_vehicle_dialog(self):
        if not self.current_selected_sheet:
            QMessageBox.warning(self, "Warning", "Please select an agency first.")
            return

        dialog = AddVehicleDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            if not data['vehicle_no']:
                QMessageBox.warning(self, "Warning", "Vehicle number cannot be empty.")
                return

            v_list = self.agency_vehicles_data.get(self.current_selected_sheet, [])
            
            existing = next((v for v in v_list if v['vehicle_no'] == data['vehicle_no']), None)
            if existing:
                existing['rate'] = data['rate']
                existing['loads'] = data['loads']
                existing['total_amount'] = data['rate'] * data['loads']
            else:
                v_list.append({
                    'sl': len(v_list) + 1,
                    'hsn': "",
                    'vehicle_no': data['vehicle_no'],
                    'rate': data['rate'],
                    'loads': data['loads'],
                    'total_amount': data['rate'] * data['loads']
                })

            self.display_agency_vehicles(self.current_selected_sheet)
            self.populate_agency_list()
            self.save_app_state()
            self.log(f"Added vehicle {data['vehicle_no']} to agency {self.current_selected_sheet}.")

    def open_edit_agency_dialog(self):
        """Opens dialog to edit agency name, PAN, GST, address, vendor code, and phone numbers."""
        if not self.current_selected_sheet:
            QMessageBox.warning(self, "Warning", "Please select an agency from the left panel first.")
            return

        try:
            sheet_name = self.current_selected_sheet
            meta = dict(self.agency_meta_map.get(sheet_name, {}))
            meta['sheet_name'] = sheet_name
            phone_str = self.contacts_mgr.get_phone_str(sheet_name)

            dialog = EditAgencyDialog(meta, phone_str, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                data = dialog.get_data()

                # 1. Update agency_meta_map
                self.agency_meta_map[sheet_name] = {
                    'agency_name': data['agency_name'],
                    'pan_no': data['pan_no'],
                    'vendor_code': data['vendor_code'],
                    'gst_no': data['gst_no'],
                    'gstin': data['gstin'],
                    'address': data['address']
                }

                # 2. Update parsed_agencies entry
                for a in self.parsed_agencies:
                    if a['sheet_name'] == sheet_name:
                        a['agency_name'] = data['agency_name']
                        a['pan_no'] = data['pan_no']
                        a['vendor_code'] = data['vendor_code']
                        a['gst_no'] = data['gst_no']
                        a['gstin'] = data['gstin']
                        a['address'] = data['address']
                        break

                # 3. Update custom_agencies_list if it is in custom agencies
                for ca in self.custom_agencies_list:
                    if ca['sheet_name'] == sheet_name:
                        ca['agency_name'] = data['agency_name']
                        ca['pan_no'] = data['pan_no']
                        ca['vendor_code'] = data['vendor_code']
                        ca['gst_no'] = data['gst_no']
                        ca['gstin'] = data['gstin']
                        ca['address'] = data['address']
                        break

                # 4. Update generated_agency_list if present
                for ga in self.generated_agency_list:
                    if ga['sheet_name'] == sheet_name:
                        ga['agency_name'] = data['agency_name']
                        ga['pan_no'] = data['pan_no']
                        ga['vendor_code'] = data['vendor_code']
                        ga['gst_no'] = data['gst_no']
                        ga['gstin'] = data['gstin']
                        ga['address'] = data['address']
                        break

                # 5. Update contacts manager phone
                self.contacts_mgr.update_phone(sheet_name, data['phone'], data['agency_name'])

                # 6. Refresh UI components
                self.lbl_agency_title.setText(f"{data['agency_name']}  (Sheet: {sheet_name})")
                
                # Refresh agency list widget items
                for i in range(self.agency_list_widget.count()):
                    item = self.agency_list_widget.item(i)
                    if item.data(Qt.ItemDataRole.UserRole) == sheet_name:
                        v_count = len(self.agency_vehicles_data.get(sheet_name, []))
                        item.setText(f"{data['agency_name']} ({v_count} vehicles)")
                        break

                # Refresh contacts table
                self.populate_contacts_table()

                # Refresh dispatch table row if present
                if hasattr(self, 'dispatch_table'):
                    self.dispatch_table.blockSignals(True)
                    for r in range(self.dispatch_table.rowCount()):
                        s_item = self.dispatch_table.item(r, 0)
                        if s_item and (s_item.data(Qt.ItemDataRole.UserRole) == sheet_name or s_item.text() in (sheet_name, meta.get('agency_name', ''))):
                            s_item.setText(data['agency_name'])
                            s_item.setData(Qt.ItemDataRole.UserRole, sheet_name)
                            phone_item = self.dispatch_table.item(r, 3)
                            if phone_item:
                                phones = self.contacts_mgr.get_phones(sheet_name)
                                phone_display = ", ".join(phones) if phones else "No Phone"
                                if len(phones) > 1:
                                    phone_display += f" ({len(phones)} contacts)"
                                phone_item.setText(phone_display)
                            break
                    self.dispatch_table.blockSignals(False)

                # 7. Persist changes
                self.save_app_state()
                self.log(f"Updated agency info for {data['agency_name']} ({sheet_name}).")
                QMessageBox.information(self, "Agency Info Updated", f"Agency details for '{data['agency_name']}' updated successfully!")
        except Exception as e:
            self.log(f"Error opening agency edit dialog: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to open agency edit dialog:\n{str(e)}")

    def save_app_state(self):
        """Persists current working state to app_state.json so user edits are retained across sessions."""
        try:
            target_m = self.month_edit.text().strip().upper() if hasattr(self, 'month_edit') else ""
            date_s = self.date_edit.date().toString("yyyy-MM-dd") if hasattr(self, 'date_edit') else ""
            start_i = self.starting_inv_spin.value() if hasattr(self, 'starting_inv_spin') else 101

            state = {
                'target_month': target_m,
                'invoice_date': date_s,
                'starting_inv_no': start_i,
                'agency_inv_map': self.agency_inv_map,
                'agency_meta_overrides': self.agency_meta_map,
                'agency_loads': self.agency_vehicles_data,
                'custom_agencies': self.custom_agencies_list
            }
            self.state_mgr.save_state(state)
        except Exception as e:
            self.log(f"Warning: Failed to save application state: {str(e)}")

    def reset_to_excel_defaults(self):
        """Resets all overrides and restores default data from master Excel workbook."""
        reply = QMessageBox.question(
            self, "Reset to Master Defaults",
            "Are you sure you want to reset all data back to the default master Excel workbook?\n\n"
            "This will clear all edited loads, custom rates, added agencies, and manual overrides.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.state_mgr.clear_state()
            self.custom_agencies_list = []
            self.agency_inv_map = {}
            self.load_master_file()
            self.log("Reset all data to default Excel workbook contents.")
            QMessageBox.information(self, "Reset Complete", "All data has been reset to the default master Excel workbook.")

    def get_or_generate_agency_pdf(self, sheet_name, force_regenerate=True):
        """
        Retrieves or on-demand generates the PDF invoice for a single agency.
        Guarantees that the PDF exists and is up-to-date with current GUI values
        without requiring batch generation of all agencies.
        """
        if not sheet_name:
            return None, None

        # 1. Commit active editor in agency_vehicle_table if this agency is currently viewed
        if self.current_selected_sheet == sheet_name and hasattr(self, 'agency_vehicle_table') and self.agency_vehicle_table.rowCount() > 0:
            for r in range(self.agency_vehicle_table.rowCount()):
                rate_item = self.agency_vehicle_table.item(r, 2)
                loads_item = self.agency_vehicle_table.item(r, 4)
                v_list = self.agency_vehicles_data.get(sheet_name, [])
                if r < len(v_list):
                    if rate_item:
                        try:
                            v_list[r]['rate'] = float(rate_item.text().strip())
                        except ValueError:
                            pass
                    if loads_item:
                        try:
                            v_list[r]['loads'] = float(loads_item.text().strip())
                        except ValueError:
                            pass
                    v_list[r]['total_amount'] = v_list[r]['rate'] * v_list[r]['loads']
            self.save_app_state()

        # 2. If not forcing regeneration, check existing cache
        if not force_regenerate:
            existing_pdf = next((p for p in self.generated_pdf_list if p['sheet_name'] == sheet_name), None)
            existing_agency = next((a for a in self.generated_agency_list if a['sheet_name'] == sheet_name), None)
            if existing_pdf and existing_agency and os.path.exists(existing_pdf['pdf_path']):
                return existing_agency, existing_pdf['pdf_path']

        # 3. Determine parameters for this agency
        target_month = self.month_edit.text().strip().upper() if hasattr(self, 'month_edit') else ""
        if not target_month:
            target_month = datetime.now().strftime("%B %Y").upper()

        date_str = self.date_edit.date().toString("dd-MM-yyyy") if hasattr(self, 'date_edit') else datetime.now().strftime("%d-%m-%Y")

        # Determine invoice number
        if sheet_name in self.agency_inv_map:
            inv_no = self.agency_inv_map[sheet_name]
        else:
            existing_agency = next((a for a in self.generated_agency_list if a['sheet_name'] == sheet_name), None)
            if existing_agency and existing_agency.get('invoice_no'):
                inv_no = existing_agency['invoice_no']
            else:
                base_inv = self.starting_inv_spin.value() if hasattr(self, 'starting_inv_spin') else 101
                idx = 0
                for i, a in enumerate(self.parsed_agencies):
                    if a['sheet_name'] == sheet_name:
                        idx = i
                        break
                inv_no = base_inv + idx
            self.agency_inv_map[sheet_name] = inv_no

        # Get vehicles and meta
        vehicles = self.agency_vehicles_data.get(sheet_name)
        meta = self.agency_meta_map.get(sheet_name)

        mgr = ExcelManager(self.master_excel_path)
        agency_data = mgr.build_agency_invoice_data(
            sheet_name=sheet_name,
            target_month_year=target_month,
            target_date_str=date_str,
            invoice_no=inv_no,
            vehicles_list=vehicles,
            agency_meta=meta
        )

        pdf_dir = os.path.join(self.output_dir, f"Invoices_{target_month.replace(' ', '_')}")
        os.makedirs(pdf_dir, exist_ok=True)
        self.current_pdf_dir = pdf_dir

        pdf_gen = PDFGenerator(pdf_dir)
        pdf_path = pdf_gen.generate_agency_pdf(agency_data)

        pdf_info = {
            'sheet_name': sheet_name,
            'agency_name': agency_data['agency_name'],
            'invoice_no': inv_no,
            'pdf_path': pdf_path
        }

        # Cache in memory
        idx_exist = next((i for i, a in enumerate(self.generated_agency_list) if a['sheet_name'] == sheet_name), None)
        if idx_exist is not None:
            self.generated_agency_list[idx_exist] = agency_data
            self.generated_pdf_list[idx_exist] = pdf_info
        else:
            self.generated_agency_list.append(agency_data)
            self.generated_pdf_list.append(pdf_info)

        # If Dispatch table has rows, update this row
        self.dispatch_table.blockSignals(True)
        for r in range(self.dispatch_table.rowCount()):
            s_name_item = self.dispatch_table.item(r, 0)
            if s_name_item and (s_name_item.text() == agency_data['agency_name'] or s_name_item.text() == sheet_name):
                inv_item = QTableWidgetItem(str(inv_no))
                inv_item.setBackground(QColor("#0f172a"))
                inv_item.setForeground(QColor("#38bdf8"))
                font_inv = QFont()
                font_inv.setBold(True)
                inv_item.setFont(font_inv)
                inv_item.setToolTip("Click to manually edit invoice number")
                self.dispatch_table.setItem(r, 1, inv_item)
                self.dispatch_table.setItem(r, 2, QTableWidgetItem(f"{agency_data['grand_total']:.2f}"))
                view_btn = QPushButton(f"👁️ {os.path.basename(pdf_path)}")
                view_btn.setObjectName("viewPdfBtn")
                view_btn.clicked.connect(lambda _, p=pdf_path: self.open_pdf_file(p))
                self.dispatch_table.setCellWidget(r, 4, view_btn)
                break
        self.dispatch_table.blockSignals(False)

        self.log(f"Generated single invoice PDF for {agency_data['agency_name']}: {os.path.basename(pdf_path)}")
        return agency_data, pdf_path

    def view_current_agency_pdf(self):
        """Opens the generated PDF invoice for the currently selected agency, generating on-the-fly if needed."""
        if not self.current_selected_sheet:
            QMessageBox.warning(self, "Warning", "Please select an agency first.")
            return

        agency_data, pdf_path = self.get_or_generate_agency_pdf(self.current_selected_sheet)
        if pdf_path and os.path.exists(pdf_path):
            self.open_pdf_file(pdf_path)
        else:
            QMessageBox.warning(self, "File Not Found", f"Failed to generate PDF invoice for {self.current_selected_sheet}.")

    def view_specific_agency_pdf(self, sheet_name):
        """Generates and opens PDF invoice for a specific agency."""
        agency_data, pdf_path = self.get_or_generate_agency_pdf(sheet_name)
        if pdf_path and os.path.exists(pdf_path):
            self.open_pdf_file(pdf_path)
        else:
            QMessageBox.warning(self, "Error", f"Failed to generate PDF invoice for {sheet_name}.")

    def open_pdf_file(self, pdf_path):
        """Launches the system default PDF viewer for the given PDF path (cross-platform)."""
        if os.path.exists(pdf_path):
            try:
                if hasattr(os, 'startfile'):
                    os.startfile(pdf_path)
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', pdf_path])
                else:
                    subprocess.Popen(['xdg-open', pdf_path])
                self.log(f"Opened PDF file: {pdf_path}")
            except Exception as e:
                self.log(f"Error opening PDF: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to open PDF:\n{str(e)}")
        else:
            QMessageBox.warning(self, "File Not Found", f"PDF file does not exist:\n{pdf_path}")

    def open_output_folder(self):
        """Opens the generated PDF output directory in system file explorer (cross-platform)."""
        target_dir = self.current_pdf_dir if os.path.exists(self.current_pdf_dir) else self.output_dir
        if not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)
        try:
            abs_dir = os.path.abspath(target_dir)
            if hasattr(os, 'startfile'):
                os.startfile(abs_dir)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', abs_dir])
            else:
                subprocess.Popen(['xdg-open', abs_dir])
            self.log(f"Opened output directory: {target_dir}")
        except Exception as e:
            self.log(f"Error opening folder: {str(e)}")

    def send_current_agency_whatsapp(self):
        """Sends WhatsApp invoice for the currently selected agency in Tab 1, generating on-the-fly if needed."""
        if not self.current_selected_sheet:
            QMessageBox.warning(self, "Warning", "Please select an agency first.")
            return

        agency_data, pdf_path = self.get_or_generate_agency_pdf(self.current_selected_sheet)
        if not agency_data or not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Error", f"Could not generate invoice PDF for {self.current_selected_sheet}.")
            return

        row_idx = 0
        for i, a in enumerate(self.generated_agency_list):
            if a['sheet_name'] == self.current_selected_sheet:
                row_idx = i
                break

        item = {
            'row_idx': row_idx,
            'sheet_name': self.current_selected_sheet,
            'agency_name': agency_data['agency_name'],
            'phone': self.contacts_mgr.get_phones(self.current_selected_sheet),
            'inv_no': agency_data['invoice_no'],
            'month_desc': agency_data['month_desc'],
            'pdf_path': pdf_path
        }

        self.send_single_agency_whatsapp(item)

    def share_current_agency_whatsapp(self):
        """Shares invoice for currently selected agency in Tab 1 directly via native WhatsApp App, generating on-the-fly if needed."""
        if not self.current_selected_sheet:
            QMessageBox.warning(self, "Warning", "Please select an agency first.")
            return

        agency_data, pdf_path = self.get_or_generate_agency_pdf(self.current_selected_sheet)
        if not agency_data or not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Error", f"Failed to generate invoice PDF for {self.current_selected_sheet}.")
            return

        row_idx = None
        for i, a in enumerate(self.generated_agency_list):
            if a['sheet_name'] == self.current_selected_sheet:
                row_idx = i
                break

        item = {
            'row_idx': row_idx,
            'sheet_name': self.current_selected_sheet,
            'agency_name': agency_data['agency_name'],
            'phone': self.contacts_mgr.get_phones(self.current_selected_sheet),
            'inv_no': agency_data['invoice_no'],
            'month_desc': agency_data['month_desc'],
            'pdf_path': pdf_path
        }

        self.on_share_button_clicked(item, self.btn_share_current_agency)

    def on_share_button_clicked(self, item, source_widget=None):
        """
        Handles the '📤 Share App' button action.
        - Automatically generates PDF for this agency on-the-fly if needed.
        - Displays a menu allowing the user to choose:
          * Send via WhatsApp Web (in default browser - 100% reliable).
          * Send via WhatsApp Desktop App.
          * Open Contact Selector in Web or Desktop.
          * Highlight PDF in Windows Explorer.
          * Copy PDF to Clipboard.
        """
        sheet_name = item.get('sheet_name', '')
        pdf_path = item.get('pdf_path', '')
        agency_name = item.get('agency_name', '')
        inv_no = item.get('inv_no', '')
        month_desc = item.get('month_desc', '')
        row_idx = item.get('row_idx')

        # Automatically generate ONLY this agency's PDF if missing or not generated!
        if (not pdf_path or not os.path.exists(pdf_path)) and sheet_name:
            agency_data, generated_path = self.get_or_generate_agency_pdf(sheet_name)
            if generated_path and os.path.exists(generated_path):
                pdf_path = generated_path
                item['pdf_path'] = pdf_path
                if agency_data:
                    agency_name = agency_data['agency_name']
                    inv_no = agency_data['invoice_no']
                    month_desc = agency_data['month_desc']
                    item['agency_name'] = agency_name
                    item['inv_no'] = inv_no
                    item['month_desc'] = month_desc

        if not inv_no and sheet_name:
            inv_no = self.agency_inv_map.get(sheet_name, '')
            item['inv_no'] = inv_no

        if not month_desc:
            target_m = self.month_edit.text().strip().upper() if hasattr(self, 'month_edit') else datetime.now().strftime("%B %Y").upper()
            month_desc = f"LPG CYLINDER LOADING AND UNLOADING CHARGES FOR THE MONTH {target_m}"
            item['month_desc'] = month_desc

        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(
                self, "PDF Generation Failed",
                f"Could not generate PDF invoice for {agency_name}."
            )
            return

        s_name = item.get('sheet_name', '')
        if s_name:
            phones = self.contacts_mgr.get_phones(s_name)
        else:
            phones = item.get('phone', [])
            if isinstance(phones, str):
                phones = [p.strip() for p in phones.split(',') if p.strip()]

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #1e293b;
                color: #f8fafc;
                border: 1px solid #475569;
                border-radius: 6px;
                padding: 4px;
            }
            QMenu::item {
                padding: 8px 18px;
                font-size: 13px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #059669;
                color: #ffffff;
            }
            QMenu::separator {
                height: 1px;
                background: #334155;
                margin: 4px 8px;
            }
        """)

        header_act = menu.addAction(f"📤 Share Invoice: {agency_name} (#{inv_no})")
        header_act.setEnabled(False)
        menu.addSeparator()

        if phones:
            for phone_num in phones:
                act_app = menu.addAction(f"📱 Send via WhatsApp App ({phone_num})")
                act_app.setToolTip("Copies the PDF and opens this chat with the message: press Ctrl+V in the chat, then Send")
                act_app.triggered.connect(
                    lambda _, p=phone_num, i_no=inv_no, m_desc=month_desc, p_path=pdf_path, r_i=row_idx:
                    self.execute_direct_share(agency_name, p, i_no, m_desc, p_path, r_i, target="desktop")
                )

                act_web = menu.addAction(f"🌐 Send via WhatsApp Web ({phone_num})")
                act_web.setToolTip("Opens chat via WhatsApp Web with pre-filled message")
                act_web.triggered.connect(
                    lambda _, p=phone_num, i_no=inv_no, m_desc=month_desc, p_path=pdf_path, r_i=row_idx:
                    self.execute_direct_share(agency_name, p, i_no, m_desc, p_path, r_i, target="web")
                )
            menu.addSeparator()
        else:
            act_enter = menu.addAction("📱 Send via WhatsApp App (enter number)...")
            act_enter.setToolTip("Asks for this agency's WhatsApp number, saves it, copies the PDF and opens the chat")
            act_enter.triggered.connect(
                lambda _, i_no=inv_no, m_desc=month_desc, p_path=pdf_path, r_i=row_idx, s_n=s_name:
                self.prompt_number_and_share(agency_name, s_n, i_no, m_desc, p_path, r_i)
            )
            menu.addSeparator()

        act_picker_app = menu.addAction("📱 Open WhatsApp App (Select Contact)...")
        act_picker_app.setToolTip("Copies the PDF and opens WhatsApp with the message: pick the contact, press Ctrl+V in the chat, then Send")
        act_picker_app.triggered.connect(
            lambda _, i_no=inv_no, m_desc=month_desc, p_path=pdf_path, r_i=row_idx:
            self.execute_direct_share(agency_name, None, i_no, m_desc, p_path, r_i, target="desktop")
        )

        act_picker_web = menu.addAction("🌐 Open WhatsApp Web (Select Contact)...")
        act_picker_web.triggered.connect(
            lambda _, i_no=inv_no, m_desc=month_desc, p_path=pdf_path, r_i=row_idx:
            self.execute_direct_share(agency_name, None, i_no, m_desc, p_path, r_i, target="web")
        )

        if phones:
            act_link = menu.addAction("💬 Open via WhatsApp Link (api.whatsapp.com)...")
            act_link.triggered.connect(
                lambda _, p=phones[0], i_no=inv_no, m_desc=month_desc, p_path=pdf_path, r_i=row_idx:
                self.execute_direct_share(agency_name, p, i_no, m_desc, p_path, r_i, target="auto")
            )

        menu.addSeparator()

        act_explorer = menu.addAction("📂 Highlight PDF in Explorer")
        act_explorer.triggered.connect(
            lambda _, p_path=pdf_path: self.highlight_pdf_file(p_path)
        )

        act_copy = menu.addAction("📋 Copy PDF to Clipboard (for Ctrl+V)")
        act_copy.triggered.connect(
            lambda _, p_path=pdf_path: self.copy_pdf_file_to_clipboard(p_path)
        )

        if source_widget:
            menu.exec(source_widget.mapToGlobal(source_widget.rect().bottomLeft()))
        else:
            menu.exec(QCursor.pos())

    def execute_direct_share(self, agency_name, phone, inv_no, month_desc, pdf_path, row_idx=None, target=None):
        """
        Shares one agency invoice:
        - target 'desktop'/'app' (and 'web' without a session) -> WhatsAppDispatcher copies the PDF to the
          clipboard and opens WhatsApp with the message; the user presses Ctrl+V in the chat, then Send.
        - target 'web' with a connected WhatsApp Web session -> Playwright attaches the PDF and fills
          the caption (result arrives in on_share_result); the user presses Send.
        """
        item = {
            'row_idx': row_idx,
            'agency_name': agency_name,
            'phone': phone,
            'inv_no': inv_no,
            'month_desc': month_desc,
            'pdf_path': pdf_path,
            'target': target,
        }

        if target == "web" and phone and self.whatsapp_connected and self.wa_worker:
            item['route'] = 'web_session'
            self._set_share_status(row_idx, "⏳ Attaching in WhatsApp Web...", "#eab308")
            self.log(f"Attaching invoice for {agency_name} in the connected WhatsApp Web session...")
            self.wa_worker.queue_share_preview(item)
            return

        success, msg = self.wa_dispatcher.share_invoice_to_whatsapp(
            agency_name=agency_name,
            phone=phone,
            invoice_no=inv_no,
            month_desc=month_desc,
            pdf_path=pdf_path,
            target=target,
        )
        self.log(msg)

        if success:
            self._set_share_status(row_idx, "📋 PDF copied — press Ctrl+V in the chat", "#38bdf8")
            self._show_share_hint(msg)
        else:
            self._set_share_status(row_idx, "❌ Could not open WhatsApp", "#ef4444")
            self._show_share_hint(msg)

    def prompt_number_and_share(self, agency_name, sheet_name, inv_no, month_desc, pdf_path, row_idx=None):
        """Asks for a missing WhatsApp number, saves it to contacts, then shares straight into that chat."""
        prompt = (f"WhatsApp number for {agency_name}:\n"
                  f"(e.g. 98765 43210 — it will be saved for next time)")
        text = ""
        while True:
            text, ok = QInputDialog.getText(self, "WhatsApp Number", prompt, text=text)
            if not ok or not text.strip():
                return
            clean = normalize_phone(text)
            if clean:
                break
            QMessageBox.warning(self, "Invalid Number",
                                f"'{text}' is not a valid WhatsApp number. Enter a 10-digit mobile number.")

        if sheet_name:
            self.contacts_mgr.update_phone(sheet_name, clean, agency_name)
            self.populate_contacts_table()
        if row_idx is not None and row_idx < self.dispatch_table.rowCount():
            self.dispatch_table.blockSignals(True)
            self.dispatch_table.setItem(row_idx, 3, QTableWidgetItem(clean))
            self.dispatch_table.blockSignals(False)
        self.log(f"Saved WhatsApp number {clean} for {agency_name}.")
        self.execute_direct_share(agency_name, clean, inv_no, month_desc, pdf_path, row_idx, target="desktop")

    def _set_share_status(self, row_idx, text, color):
        if row_idx is None or row_idx >= self.dispatch_table.rowCount():
            return
        status_item = QTableWidgetItem(text)
        status_item.setForeground(QColor(color))
        font = QFont()
        font.setBold(True)
        status_item.setFont(font)
        self.dispatch_table.setItem(row_idx, 5, status_item)

    def _show_share_hint(self, text):
        """Non-intrusive hint next to the cursor (no modal dialog)."""
        QToolTip.showText(QCursor.pos(), text, None, self.rect(), 8000)

    @pyqtSlot(object, bool, str)
    def on_share_result(self, item, attached, msg):
        """Result of an interactive share from the connected WhatsApp Web (Playwright) worker."""
        agency_name = item.get('agency_name', '')
        row_idx = item.get('row_idx')
        via_web = item.get('route') == 'web_session'
        badge = "Web" if via_web else "App"

        if attached:
            self._set_share_status(row_idx, f"✅ PDF Attached [{badge}] — press Send", "#22c55e")
            self.log(f"✅ {agency_name}: {msg}")
            return

        if via_web:
            # The desktop thread already ran the fallback; the Playwright path has not.
            pdf_path = item.get('pdf_path', '')
            if pdf_path and os.path.exists(pdf_path):
                msg = self.wa_dispatcher.run_fallback(os.path.abspath(pdf_path), reason=f"Auto-attach failed ({msg}).")
        self._set_share_status(row_idx, "📋 PDF ready to drop", "#38bdf8")
        self.log(f"⚠️ {agency_name}: {msg}")
        self._show_share_hint(msg)

    def highlight_pdf_file(self, pdf_path):
        """Highlights the specified PDF invoice in system file manager (cross-platform)."""
        if pdf_path and os.path.exists(pdf_path):
            try:
                abs_p = os.path.abspath(pdf_path)
                if sys.platform == 'win32':
                    subprocess.Popen(f'explorer /select,"{abs_p}"')
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', '-R', abs_p])
                else:
                    subprocess.Popen(['xdg-open', os.path.dirname(abs_p)])
                self.log(f"Highlighted PDF in file manager: {pdf_path}")
            except Exception as e:
                self.log(f"Error highlighting PDF: {str(e)}")

    def copy_pdf_file_to_clipboard(self, pdf_path):
        """Copies the PDF invoice to Windows clipboard (CF_HDROP)."""
        if self.wa_dispatcher.copy_pdf_to_clipboard(pdf_path):
            self.log(f"Copied PDF to Windows clipboard: {os.path.basename(pdf_path)}")
            QMessageBox.information(
                self, "Copied",
                "📋 PDF invoice has been copied to your clipboard!\n\nYou can now press Ctrl+V in WhatsApp to attach it."
            )
        else:
            QMessageBox.warning(self, "Error", "Failed to copy PDF to clipboard.")

    def populate_contacts_table(self):
        self.contacts_table.blockSignals(True)
        self.contacts_table.setRowCount(0)
        for idx, a in enumerate(self.parsed_agencies):
            self.contacts_table.insertRow(idx)

            s_name = a['sheet_name']
            phone_str = self.contacts_mgr.get_phone_str(s_name)

            item_sname = QTableWidgetItem(s_name)
            item_sname.setFlags(item_sname.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_sname.setForeground(QColor("#f8fafc"))
            self.contacts_table.setItem(idx, 0, item_sname)

            item_aname = QTableWidgetItem(a['agency_name'])
            item_aname.setFlags(item_aname.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_aname.setForeground(QColor("#f8fafc"))
            self.contacts_table.setItem(idx, 1, item_aname)

            item_phone = QTableWidgetItem(phone_str)
            item_phone.setBackground(QColor("#0f172a"))
            item_phone.setForeground(QColor("#38bdf8"))
            self.contacts_table.setItem(idx, 2, item_phone)
        self.contacts_table.blockSignals(False)

    def save_contacts_from_table(self):
        for r in range(self.contacts_table.rowCount()):
            s_name = self.contacts_table.item(r, 0).text()
            a_name = self.contacts_table.item(r, 1).text()
            phone_input = self.contacts_table.item(r, 2).text().strip()
            self.contacts_mgr.update_phone(s_name, phone_input, a_name)
        self.log("WhatsApp phone numbers saved successfully.")
        QMessageBox.information(self, "Success", "Agency WhatsApp numbers updated.")

    def run_batch_generation(self):
        if not self.parsed_agencies:
            QMessageBox.warning(self, "Warning", "No master agencies loaded.")
            return

        target_month = self.month_edit.text().strip().upper()
        if not target_month:
            QMessageBox.warning(self, "Warning", "Please enter target month (e.g. AUGUST 2026).")
            return

        date_str = self.date_edit.date().toString("dd-MM-yyyy")
        start_inv = self.starting_inv_spin.value()

        # 1. Commit any in-progress cell editor in the vehicle loads table
        if self.current_selected_sheet and self.agency_vehicle_table.rowCount() > 0:
            for r in range(self.agency_vehicle_table.rowCount()):
                rate_item = self.agency_vehicle_table.item(r, 2)
                loads_item = self.agency_vehicle_table.item(r, 4)
                v_list = self.agency_vehicles_data.get(self.current_selected_sheet, [])
                if r < len(v_list):
                    if rate_item:
                        try:
                            v_list[r]['rate'] = float(rate_item.text().strip())
                        except ValueError:
                            pass
                    if loads_item:
                        try:
                            v_list[r]['loads'] = float(loads_item.text().strip())
                        except ValueError:
                            pass
                    v_list[r]['total_amount'] = v_list[r]['rate'] * v_list[r]['loads']
            self.save_app_state()

        # 2. Auto-save all contacts from Tab 2
        if hasattr(self, 'contacts_table') and self.contacts_table.rowCount() > 0:
            for r in range(self.contacts_table.rowCount()):
                s_item = self.contacts_table.item(r, 0)
                a_item = self.contacts_table.item(r, 1)
                p_item = self.contacts_table.item(r, 2)
                if s_item and p_item:
                    s_name = s_item.text()
                    a_name = a_item.text() if a_item else ""
                    phone_input = p_item.text().strip()
                    self.contacts_mgr.update_phone(s_name, phone_input, a_name)

        self.log("Starting Batch Invoice & PDF Generation...")
        self.log(f"Target Month: {target_month} | Date: {date_str} | Starting Inv #: {start_inv}")

        try:
            # 1. Generate updated Excel file
            excel_out_name = f"Ananya_Bill_{target_month.replace(' ', '_')}.xlsx"
            excel_out_path = os.path.join(self.output_dir, excel_out_name)
            
            mgr = ExcelManager(self.master_excel_path)
            self.generated_agency_list = mgr.generate_updated_workbook(
                target_month_year=target_month,
                target_date_str=date_str,
                starting_inv_no=start_inv,
                updated_loads_map=self.agency_vehicles_data,
                output_file_path=excel_out_path,
                agency_meta_map=self.agency_meta_map,
                custom_inv_map=self.agency_inv_map
            )
            for a in self.generated_agency_list:
                self.agency_inv_map[a['sheet_name']] = a['invoice_no']

            saved_excel_path = getattr(mgr, 'last_output_file_path', excel_out_path)
            self.log(f"Saved updated master Excel: {saved_excel_path}")

            # 2. Batch generate PDFs (clean output folder first so no duplicate PDFs exist)
            pdf_dir = os.path.join(self.output_dir, f"Invoices_{target_month.replace(' ', '_')}")
            self.current_pdf_dir = pdf_dir
            pdf_gen = PDFGenerator(pdf_dir)
            self.generated_pdf_list = pdf_gen.batch_generate(self.generated_agency_list, clean_output_dir=True)
            self.log(f"Successfully generated {len(self.generated_pdf_list)} PDF invoices in {pdf_dir}")

            # 3. Populate Dispatch Summary Table
            self.populate_dispatch_table(self.generated_agency_list, self.generated_pdf_list)
            self.tabs.setCurrentIndex(2) # Switch to dispatch tab

            QMessageBox.information(
                self, "Batch Complete",
                f"Successfully generated:\n- Master Excel File\n- {len(self.generated_pdf_list)} PDF Invoices\n\nYou can click '👁️ View PDF' next to any invoice to preview it!"
            )

        except Exception as e:
            self.log(f"Batch generation error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed batch generation:\n{str(e)}")

    def populate_dispatch_table(self, agency_list, pdf_list):
        pdf_map = {p['sheet_name']: p['pdf_path'] for p in pdf_list}
        self.dispatch_table.blockSignals(True)
        self.dispatch_table.setRowCount(0)

        for idx, a in enumerate(agency_list):
            self.dispatch_table.insertRow(idx)
            s_name = a['sheet_name']
            phones = self.contacts_mgr.get_phones(s_name)
            phone_display = ", ".join(phones) if phones else "No Phone"
            if len(phones) > 1:
                phone_display += f" ({len(phones)} contacts)"

            pdf_path = pdf_map.get(s_name, '')

            item_aname = QTableWidgetItem(a['agency_name'])
            item_aname.setData(Qt.ItemDataRole.UserRole, s_name)
            item_aname.setFlags(item_aname.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_aname.setForeground(QColor("#f8fafc"))
            self.dispatch_table.setItem(idx, 0, item_aname)

            inv_no_val = self.agency_inv_map.get(s_name, a['invoice_no'])
            item_inv = QTableWidgetItem(str(inv_no_val))
            item_inv.setBackground(QColor("#0f172a"))
            item_inv.setForeground(QColor("#38bdf8"))
            font_inv = QFont()
            font_inv.setBold(True)
            item_inv.setFont(font_inv)
            item_inv.setToolTip("Click to manually edit invoice number")
            self.dispatch_table.setItem(idx, 1, item_inv)

            item_total = QTableWidgetItem(f"{a['grand_total']:.2f}")
            item_total.setFlags(item_total.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_total.setForeground(QColor("#f8fafc"))
            self.dispatch_table.setItem(idx, 2, item_total)

            item_phone = QTableWidgetItem(phone_display)
            item_phone.setFlags(item_phone.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_phone.setForeground(QColor("#38bdf8") if phones else QColor("#ef4444"))
            self.dispatch_table.setItem(idx, 3, item_phone)

            # Interactive View PDF Button Column
            if pdf_path and os.path.exists(pdf_path):
                btn_label = f"👁️ {os.path.basename(pdf_path)}"
                view_btn = QPushButton(btn_label)
                view_btn.setObjectName("viewPdfBtn")
                view_btn.clicked.connect(lambda _, path=pdf_path: self.open_pdf_file(path))
            else:
                view_btn = QPushButton("👁️ View PDF")
                view_btn.setObjectName("viewPdfBtn")
                view_btn.clicked.connect(lambda _, sn=s_name: self.view_specific_agency_pdf(sn))
            self.dispatch_table.setCellWidget(idx, 4, view_btn)

            item_status = QTableWidgetItem("Ready to Send" if phones else "Missing Phone")
            item_status.setForeground(QColor("#e2e8f0") if phones else QColor("#ef4444"))
            self.dispatch_table.setItem(idx, 5, item_status)

            # Individual Auto-Send Action Button (WhatsApp Web)
            send_btn = QPushButton("📱 Auto-Send")
            send_btn.setObjectName("singleSendBtn")
            send_btn.setToolTip("Automated background send via WhatsApp Web")
            item_dict = {
                'row_idx': idx,
                'sheet_name': s_name,
                'agency_name': a['agency_name'],
                'phone': phones,
                'inv_no': a['invoice_no'],
                'month_desc': a['month_desc'],
                'pdf_path': pdf_path
            }
            send_btn.clicked.connect(lambda _, it=item_dict: self.send_single_agency_whatsapp(it))
            self.dispatch_table.setCellWidget(idx, 6, send_btn)

            # Direct Native WhatsApp Share App Button
            share_btn = QPushButton("📤 Share App")
            share_btn.setObjectName("shareBtn")
            share_btn.setToolTip("Open in native WhatsApp App (Copies PDF to clipboard for Ctrl+V)")
            share_btn.clicked.connect(lambda _, it=item_dict, btn=share_btn: self.on_share_button_clicked(it, btn))
            self.dispatch_table.setCellWidget(idx, 7, share_btn)

        self.dispatch_table.blockSignals(False)

    def send_single_agency_whatsapp(self, item):
        """Sends single agency WhatsApp invoice directly on main GUI thread to avoid greenlet errors."""
        if not self.whatsapp_connected or not self.wa_worker:
            reply = QMessageBox.question(
                self, "WhatsApp Connection Required",
                "WhatsApp Web is not connected yet.\n\nWould you like to connect now to scan your QR code / log in?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.connect_whatsapp_account()
            return

        sheet_name = item.get('sheet_name', '')
        pdf_path = item.get('pdf_path', '')
        if (not pdf_path or not os.path.exists(pdf_path)) and sheet_name:
            agency_data, generated_path = self.get_or_generate_agency_pdf(sheet_name)
            if generated_path and os.path.exists(generated_path):
                pdf_path = generated_path
                item['pdf_path'] = pdf_path
                if agency_data:
                    item['agency_name'] = agency_data['agency_name']
                    item['inv_no'] = agency_data['invoice_no']
                    item['month_desc'] = agency_data['month_desc']

        row_idx = item['row_idx']
        agency_name = item['agency_name']
        phone = self.contacts_mgr.get_phones(sheet_name) if sheet_name else item.get('phone')

        if not phone:
            QMessageBox.warning(self, "Missing Phone", f"No phone number configured for {agency_name}. Please add it in the WhatsApp Contacts tab.")
            return

        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "PDF Not Found", f"Could not generate PDF invoice for {agency_name}.")
            return

        # Mark table row as sending
        status_item = QTableWidgetItem("⏳ Sending...")
        status_item.setForeground(QColor("#eab308"))
        self.dispatch_table.setItem(row_idx, 5, status_item)
        QApplication.processEvents()

        self.log(f"Sending individual WhatsApp invoice to {agency_name}...")
        item = dict(item)
        item['phone'] = phone
        self.wa_worker.queue_single_send(item)

    @pyqtSlot(object, bool, str)
    def on_single_send_result(self, item, success, msg):
        row_idx = item['row_idx']
        color = "#22c55e" if success else "#ef4444"
        res_item = QTableWidgetItem(f"✅ {msg}" if success else f"❌ {msg}")
        res_item.setForeground(QColor(color))
        font = QFont()
        font.setBold(True)
        res_item.setFont(font)
        self.dispatch_table.setItem(row_idx, 5, res_item)
        self.log(f"{'✅ Sent' if success else '❌ Failed'} PDF invoice to {item['agency_name']}: {msg}")

    def run_bulk_whatsapp_send(self):
        """Executes bulk WhatsApp dispatch sequentially on main thread with non-blocking processEvents."""
        if not self.generated_agency_list or not self.generated_pdf_list:
            QMessageBox.warning(self, "Warning", "Please click '1. Generate All PDFs & Excel' first.")
            return

        if not self.whatsapp_connected or not self.wa_worker:
            reply = QMessageBox.question(
                self, "WhatsApp Connection Required",
                "WhatsApp Web is not connected yet.\n\nWould you like to connect now to scan your QR code / log in?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.connect_whatsapp_account()
            return

        pdf_map = {p['sheet_name']: p['pdf_path'] for p in self.generated_pdf_list}
        total = len(self.generated_agency_list)
        self.btn_send_all_wa.setEnabled(False)
        self.btn_send_all_wa.setText("⏳ Sending Invoices...")
        self.send_progress_bar.setValue(0)
        jobs = []
        for idx, a in enumerate(self.generated_agency_list):
            phones = self.contacts_mgr.get_phones(a['sheet_name'])
            jobs.append({
                'row_idx': idx,
                'item': {
                    'sheet_name': a['sheet_name'],
                    'agency_name': a['agency_name'],
                    'phone': phones,
                    'inv_no': a['invoice_no'],
                    'month_desc': a['month_desc'],
                    'pdf_path': pdf_map.get(a['sheet_name'], ''),
                },
            })
            status = QTableWidgetItem("⏳ Queued..." if phones else "❌ No Phone Number")
            status.setForeground(QColor("#eab308") if phones else QColor("#ef4444"))
            self.dispatch_table.setItem(idx, 5, status)

        self.wa_worker.queue_bulk_send(jobs)

    @pyqtSlot(int, bool, str)
    def on_bulk_send_result(self, row_idx, success, msg):
        item = QTableWidgetItem(f"✅ {msg}" if success else f"❌ {msg}")
        item.setForeground(QColor("#22c55e") if success else QColor("#ef4444"))
        self.dispatch_table.setItem(row_idx, 5, item)
        self.send_progress_bar.setValue(int(((row_idx + 1) / max(1, len(self.generated_agency_list))) * 100))

    @pyqtSlot(int, int)
    def on_bulk_send_finished(self, success_count, total):
        self.send_progress_bar.setValue(100)
        self.btn_send_all_wa.setEnabled(True)
        self.btn_send_all_wa.setText("2. 🚀 Automated Send ALL Invoices via WhatsApp")
        self.log(f"Bulk WhatsApp Dispatch finished: {success_count}/{total} sent successfully.")
        QMessageBox.information(
            self, "Dispatch Complete",
            f"WhatsApp Bulk Dispatch Complete!\n\nSuccessfully sent {success_count} out of {total} PDF invoices."
        )

    def on_whatsapp_thread_finished(self):
        self.whatsapp_connected = False
        self.wa_automator = None

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_console.append(f"[{timestamp}] {message}")

    def closeEvent(self, event):
        self.save_app_state()
        if self.wa_worker:
            try:
                self.wa_worker.close()
            except Exception:
                pass
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = InvoiceAutomationApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
