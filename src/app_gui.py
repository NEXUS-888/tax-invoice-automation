import os
import sys
import subprocess
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QTabWidget, QFileDialog, QMessageBox, QTextEdit, QHeaderView,
    QGroupBox, QSpinBox, QDateEdit, QSplitter, QListWidget, QListWidgetItem,
    QDialog, QFormLayout, QDoubleSpinBox, QProgressBar
)
from PyQt6.QtCore import Qt, QDate, QThread, QObject, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QFont, QColor, QIcon

from excel_manager import ExcelManager
from pdf_generator import PDFGenerator
from contacts_manager import ContactsManager
from whatsapp_dispatcher import WhatsAppDispatcher
from whatsapp_automator import WhatsAppAutomator


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

        # Step 1: Launch Playwright session
        try:
            self.automator = WhatsAppAutomator(self.session_dir, log_callback=self.log_message.emit)
            ok = self.automator.launch_session(headless=True)
            self.connected.emit(ok, "WhatsApp Web session started (Background Mode)." if ok else "Failed to start WhatsApp Web.")
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
            for job in jobs:
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
        self.base_dir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
        self.master_excel_path = os.path.join(self.base_dir, "Ananya Bill.xlsm")
        self.contacts_path = os.path.join(self.base_dir, "data", "contacts.json")
        self.wa_session_dir = os.path.join(self.base_dir, "data", "wa_session")
        self.output_dir = os.path.join(self.base_dir, "output")

        # Managers
        self.contacts_mgr = ContactsManager(self.contacts_path)
        self.wa_dispatcher = WhatsAppDispatcher()
        self.wa_automator = None
        self.wa_thread = None
        self.wa_worker = None
        self.whatsapp_connected = False
        self.parsed_agencies = []
        self.agency_vehicles_data = {}
        self.agency_meta_map = {}
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
            QMainWindow { background-color: #121418; }
            
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

        # Row 1: Target Month, Date, Starting Inv
        top_grid.addWidget(QLabel("Target Month:"), 1, 0)
        self.month_edit = QLineEdit("AUGUST 2026")
        self.month_edit.setPlaceholderText("e.g. AUGUST 2026")
        top_grid.addWidget(self.month_edit, 1, 1)

        top_grid.addWidget(QLabel("Invoice Date:"), 1, 2)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        top_grid.addWidget(self.date_edit, 1, 3)

        top_grid.addWidget(QLabel("Starting Inv #:"), 1, 4)
        self.starting_inv_spin = QSpinBox()
        self.starting_inv_spin.setRange(1, 99999)
        self.starting_inv_spin.setValue(1721)
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

        self.btn_view_current_pdf = QPushButton("👁️ View PDF Invoice")
        self.btn_view_current_pdf.setObjectName("viewPdfBtn")
        self.btn_view_current_pdf.clicked.connect(self.view_current_agency_pdf)
        banner_layout.addWidget(self.btn_view_current_pdf)

        self.btn_send_current_agency = QPushButton("📱 Send WhatsApp")
        self.btn_send_current_agency.setObjectName("singleSendBtn")
        self.btn_send_current_agency.clicked.connect(self.send_current_agency_whatsapp)
        banner_layout.addWidget(self.btn_send_current_agency)

        self.btn_add_vehicle = QPushButton("➕ Add Vehicle")
        self.btn_add_vehicle.setObjectName("addBtn")
        self.btn_add_vehicle.clicked.connect(self.open_add_vehicle_dialog)
        banner_layout.addWidget(self.btn_add_vehicle)

        right_layout.addWidget(self.agency_banner)

        # Vehicle Table for Selected Agency
        self.agency_vehicle_table = QTableWidget()
        self.agency_vehicle_table.setColumnCount(5)
        self.agency_vehicle_table.setHorizontalHeaderLabels([
            "Sl No", "Vehicle Reg Number", "Rate per Load (₹)", "Current Loads", "Next Month Loads (Edit Here)"
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
        layout.addWidget(self.contacts_table)

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
        self.dispatch_table.setColumnCount(7)
        self.dispatch_table.setHorizontalHeaderLabels([
            "Agency", "Inv No", "Grand Total (₹)", "WhatsApp Numbers", "PDF Invoice", "WhatsApp Status", "Action"
        ])
        self.dispatch_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
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
                self.starting_inv_spin.setValue(max_inv + 1)

            self.populate_agency_list()
            self.populate_contacts_table()
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

            self.parsed_agencies.append({
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
            })

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

            # Rate
            item_rate = QTableWidgetItem(f"{float(v['rate']):.2f}")
            item_rate.setFlags(item_rate.flags() ^ Qt.ItemFlag.ItemIsEditable)
            item_rate.setForeground(QColor("#f8fafc"))
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
            self.agency_vehicle_table.setItem(r_idx, 4, item_next)

            subtotal += float(v['loads']) * float(v['rate'])

        grand_total = subtotal * 1.18
        self.lbl_agency_subtotal.setText(f"Total Subtotal: ₹ {subtotal:.2f}")
        self.lbl_agency_grand.setText(f"Grand Total (incl 18% GST): ₹ {grand_total:.2f}")

        self.agency_vehicle_table.blockSignals(False)

    def on_load_cell_changed(self, row, col):
        if col != 4 or not self.current_selected_sheet:
            return

        item = self.agency_vehicle_table.item(row, 4)
        if not item:
            return

        try:
            new_loads = float(item.text().strip())
        except ValueError:
            new_loads = 0.0

        v_list = self.agency_vehicles_data.get(self.current_selected_sheet, [])
        if row < len(v_list):
            v_list[row]['loads'] = new_loads
            self.display_agency_vehicles(self.current_selected_sheet)

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
            self.log(f"Added vehicle {data['vehicle_no']} to agency {self.current_selected_sheet}.")

    def view_current_agency_pdf(self):
        """Opens the generated PDF invoice for the currently selected agency."""
        if not self.current_selected_sheet:
            QMessageBox.warning(self, "Warning", "Please select an agency first.")
            return

        if not self.generated_pdf_list:
            QMessageBox.warning(self, "PDF Not Generated", "Please click '1. Generate All PDFs & Excel' first to create the PDF invoice.")
            return

        pdf_info = next((p for p in self.generated_pdf_list if p['sheet_name'] == self.current_selected_sheet), None)
        if pdf_info and os.path.exists(pdf_info['pdf_path']):
            self.open_pdf_file(pdf_info['pdf_path'])
        else:
            QMessageBox.warning(self, "File Not Found", f"PDF invoice for {self.current_selected_sheet} not found.")

    def open_pdf_file(self, pdf_path):
        """Launches the system default PDF viewer for the given PDF path."""
        if os.path.exists(pdf_path):
            try:
                os.startfile(pdf_path)
                self.log(f"Opened PDF file: {pdf_path}")
            except Exception as e:
                self.log(f"Error opening PDF: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to open PDF:\n{str(e)}")
        else:
            QMessageBox.warning(self, "File Not Found", f"PDF file does not exist:\n{pdf_path}")

    def open_output_folder(self):
        """Opens the generated PDF output directory in Windows Explorer."""
        target_dir = self.current_pdf_dir if os.path.exists(self.current_pdf_dir) else self.output_dir
        if not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)
        try:
            subprocess.Popen(f'explorer "{os.path.abspath(target_dir)}"')
            self.log(f"Opened output directory: {target_dir}")
        except Exception as e:
            self.log(f"Error opening folder: {str(e)}")

    def send_current_agency_whatsapp(self):
        """Sends WhatsApp invoice for the currently selected agency in Tab 1."""
        if not self.current_selected_sheet:
            QMessageBox.warning(self, "Warning", "Please select an agency first.")
            return

        if not self.generated_agency_list or not self.generated_pdf_list:
            QMessageBox.warning(self, "PDFs Not Generated", "Please click '1. Generate All PDFs & Excel' first to generate current invoices.")
            return

        agency_info = next((a for a in self.generated_agency_list if a['sheet_name'] == self.current_selected_sheet), None)
        pdf_info = next((p for p in self.generated_pdf_list if p['sheet_name'] == self.current_selected_sheet), None)

        if not agency_info or not pdf_info:
            QMessageBox.warning(self, "Warning", "Invoice not generated for this agency yet.")
            return

        row_idx = 0
        for i, a in enumerate(self.generated_agency_list):
            if a['sheet_name'] == self.current_selected_sheet:
                row_idx = i
                break

        item = {
            'row_idx': row_idx,
            'agency_name': agency_info['agency_name'],
            'phone': self.contacts_mgr.get_phones(self.current_selected_sheet),
            'inv_no': agency_info['invoice_no'],
            'month_desc': agency_info['month_desc'],
            'pdf_path': pdf_info['pdf_path']
        }

        self.send_single_agency_whatsapp(item)

    def populate_contacts_table(self):
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
                agency_meta_map=self.agency_meta_map
            )
            saved_excel_path = getattr(mgr, 'last_output_file_path', excel_out_path)
            self.log(f"Saved updated master Excel: {saved_excel_path}")

            # 2. Batch generate PDFs
            pdf_dir = os.path.join(self.output_dir, f"Invoices_{target_month.replace(' ', '_')}")
            self.current_pdf_dir = pdf_dir
            pdf_gen = PDFGenerator(pdf_dir)
            self.generated_pdf_list = pdf_gen.batch_generate(self.generated_agency_list)
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
            item_aname.setForeground(QColor("#f8fafc"))
            self.dispatch_table.setItem(idx, 0, item_aname)

            item_inv = QTableWidgetItem(str(a['invoice_no']))
            item_inv.setForeground(QColor("#f8fafc"))
            self.dispatch_table.setItem(idx, 1, item_inv)

            item_total = QTableWidgetItem(f"{a['grand_total']:.2f}")
            item_total.setForeground(QColor("#f8fafc"))
            self.dispatch_table.setItem(idx, 2, item_total)

            item_phone = QTableWidgetItem(phone_display)
            item_phone.setForeground(QColor("#38bdf8") if phones else QColor("#ef4444"))
            self.dispatch_table.setItem(idx, 3, item_phone)

            # Interactive View PDF Button Column
            view_btn = QPushButton(f"👁️ {os.path.basename(pdf_path)}")
            view_btn.setObjectName("viewPdfBtn")
            view_btn.clicked.connect(lambda _, path=pdf_path: self.open_pdf_file(path))
            self.dispatch_table.setCellWidget(idx, 4, view_btn)

            item_status = QTableWidgetItem("Ready to Send" if phones else "Missing Phone")
            item_status.setForeground(QColor("#e2e8f0") if phones else QColor("#ef4444"))
            self.dispatch_table.setItem(idx, 5, item_status)

            # Individual Send Action Button
            send_btn = QPushButton("📱 Send")
            send_btn.setObjectName("singleSendBtn")
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

        row_idx = item['row_idx']
        agency_name = item['agency_name']
        phone = self.contacts_mgr.get_phones(item['sheet_name']) if item.get('sheet_name') else item['phone']

        if not phone:
            QMessageBox.warning(self, "Missing Phone", f"No phone number configured for {agency_name}. Please add it in the WhatsApp Contacts tab.")
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
