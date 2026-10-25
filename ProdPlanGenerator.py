"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Fardan Apex --- ProdPlanGenerator ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
This software automates processing factory orders by generating, organizing,
and exporting production documents into structured PDF files.

Author: Behnam Rabieyan
Company: Garma Gostar Fardan
Created: 2025
"""

# === Standard Library ===
import os
import re
import sys
import json
import time
import shutil
import tempfile
import html as html_lib

import keyring

from erp_repository import ErpRepository
from order_sheet_pdf import create_order_sheet_pdf, merged_pdf_filename, jalali_date_text

# === Third-Party Libraries ===
import pandas as pd
import xlwings as xw
from PyPDF2 import PdfWriter
from PyQt5.QtCore import QObject, QThread, pyqtSignal, Qt, QTimer, QPointF, QRectF
from PyQt5.QtGui import (
    QFont, QIcon, QPixmap, QFontDatabase, QPainter, QPainterPath, QColor, QPolygonF, QPen
)
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTextEdit, QLabel, QMessageBox, QSplashScreen, QProgressBar, QStyle,
    QGroupBox, QDialog, QLineEdit, QFileDialog, QCheckBox, QRadioButton,
    QListWidget, QListWidgetItem, QComboBox, QGridLayout, QFrame, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QStackedWidget
)

# ==============================================================================
# Helper Function for File Paths
# ==============================================================================
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller. """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ==============================================================================
# Configuration Management
# ==============================================================================
class ConfigManager:
    """ Manages loading and saving of application settings from a JSON file. """
    def __init__(self, config_file='config.json'):
        self.config_file = resource_path(config_file)
        self.settings = {}
        self.load()
    
    def load(self):
        """ Loads settings from the JSON file. """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self._create_default_config()
    
    def save(self):
        """ Saves the current settings to the JSON file. """
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.settings, f, indent=4, ensure_ascii=False)
    
    def _create_default_config(self):
        """ Creates a default configuration file if one doesn't exist. """
        self.settings = {
            "database_file_path": "",
            "output_base_path": "",
            "sql_server": "",
            "sql_database": "FARDAN_OLTP",
            "sql_driver": "ODBC Driver 18 for SQL Server",
            "delete_temp_files": True,
            "create_preparation_excel": True,
            "print_preparation_pdf": True,
            "print_timing_pdf": True,
            "report_level": "detailed"
        }
        self.save()

CONFIG = ConfigManager()

# ==============================================================================
# Constants
# ==============================================================================
# --- Sheet and Column Names ---
ORDER_SHEET_NAME = "OrderList"
DATABASE_SHEET_NAME = "LOM"
COL_ORDER_NUM = "شماره سفارش"
COL_PRODUCT_CODE = "کد محصول"
COL_QUANTITY = "تعداد"

# ==============================================================================
# Main window size settings (pixels) - change these numbers to tune the window height
# ==============================================================================
# The window opens at the SHORTEST height its boxes allow, so these numbers decide the final height:
# LOWER them for a shorter window, RAISE them for a taller one. Both boxes scroll when they have
# more content than fits, so they never make the window taller than this.
ORDER_TABLE_VISIBLE_ROWS = 4   # rows of the orders table visible before it scrolls
ORDER_ROW_HEIGHT = 30          # height of one table row (and of the table header)
LOG_BOX_MIN_HEIGHT = 110       # minimum height of the "گزارش پردازش" box
WINDOW_DEFAULT_WIDTH = 950    # window width when the program opens
WINDOW_DEFAULT_HEIGHT = 500   # None = shortest possible height; or a number such as 500 to choose it
                               # (never below the minimum the boxes need, never above the safety limit)
# Safety limit: the window is never taller than this fraction of the usable screen height
# (the Windows taskbar is not counted as usable), so it can not slide under the taskbar.
WINDOW_MAX_HEIGHT_RATIO = 0.95

# --- Input Cells in LOM Sheet ---
CELL_PRODUCT_CODE = "I4"
CELL_CHECK = "D1"
CELL_QUANTITY = "J6"
CELL_ORDER_NUM_DB = "W3"

# --- Main Print Jobs from LOM Sheet ---
LOM_PRINT_JOBS = [
    {"suffix": "LOM", "type": "main"},
    {"suffix": "زمانسنجی", "type": "timing"},
    {"suffix": "آماده سازی", "type": "preparation"}
]

# --- Conditional Check Parameter ---
CONDITIONAL_CHECK_CELL = 'D3'

# --- Conditional Sheets ---
MF_SHEET_NAME = "برنامه فنرپیچ"
ST_SHEET_NAME = "برنامه سیم تابنده"
KL_SHEET_NAME = "برنامه خم لوله‌ای"

# --- Conditional Sheet Print Settings ---
MF_CONFIG = {
    "print_range": "B2:Y54",
    "cell_product": "K5",
    "cell_order": "W5",
    "cell_flag": "Z47",
    "check_cell": "P49"
}
ST_CONFIG = {
    "print_range": "B2:Y68",
    "cell_product": "G7",
    "cell_flag": "Z63",
    "check_cell": "P64"
}
KL_CONFIG = {
    "print_range": "B1:L41",
    "cell_product": "E2"
}

# ==============================================================================
# Directory Scanner Worker (for non-blocking startup)
# ==============================================================================
class DirectoryScannerWorker(QObject):
    """ Scans the order directory in a background thread to keep the UI responsive. """
    scan_complete = pyqtSignal(list, list)
    finished = pyqtSignal()

    def __init__(self, path):
        super().__init__()
        self.path = path
    
    def run(self):
        """ Performs the directory scan and emits the results. """
        confirmed, pending = [], []
        if self.path and os.path.isdir(self.path):
            for filename in os.listdir(self.path):
                if not filename.lower().endswith('.pdf'):
                    continue
                match = re.search(r'\((\d+)\)', filename)
                if match:
                    order_num = match.group(1)
                    base_name = os.path.splitext(filename)[0]
                    if re.search(r'\s*ok$', base_name, re.IGNORECASE):
                        if order_num not in confirmed:
                            confirmed.append(order_num)
                    else:
                        if order_num not in pending:
                            pending.append(order_num)
        
        self.scan_complete.emit(confirmed, pending)
        self.finished.emit()

# ==============================================================================
# Animated "busy" bar (striped, no percentage)
# ==============================================================================
class StripedProgressBar(QWidget):
    """ Progress bar with moving stripes. While busy, the stripes keep moving so the user can see the
    app is working; the solid part grows with set_progress(0..1) as orders are completed. """
    BAR_COLOR = "#2e7dff"      # filled part
    TRACK_COLOR = "#dce6f7"    # empty part
    STRIPE_WIDTH = 14          # width of each stripe (px)
    INTERVAL_MS = 30           # animation speed: lower = faster

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(12)
        self._offset = 0
        self._fraction = 0.0
        self._busy = False
        self._timer = QTimer(self)
        self._timer.setInterval(self.INTERVAL_MS)
        self._timer.timeout.connect(self._advance)

    def set_progress(self, fraction):
        self._fraction = max(0.0, min(1.0, float(fraction)))
        self.update()

    def start(self):
        self._busy = True
        self._timer.start()
        self.update()

    def stop(self):
        self._busy = False
        self._timer.stop()
        self.update()

    def _advance(self):
        self._offset = (self._offset + 1) % (self.STRIPE_WIDTH * 2)
        self.update()

    def _draw_stripes(self, painter, left, right, height, color):
        painter.setPen(Qt.NoPen)
        painter.setBrush(color)
        stripe = self.STRIPE_WIDTH
        step = stripe * 2
        x = left - height - step + self._offset
        while x < right:
            painter.drawPolygon(QPolygonF([
                QPointF(x, height), QPointF(x + stripe, height),
                QPointF(x + stripe + height, 0), QPointF(x + height, 0),
            ]))
            x += step

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        height = rect.height()
        path = QPainterPath()
        path.addRoundedRect(rect, height / 2, height / 2)
        painter.setClipPath(path)
        painter.fillPath(path, QColor(self.TRACK_COLOR))
        if self._busy:   # faint moving stripes over the whole bar: "still working"
            self._draw_stripes(painter, 0, self.width(), height, QColor(46, 125, 255, 45))
        fill_width = rect.width() * self._fraction
        if fill_width > 0:
            painter.fillRect(QRectF(rect.left(), rect.top(), fill_width, height), QColor(self.BAR_COLOR))
            if self._busy:
                painter.setClipRect(QRectF(rect.left(), rect.top(), fill_width, height))
                self._draw_stripes(painter, 0, self.width(), height, QColor(255, 255, 255, 80))

# ==============================================================================
# Core Application Logic (Worker Thread)
# ==============================================================================
class Worker(QObject):
    """ Handles the core data processing in a separate thread. """
    status_update = pyqtSignal(str)
    busy_changed = pyqtSignal(bool)   # True: stripes moving, False: stopped
    progress_changed = pyqtSignal(int, int)   # (orders done, orders total)
    finished = pyqtSignal()
    error_signal = pyqtSignal(str, str)
    warning_signal = pyqtSignal(str, str)
    info_signal = pyqtSignal(str, str)

    def __init__(self, order_numbers_str, config_settings):
        super().__init__()
        self.order_numbers_str = order_numbers_str
        self.config = config_settings
        self.db_wb = None # To hold the workbook object
    
    # --- Report detail level (the "report_level" setting) -----------------------------
    # "detailed": every message is shown.  "summary": only the messages matched below.
    _SUMMARY_MARKS = ("✘", "❗", "⚠", "خطا")                    # errors and warnings
    _SUMMARY_PHRASES = (
        "ادغام شده برای سفارش",                                 # final merged PDFs of an order
        "فایل اکسل آماده سازی برای سفارش",                       # final preparation Excel of an order
        "💯",                                                   # end of processing
    )                                                           # (order headers start with "=====")

    def _log(self, message):
        """ Sends a message to the report box, respecting the chosen report detail level. """
        if self.config.get("report_level", "detailed") == "summary":
            text = message.strip()
            if not (text.startswith("=====")
                    or any(mark in text for mark in self._SUMMARY_MARKS + self._SUMMARY_PHRASES)):
                return
        self.status_update.emit(message)

    def find_last_numeric_row(self, sheet, search_range):
        """ Finds the last row with a numeric value in a given range. """
        values = sheet.range(search_range).options(ndim=1).value
        start_row = sheet.range(search_range).row
        for i in range(len(values) - 1, -1, -1):
            if isinstance(values[i], (int, float)) and values[i] is not None:
                return start_row + i
        return 0
    
    def print_conditional_sheet(self, sheet, product_code, pdf_filepath, config, order_num=None):
        """ Prints a conditional sheet to PDF after setting necessary values. """
        try:
            sheet.range(config['cell_product']).value = product_code
            if order_num and 'cell_order' in config:
                sheet.range(config['cell_order']).value = order_num
            if 'check_cell' in config and 'cell_flag' in config:
                check_val = str(sheet.range(config['check_cell']).value).strip().upper()
                sheet.range(config['cell_flag']).value = (check_val == 'FALSE')
            sheet.range(config['print_range']).api.ExportAsFixedFormat(0, pdf_filepath)
            self._log(
                f"    ✔ چاپ {sheet.name} انجام شد:\n"
            )
            return True
        except Exception as e:
            self._log(
                f"    ✘ خطا در چاپ {sheet.name}: {e}\n"
            )
            return False
    
    def _process_product(self, product_code, order_num, quantity, order_folder, db_sheet, mergers, files_to_delete, preparation_excel_data, technical_drawing_paths):
        """
        Processes a single product code (main or sub-component).
        Prints all necessary documents and updates mergers and file lists.
        """
        # Set main values in the database sheet
        db_sheet.range(CELL_ORDER_NUM_DB).value = order_num
        db_sheet.range(CELL_QUANTITY).value = quantity
        db_sheet.range(CELL_PRODUCT_CODE).value = product_code

        # Check if the product code is valid in the database
        if str(db_sheet.range(CELL_CHECK).value).strip().lower() == 'empty':
            self._log(
                f"  ❗ هشدار: کد محصول {product_code} در دیتابیس نامعتبر است. این آیتم نادیده گرفته شد.\n"
            )
            return mergers, files_to_delete, preparation_excel_data
        
        main_merger, preparation_merger, timing_merger = mergers

        # --- Process standard LOM jobs ---
        for job in LOM_PRINT_JOBS:
            suffix, job_type = job['suffix'], job['type']
            has_data, print_range = True, ""

            if suffix == "LOM":
                last_row = self.find_last_numeric_row(db_sheet, 'B5:B65')
                print_range = f"B1:G{last_row}" if last_row > 0 else ""
            elif suffix == "زمانسنجی":
                if db_sheet.range('P9').value is None: has_data = False
                else: last_row = self.find_last_numeric_row(db_sheet, 'N9:N47'); print_range = f"N4:Q{last_row}" if last_row > 0 else ""
            elif suffix == "آماده سازی":
                if db_sheet.range('U5').value is None: has_data = False
                else: last_row = self.find_last_numeric_row(db_sheet, 'S5:S24'); print_range = f"S1:Y{last_row}" if last_row > 0 else ""
            
            if not print_range: has_data = False

            if not has_data:
                self._log(
                    f"    ✘ {suffix} اطلاعاتی برای پردازش ندارد.\n"
                    )
                continue

            print_this_pdf = True
            if suffix == "زمانسنجی" and not self.config.get('print_timing_pdf', True):
                print_this_pdf = False
                self._log(
                    f"    - چاپ {suffix} بر اساس تنظیمات غیرفعال است.\n"
                )
            if suffix == "آماده سازی" and not self.config.get('print_preparation_pdf', True):
                print_this_pdf = False
                self._log(
                    f"    - چاپ {suffix} بر اساس تنظیمات غیرفعال است.\n"
                )
            
            if print_this_pdf:
                pdf_filepath = os.path.join(order_folder, f"{product_code}_{suffix}.pdf")
                db_sheet.range(print_range).api.ExportAsFixedFormat(0, pdf_filepath)
                files_to_delete.append(pdf_filepath)
                if job_type == 'main': main_merger.append(pdf_filepath)
                elif job_type == 'preparation': preparation_merger.append(pdf_filepath)
                elif job_type == 'timing': timing_merger.append(pdf_filepath)
                self._log(
                    f"    ✔ پرینت {suffix} برای {product_code} ذخیره شد.\n"
                )
            
            if suffix == "آماده سازی" and self.config.get('create_preparation_excel', True):
                try:
                    prep_data_range = db_sheet.range('T5:V24').options(ndim=2).value
                    for row_data in prep_data_range:
                        if row_data[0]:
                            new_row = {"شماره سفارش": order_num, "کد محصول": product_code, "شرح کالا": row_data[0], "تعداد": row_data[1], "اندازه برش": row_data[2]}
                            preparation_excel_data.append(new_row)
                    self._log(
                        f"    ✔ داده‌های اکسل آماده‌سازی برای {product_code} استخراج شد.\n"
                    )
                except Exception as e:
                    self._log(
                        f"    ✘ خطا در استخراج داده‌های اکسل آماده‌سازی: {e}\n"
                    )

        # --- HYBRID SOLUTION: Determine product type reliably ---
        product_type = ""
        code_upper = product_code.upper()

        # Step 1: First, check for TS/TF cases by their reliable prefix.
        if code_upper.startswith('TS-'):
            product_type = 'TS'
        elif code_upper.startswith('TF-'):
            product_type = 'TF'
        
        # Step 2: If it's not a special case, fall back to the trusted method of reading cell D3 for all other product types.
        if not product_type:
            product_type = str(db_sheet.range(CONDITIONAL_CHECK_CELL).value)[:2].upper()

        # --- Process conditional sheets and drawings based on the correctly identified product_type ---
        
        if product_type == 'MF':
            mf_sheet = self.db_wb.sheets[MF_SHEET_NAME]
            pdf_filepath = os.path.join(order_folder, f"{product_code}_{MF_SHEET_NAME}.pdf")
            if self.print_conditional_sheet(mf_sheet, product_code, pdf_filepath, MF_CONFIG, order_num=order_num):
                main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)
        
        if product_type in ('DS', 'DF', 'NL', 'DL'):
            st_sheet = self.db_wb.sheets[ST_SHEET_NAME]
            pdf_filepath = os.path.join(order_folder, f"{product_code}_{ST_SHEET_NAME}.pdf")
            if self.print_conditional_sheet(st_sheet, product_code, pdf_filepath, ST_CONFIG):
                main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)

        if product_type in ('NL', 'DL'):
            kl_sheet = self.db_wb.sheets[KL_SHEET_NAME]
            pdf_filepath = os.path.join(order_folder, f"{product_code}_{KL_SHEET_NAME}.pdf")
            if self.print_conditional_sheet(kl_sheet, product_code, pdf_filepath, KL_CONFIG):
                main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)

        if product_type in ('TS', 'TF'):
            self._log(
                f"    - در حال بررسی نقشه فنی برای محصول ({product_code})\n"
            )

        drawing_source_folder = technical_drawing_paths.get(product_type)
        if drawing_source_folder:
            source_drawing_path = os.path.join(drawing_source_folder, f"{product_code}.pdf")
            dest_drawing_path = os.path.join(order_folder, f"{product_code}_نقشه.pdf")
            
            if os.path.exists(source_drawing_path):
                shutil.copy(source_drawing_path, dest_drawing_path)
                main_merger.append(dest_drawing_path)
                files_to_delete.append(dest_drawing_path)
                self._log(f"    ✔ نقشه فنی برای {product_code} کپی شد.\n")
            else:
                self._log(f"    ✘ هشدار: نقشه فنی {os.path.basename(source_drawing_path)} یافت نشد.\n")

        return (main_merger, preparation_merger, timing_merger), files_to_delete, preparation_excel_data

    def run(self):
        """ Main processing logic. """
        try:
            database_file_path = os.path.normpath(self.config['database_file_path'])
            output_base_path = os.path.normpath(self.config['output_base_path'])
            
            technical_drawing_paths = {
                "TS": r"\\fileserver\mohandesi\PDF Plan\ترموسوئیچ",
                "TF": r"\\fileserver\mohandesi\PDF Plan\ترموفیوز\نقشه های معتبر",
                "DS": r"\\fileserver\mohandesi\PDF Plan\هیتر سیمی\نقشه های معتبر",
                "DF": r"\\fileserver\mohandesi\PDF Plan\فویلی\نقشه های معتبر",
                "NL": r"\\fileserver\mohandesi\PDF Plan\لوله ای\نقشه های معتبر",
                "DL": r"\\fileserver\mohandesi\PDF Plan\لوله ای\نقشه های معتبر",
                "MF": r"\\fileserver\mohandesi\PDF Plan\میله ای\نقشه های معتبر",
                "MR": r"\\fileserver\mohandesi\PDF Plan\میله ای\نقشه های معتبر"
            }

            if not all([database_file_path, output_base_path]):
                self.error_signal.emit(
                    "مسیرها تنظیم نشده",
                    "لطفا مسیر فایل اکسل مهندسی و پوشه خروجی را در تنظیمات مشخص کنید."
                )
                self.finished.emit()
                return

            # The UI accepts/displays only the numeric part, while ERP stores ORD-<number>.
            entered_order_numbers = [
                re.sub(r"^ORD-", "", num.strip(), flags=re.IGNORECASE)
                for num in self.order_numbers_str.strip().split('\n') if num.strip()
            ]
            if not entered_order_numbers:
                self.error_signal.emit(
                    "ورودی خالی", 
                    "هیچ شماره سفارشی برای پردازش وارد نشده است."
                )
                self.finished.emit()
                return

            # Avoid double-prefixing if a user pastes a number in ORD-<number> format.
            order_numbers_list = [f"ORD-{number}" for number in entered_order_numbers]
            self._log(
                f"شماره‌های سفارش برای پردازش:\n{entered_order_numbers}\n"
            )
            self._log("⏳ در حال دریافت سفارش‌ها از ERP (فقط‌خواندنی) . . .\n")
            erp_orders_raw = ErpRepository(self.config).get_orders_for_processing(order_numbers_list)
            if not erp_orders_raw:
                self.warning_signal.emit("یافت نشد", "برای شماره‌های واردشده، سفارش فعالی در ERP یافت نشد.")
                self.finished.emit()
                return

            # Use numeric order numbers throughout the app after retrieving matching ORD-<number> records.
            erp_orders = {}
            for stored_order_number, order in erp_orders_raw.items():
                numeric_order_number = re.sub(r"^ORD-", "", str(stored_order_number).strip(), flags=re.IGNORECASE)
                erp_orders[numeric_order_number] = order
                if isinstance(order, dict) and "order_number" in order:
                    order["order_number"] = numeric_order_number

            rows = []
            for order_number, order in erp_orders.items():
                for item in order["items"]:
                    rows.append({COL_ORDER_NUM: order_number, COL_PRODUCT_CODE: item.get("technical_code") or "", COL_QUANTITY: item.get("quantity") or 0})
            filtered_df = pd.DataFrame(rows)
            missing_orders = [number for number in entered_order_numbers if number not in erp_orders]
            if missing_orders:
                self._log(f"⚠️ سفارش‌های فعال یافت‌نشده: {', '.join(missing_orders)}\n")
            self._log(f"🔎 تعداد {len(filtered_df)} قلم از ERP دریافت شد.\n")

            self._log(
                "   🌀 شروع پردازش اکسل دیتابیس محصولات و چاپ برگه‌ها . . .\n"
            )
            self.busy_changed.emit(True)   # the animated bar appears right after the message above
            with xw.App(visible=False) as app:
                self.db_wb = app.books.open(database_file_path, read_only=True)
                db_sheet = self.db_wb.sheets[DATABASE_SHEET_NAME]

                total_orders = int(filtered_df[COL_ORDER_NUM].nunique())
                processed_orders = 0
                self.progress_changed.emit(0, total_orders)
                for order_num, group in filtered_df.groupby(COL_ORDER_NUM.strip()):
                    order_num_str = str(order_num)
                    self._log(
                        f"=======   شروع پردازش سفارش شماره {order_num_str}   =======\n"
                    )
                    order_folder = os.path.join(output_base_path, order_num_str)
                    os.makedirs(order_folder, exist_ok=True)
                    main_merger, preparation_merger, timing_merger = PdfWriter(), PdfWriter(), PdfWriter()
                    files_to_delete, original_order_filename = [], None
                    order_metadata = erp_orders.get(order_num_str)
                    if order_metadata:
                        order_sheet_path = os.path.join(order_folder, f"برگه سفارش({order_num_str}).pdf")
                        try:
                            create_order_sheet_pdf(order_metadata, order_sheet_path, resource_path("IRAN.ttf"))
                            main_merger.append(order_sheet_path)
                            files_to_delete.append(order_sheet_path)
                            self._log(f"✔ برگه سفارش از ERP ساخته شد: {order_sheet_path}\n")
                        except Exception as error:
                            self._log(f"✘ ساخت برگه سفارش ERP ناموفق بود: {error}\n")
                    preparation_excel_data = []
                    preparation_excel_headers = [
                        "ردیف",
                        "شماره سفارش",
                        "کد محصول",
                        "شرح کالا",
                        "تعداد",
                        "اندازه برش",
                        "تاریخ نیاز",
                        "توضیحات",
                        "امضای تحویل گیرنده"
                    ]

                    # Current order sheet is generated from ERP; legacy order PDFs are not used.
                    for _, row in group.iterrows():
                        original_product_code = str(row[COL_PRODUCT_CODE.strip()])
                        quantity = row[COL_QUANTITY.strip()]
                        self._log(
                            f"\n   ✨  بررسی کد محصول {original_product_code}\n"
                        )
                        
                        valid_product_codes = []
                        db_sheet.range(CELL_PRODUCT_CODE).value = original_product_code
                        if str(db_sheet.range(CELL_CHECK).value).strip().lower() != 'empty':
                            valid_product_codes.append(original_product_code)
                        else:
                            for suffix in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                                variant_code = f"{original_product_code}{suffix}"
                                db_sheet.range(CELL_PRODUCT_CODE).value = variant_code
                                if str(db_sheet.range(CELL_CHECK).value).strip().lower() != 'empty':
                                    valid_product_codes.append(variant_code)
                                else: break
                        
                        if not valid_product_codes:
                            self._log(
                                f"  ❗ هشدار: محصول {original_product_code} نامعتبر است. این آیتم نادیده گرفته شد.\n"
                            )
                            continue
                        
                        for final_code in valid_product_codes:
                            self._log(
                                f"          🚀 شروع فرآیند چاپ برای کد محصول {final_code}\n"
                            )
                            
                            mergers = (main_merger, preparation_merger, timing_merger)
                            mergers, files_to_delete, preparation_excel_data = self._process_product(
                                final_code, order_num_str, quantity, order_folder, db_sheet, mergers, 
                                files_to_delete, preparation_excel_data, technical_drawing_paths
                            )
                            main_merger, preparation_merger, timing_merger = mergers

                            sub_components = []
                            try:
                                bom_range = db_sheet.range('C5:C64').options(ndim=1).value
                                for cell_value in bom_range:
                                    if isinstance(cell_value, str):
                                        found_codes = re.findall(r'(TS-\d+|TF-\d+)', cell_value, re.IGNORECASE)
                                        if found_codes:
                                            sub_components.extend(found_codes)
                                sub_components = sorted(list(set(sub_components)))
                            except Exception as e:
                                self._log(
                                    f"  ✘ خطا در جستجوی قطعات جانبی: {e}\n"
                                )

                            if sub_components:
                                self._log(
                                    f"  🔍 قطعات جانبی یافت شد: {', '.join(sub_components)}\n"
                                )
                                for sub_code in sub_components:
                                    self._log(
                                        f"    🚀 شروع فرآیند چاپ برای {sub_code}\n"
                                    )
                                    mergers, files_to_delete, preparation_excel_data = self._process_product(
                                        sub_code, order_num_str, quantity, order_folder, db_sheet, mergers,
                                        files_to_delete, preparation_excel_data, technical_drawing_paths
                                    )
                                    main_merger, preparation_merger, timing_merger = mergers
                    
                    final_main_pdf_path = None
                    if len(main_merger.pages) > 0:
                        final_main_pdf_path = os.path.join(order_folder, merged_pdf_filename(order_metadata, order_num_str))
                        with open(final_main_pdf_path, "wb") as f: main_merger.write(f)
                        self._log(f"  ✔ فایل اصلی ادغام شده برای سفارش {order_num_str} ذخیره شد.\n")

                    if len(preparation_merger.pages) > 0:
                        with open(os.path.join(order_folder, f"آماده سازی({order_num_str}).pdf"), "wb") as f: preparation_merger.write(f)
                        self._log(f"  ✔ فایل آماده سازی ادغام شده برای سفارش {order_num_str} ذخیره شد.\n")
                    
                    if len(timing_merger.pages) > 0:
                        with open(os.path.join(order_folder, f"زمانسنجی({order_num_str}).pdf"), "wb") as f: timing_merger.write(f)
                        self._log(f"  ✔ فایل زمانسنجی ادغام شده برای سفارش {order_num_str} ذخیره شد.\n")

                    if self.config.get('create_preparation_excel', False) and preparation_excel_data:
                        try:
                            prep_excel_path = os.path.join(order_folder, f"آماده سازی({order_num_str}).xlsx")
                            df_prep = pd.DataFrame(preparation_excel_data)
                            df_prep.insert(0, 'ردیف', range(1, len(df_prep) + 1))
                            df_prep = df_prep.reindex(columns=preparation_excel_headers)
                            df_prep.to_excel(prep_excel_path, index=False, engine='openpyxl')
                            self._log(f"  ✔ فایل اکسل آماده سازی برای سفارش {order_num_str} ذخیره شد.\n")
                        except Exception as e:
                            self._log(f"  ✘ خطا در ذخیره فایل اکسل آماده سازی: {e}\n")
                    
                    main_merger.close(); preparation_merger.close(); timing_merger.close()

                    if self.config['delete_temp_files'] and files_to_delete:
                        self._log(
                            f"\n  ⏳ شروع پاکسازی فایل‌های موقت برای سفارش {order_num_str} . . .\n"
                        )
                        if final_main_pdf_path and final_main_pdf_path in files_to_delete:
                            files_to_delete.remove(final_main_pdf_path)
                            self._log(
                                f"    - فایل نهایی {os.path.basename(final_main_pdf_path)} از لیست حذف خارج شد.\n"
                            )
                        deleted_count = 0
                        for file_path in files_to_delete:
                            try:
                                if os.path.exists(file_path): os.remove(file_path); deleted_count += 1
                            except Exception as e:
                                self._log(
                                    f"    ❗ خطا در حذف فایل {os.path.basename(file_path)}: {e}\n"
                                )
                        self._log(
                            f"    ✔ {deleted_count} فایل موقت با موفقیت حذف شد.\n"
                        )

                    processed_orders += 1
                    self.progress_changed.emit(processed_orders, total_orders)
                
                # Close the workbook defensively. Excel may already have exited,
                # in which case xlwings can raise an RPC error during cleanup.
                try:
                    self.db_wb.close()
                except Exception as close_error:
                    self._log(
                        f"⚠️ بستن فایل اکسل دیتابیس با خطا مواجه شد: {close_error}\n"
                    )
                finally:
                    self.db_wb = None
                self._log(
                    "\n💯 عملیات پردازش با موفقیت به پایان رسید.\n"
                )
                self.info_signal.emit(
                    "اتمام عملیات", "تمام سفارش‌ها با موفقیت پردازش شدند."
                )
        except FileNotFoundError as e:
            msg = (
                f"فایل یا مسیر مورد نظر یافت نشد:\n{e.filename}\n\n"
                   "لطفا از صحت مسیرها در بخش تنظیمات اطمینان حاصل کنید."
                )
            self.error_signal.emit(
                "خطای فایل", msg
            )
            self._log(
                f"خطا در یافتن فایل: {e}\n"
            )
        except Exception as e:
            self.error_signal.emit(
                "خطای کلی", f"یک خطای ناشناخته در برنامه رخ داد:\n{e}"
            )
            self._log(
                f"خطای بحرانی: {e}\n"
            )
        finally:
            # Never let a cleanup failure mask the original processing exception.
            if self.db_wb:
                workbook = self.db_wb
                self.db_wb = None
                try:
                    workbook.close()
                except Exception as close_error:
                    self._log(
                        f"⚠️ هنگام بستن فایل اکسل دیتابیس خطا رخ داد: {close_error}\n"
                    )
            self.busy_changed.emit(False)
            self.finished.emit()

# ==============================================================================
# Settings Dialog Window
# ==============================================================================
# ==============================================================================
# Settings window (modern card design)
# ==============================================================================
SETTINGS_DIALOG_WIDTH = 640    # minimum width of the settings window (pixels)
SETTING_CARD_MIN_HEIGHT = 112  # height of the four cards on the settings home page (pixels)

# Only used by the About window (opened from the settings window).
ABOUT_STYLESHEET = """
    QWidget { background-color: #f5f7fb; }
    QLabel { font-size: 10pt; color: #333; }
"""

SETTINGS_STYLESHEET = """
    #settingsPanel { background-color: #eef3fb; }
    #settingsPanel QLabel { background: transparent; color: #1e2a44; font-size: 10pt; }
    #settingsPanel QLabel#pageTitle { font-size: 16pt; font-weight: bold; color: #14213d; }
    #settingsPanel QLabel#cardName { font-size: 12pt; font-weight: bold; color: #14213d; }
    #settingsPanel QLabel#boldSmall { font-weight: bold; color: #3a4a68; }
    #settingsPanel QLabel#muted { color: #7a879e; font-size: 9pt; }
    #settingsPanel QLabel#formLabel { color: #3a4a68; }
    #settingsPanel QFrame#settingCard { background: #ffffff; border: 1px solid #dce5f3; border-radius: 14px; }
    #settingsPanel QFrame#settingCard:hover { background: #f8fbff; border: 1px solid #9bbcf5; }
    #settingsPanel QFrame#sectionBox { background: #ffffff; border: 1px solid #dce5f3; border-radius: 14px; }
    #settingsPanel QFrame#choiceCard { background: #ffffff; border: 2px solid #dce5f3; border-radius: 14px; }
    #settingsPanel QFrame#choiceCard:hover { border: 2px solid #9bbcf5; }
    #settingsPanel QFrame#choiceCard[selected="true"] { background: #f3f8ff; border: 2px solid #2f7dff; }
    #settingsPanel QLineEdit, #settingsPanel QComboBox {
        background: #ffffff; border: 1px solid #d3dcec; border-radius: 10px;
        padding: 6px 12px; font-size: 10pt; selection-background-color: #2f7dff;
    }
    #settingsPanel QLineEdit:focus, #settingsPanel QComboBox:focus { border: 1px solid #2f7dff; }
    #settingsPanel QCheckBox { background: transparent; font-size: 10pt; spacing: 8px; }
    #settingsPanel QCheckBox::indicator {
        width: 18px; height: 18px; border: 1.5px solid #b8c5dc; border-radius: 5px; background: #ffffff;
    }
    #settingsPanel QCheckBox::indicator:hover { border-color: #2f7dff; }
    #settingsPanel QCheckBox::indicator:checked { background: #2f7dff; border-color: #2f7dff; __CHECK_IMAGE__ }
    #settingsPanel QPushButton { font-size: 10pt; border-radius: 10px; padding: 8px 18px; border: none; }
    #settingsPanel QPushButton#primary {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3f8dff, stop:1 #1e68eb);
        color: white; font-weight: bold;
    }
    #settingsPanel QPushButton#primary:hover {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5a9fff, stop:1 #2f78f5);
    }
    #settingsPanel QPushButton#secondary { background: #ffffff; color: #2a3a5c; border: 1px solid #d3dcec; }
    #settingsPanel QPushButton#secondary:hover { background: #f1f5fc; }
    #settingsPanel QPushButton:disabled { background: #c9d3e5; color: #f3f6fb; border: none; }
"""

def _check_mark_style():
    """ A white tick image for checked boxes, written once to the temp folder (style sheets need a file). """
    try:
        path = os.path.join(tempfile.gettempdir(), "prodplan_check_mark.png")
        if not os.path.exists(path):
            pixmap = QPixmap(18, 18)
            pixmap.fill(Qt.transparent)
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.Antialiasing)
            painter.setPen(QPen(QColor("#ffffff"), 2.2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            painter.drawPolyline(QPolygonF([QPointF(4.5, 9.5), QPointF(7.8, 12.8), QPointF(13.5, 5.8)]))
            painter.end()
            pixmap.save(path, "PNG")
        return 'image: url("' + path.replace("\\", "/") + '");'
    except Exception:
        return ""     # without the image a checked box is simply filled blue


def rtl_text(text):
    """ Forces right-to-left reading order (U+202B ... U+202C): labels that start with a bullet or
    mix Persian and English then show in the correct order. """
    return "\u202b" + text + "\u202c"


class IconTile(QWidget):
    """ Small drawn icon, optionally on a rounded light-blue tile. kind: database | folder | document |
    report | gear. Drawn with QPainter, so it never depends on an icon font. """
    COLOR = QColor("#2f7dff")
    LIGHT = QColor("#8db8ff")
    TILE = QColor("#e6eefc")

    def __init__(self, kind, size=44, tile=True, color=None):
        super().__init__()
        self.kind = kind
        self.tile = tile
        self.color = QColor(color) if color else self.COLOR
        self.setFixedSize(size, size)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        full = QRectF(self.rect())
        if self.tile:
            p.setPen(Qt.NoPen)
            p.setBrush(self.TILE)
            p.drawRoundedRect(full, full.width() * 0.28, full.width() * 0.28)
            pad = full.width() * 0.27
        else:
            pad = full.width() * 0.04
        r = full.adjusted(pad, pad, -pad, -pad)
        getattr(self, "_draw_" + self.kind)(p, r)

    def _draw_database(self, p, r):
        w, eh = r.width(), r.height() * 0.30
        left, top, bottom = r.left(), r.top(), r.bottom()
        body = QPainterPath()
        body.moveTo(left, top + eh / 2)
        body.lineTo(left, bottom - eh / 2)
        body.arcTo(QRectF(left, bottom - eh, w, eh), 180, 180)
        body.lineTo(left + w, top + eh / 2)
        body.arcTo(QRectF(left, top, w, eh), 0, 180)
        p.setPen(Qt.NoPen)
        p.setBrush(self.color)
        p.drawPath(body)
        p.setBrush(self.LIGHT)
        p.drawEllipse(QRectF(left, top, w, eh))
        p.setPen(QPen(self.TILE, 1.6))
        p.setBrush(Qt.NoBrush)
        for k in (1, 2):
            y = top + (r.height() - eh) * k / 3
            p.drawArc(QRectF(left, y, w, eh), 180 * 16, 180 * 16)

    def _draw_folder(self, p, r):
        p.setPen(Qt.NoPen)
        p.setBrush(self.LIGHT)
        p.drawRoundedRect(QRectF(r.left(), r.top() + r.height() * 0.08, r.width() * 0.45, r.height() * 0.30), 2.5, 2.5)
        p.setBrush(self.color)
        p.drawRoundedRect(QRectF(r.left(), r.top() + r.height() * 0.26, r.width(), r.height() * 0.66), 3.5, 3.5)

    def _draw_document(self, p, r):
        page = QRectF(r.left() + r.width() * 0.14, r.top(), r.width() * 0.72, r.height())
        p.setPen(Qt.NoPen)
        p.setBrush(self.color)
        p.drawRoundedRect(page, 3.5, 3.5)
        p.setBrush(QColor("#ffffff"))
        for k in range(3):
            y = page.top() + page.height() * (0.36 + 0.18 * k)
            w = page.width() * (0.62 if k < 2 else 0.38)
            p.drawRoundedRect(QRectF(page.right() - page.width() * 0.16 - w, y, w, 2.2), 1.1, 1.1)

    def _draw_report(self, p, r):
        p.setPen(Qt.NoPen)
        p.setBrush(self.color)
        p.drawRoundedRect(r, 3.5, 3.5)
        p.setPen(QPen(self.TILE, 1.6))
        for k in (1, 2):
            y = r.top() + r.height() * k / 3
            p.drawLine(QPointF(r.left(), y), QPointF(r.right(), y))
        p.drawLine(QPointF(r.center().x(), r.top() + r.height() / 3), QPointF(r.center().x(), r.bottom()))

    def _draw_gear(self, p, r):
        p.translate(r.center())
        radius = r.width() / 2
        p.setPen(Qt.NoPen)
        p.setBrush(self.color)
        for _ in range(8):
            p.drawRoundedRect(QRectF(-radius * 0.17, -radius, radius * 0.34, radius * 0.42), 1.5, 1.5)
            p.rotate(45)
        p.drawEllipse(QPointF(0, 0), radius * 0.76, radius * 0.76)
        p.setBrush(QColor("#eef3fb"))
        p.drawEllipse(QPointF(0, 0), radius * 0.32, radius * 0.32)


class Chevron(QWidget):
    """ Small '<' arrow (points to the left = 'go forward' in a right-to-left window). """
    def __init__(self):
        super().__init__()
        self.setFixedSize(14, 20)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QPen(QColor("#8a96ad"), 2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawPolyline(QPolygonF([QPointF(10, 4), QPointF(4, 10), QPointF(10, 16)]))


class RadioDot(QWidget):
    """ Round radio indicator for the report-level cards. """
    def __init__(self):
        super().__init__()
        self._on = False
        self.setFixedSize(20, 20)

    def set_on(self, on):
        self._on = on
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QPen(QColor("#2f7dff" if self._on else "#b8c5dc"), 2))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(QPointF(10, 10), 8, 8)
        if self._on:
            p.setPen(Qt.NoPen)
            p.setBrush(QColor("#2f7dff"))
            p.drawEllipse(QPointF(10, 10), 4.5, 4.5)


class ClickableFrame(QFrame):
    """ A frame that behaves like a button. """
    clicked = pyqtSignal()

    def __init__(self, object_name):
        super().__init__()
        self.setObjectName(object_name)
        self.setCursor(Qt.PointingHandCursor)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.rect().contains(event.pos()):
            self.clicked.emit()
        super().mouseReleaseEvent(event)


class SettingCard(ClickableFrame):
    """ One card of the settings home page: icon, title, description, live status line, arrow. """
    def __init__(self, icon_kind, title, description):
        super().__init__("settingCard")
        self.setMinimumHeight(SETTING_CARD_MIN_HEIGHT)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(12)
        texts = QVBoxLayout()
        texts.setSpacing(2)
        name = QLabel(title)
        name.setObjectName("cardName")
        desc = QLabel(description)
        desc.setObjectName("muted")
        desc.setWordWrap(True)
        self.status_label = QLabel()
        self.status_label.setObjectName("muted")
        self.status_label.setWordWrap(True)
        texts.addWidget(name)
        texts.addWidget(desc)
        texts.addSpacing(2)
        texts.addWidget(self.status_label)
        layout.addWidget(IconTile(icon_kind, 46))
        layout.addLayout(texts, 1)
        layout.addWidget(Chevron())

    def set_status(self, text, ok=True):
        color = "#1e9e5a" if ok else "#d98200"
        self.status_label.setText(f'<span style="color:{color}">●&nbsp;&nbsp;{text}</span>')


class ChoiceCard(ClickableFrame):
    """ A selectable card (like a radio button) that lists what a report level contains. """
    def __init__(self, title, tagline, items, footer):
        super().__init__("choiceCard")
        self.setProperty("selected", False)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(6)
        head = QHBoxLayout()
        self.dot = RadioDot()
        name = QLabel(title)
        name.setObjectName("cardName")
        head.addWidget(self.dot)
        head.addWidget(name)
        head.addStretch(1)
        layout.addLayout(head)
        tag = QLabel(tagline)
        tag.setObjectName("muted")
        tag.setWordWrap(True)
        layout.addWidget(tag)
        layout.addSpacing(4)
        includes = QLabel("شامل:")
        includes.setObjectName("boldSmall")
        layout.addWidget(includes)
        for text in items:
            row = QHBoxLayout()
            row.setSpacing(6)
            bullet = QLabel("•")
            bullet.setFixedWidth(10)
            label = QLabel(text)
            label.setWordWrap(True)
            row.addWidget(bullet, 0, Qt.AlignTop)
            row.addWidget(label, 1)
            layout.addLayout(row)
        layout.addStretch(1)
        foot = QLabel(footer)
        foot.setObjectName("muted")
        foot.setWordWrap(True)
        layout.addWidget(foot)

    def set_selected(self, selected):
        self.setProperty("selected", selected)
        self.dot.set_on(selected)
        self.style().unpolish(self)
        self.style().polish(self)


REPORT_LEVELS = {
    "summary": ("خلاصه", "فقط نتیجه‌ی کار، بدون جزئیات مراحل", [
        "شروع پردازش هر سفارش",
        "خطاها و هشدارها (مثلاً نقشه‌ی فنی یافت نشد یا کد محصول نامعتبر)",
        "فایل‌های نهایی ساخته‌شده برای هر سفارش",
        "پیام پایان پردازش",
    ], "مناسب برای استفاده‌ی روزمره"),
    "detailed": ("کامل", "همه‌ی مراحل، قدم به قدم", [
        "همه‌ی موارد سطح «خلاصه»",
        "شماره سفارش‌های واردشده، دریافت از ERP و تعداد اقلام",
        "بررسی هر کد محصول و قطعات جانبی",
        "کپی نقشه‌های فنی و استخراج داده‌های اکسل آماده‌سازی",
        "چاپ هر برگه و پاک‌سازی فایل‌های موقت",
    ], "مناسب برای بررسی مشکل و عیب‌یابی"),
}


class SettingsDialog(QDialog):
    """ Settings window: a home page with four cards, each opening its own page. """
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.report_level = "detailed"
        self.setWindowTitle("ویژگی‌ها و تنظیمات")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setMinimumWidth(SETTINGS_DIALOG_WIDTH)
        self.setStyleSheet(SETTINGS_STYLESHEET.replace("__CHECK_IMAGE__", _check_mark_style()))     # set first, so every widget is styled when it is created

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.panel = QWidget()
        self.panel.setObjectName("settingsPanel")    # the style sheet only targets this container
        outer.addWidget(self.panel)
        root = QVBoxLayout(self.panel)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        self.pages = QStackedWidget()
        self.cards = {}
        self.pages.addWidget(self._build_home_page())      # 0
        self.pages.addWidget(self._build_erp_page())       # 1
        self.pages.addWidget(self._build_paths_page())     # 2
        self.pages.addWidget(self._build_options_page())   # 3
        self.pages.addWidget(self._build_report_page())    # 4
        root.addWidget(self.pages, 1)

        bottom = QHBoxLayout()                              # right to left: About ... Save
        self.about_button = QPushButton("درباره برنامه")
        self.about_button.setObjectName("secondary")
        self.save_button = QPushButton("ذخیره تنظیمات")
        self.save_button.setObjectName("primary")
        bottom.addWidget(self.about_button)
        bottom.addStretch(1)
        bottom.addWidget(self.save_button)
        root.addLayout(bottom)

        self._connect_signals()
        self._populate_fields()
        self._refresh_card_status()

    # --------------------------------------------------------------------------
    # Pages
    # --------------------------------------------------------------------------
    def _build_home_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        header = QHBoxLayout()
        header.setSpacing(12)
        titles = QVBoxLayout()
        titles.setSpacing(0)
        title = QLabel("تنظیمات")
        title.setObjectName("pageTitle")
        subtitle = QLabel("تنظیمات برنامه و مسیرهای مورد نیاز")
        subtitle.setObjectName("muted")
        titles.addWidget(title)
        titles.addWidget(subtitle)
        header.addWidget(IconTile("gear", 40, tile=False, color="#6b87b8"))
        header.addLayout(titles)
        header.addStretch(1)
        layout.addLayout(header)

        grid = QGridLayout()                                # column 0 = right side (right-to-left)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(14)
        cards = [
            ("erp", "database", "اتصال به پایگاه‌داده", "تنظیمات اتصال به دیتابیس ERP", 1),
            ("paths", "folder", "مسیرها", "تنظیم مسیرهای فایل‌ها و پوشه‌ها", 2),
            ("options", "document", "گزینه‌های پردازش", "تنظیم رفتار و خروجی‌های برنامه", 3),
            ("report", "report", "سطح گزارش‌ها", "تعیین سطح جزئیات گزارش پردازش", 4),
        ]
        for index, (key, icon, title_text, description, page_index) in enumerate(cards):
            card = SettingCard(icon, title_text, description)
            card.clicked.connect(lambda n=page_index: self.pages.setCurrentIndex(n))
            grid.addWidget(card, index // 2, index % 2)
            self.cards[key] = card
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        layout.addLayout(grid)
        layout.addStretch(1)
        return page

    def _page_shell(self, title_text, icon_kind):
        """ Common frame of an inner page: back button + title on top, a white box for the content. """
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        header = QHBoxLayout()
        header.setSpacing(10)
        back = QPushButton(rtl_text("→  بازگشت"))
        back.setObjectName("secondary")
        back.clicked.connect(self._go_home)
        title = QLabel(title_text)
        title.setObjectName("pageTitle")
        header.addWidget(back)
        header.addWidget(IconTile(icon_kind, 36))
        header.addWidget(title)
        header.addStretch(1)
        layout.addLayout(header)
        box = QFrame()
        box.setObjectName("sectionBox")
        box_layout = QVBoxLayout(box)
        box_layout.setContentsMargins(16, 14, 16, 14)
        box_layout.setSpacing(10)
        layout.addWidget(box)
        layout.addStretch(1)
        return page, box_layout

    @staticmethod
    def _add_row(layout, label_text, widget):
        row = QHBoxLayout()
        label = QLabel(label_text)
        label.setObjectName("formLabel")
        label.setFixedWidth(130)
        row.addWidget(label)
        row.addWidget(widget, 1)
        layout.addLayout(row)

    def _build_erp_page(self):
        page, box = self._page_shell("اتصال به پایگاه‌داده", "database")
        self.sql_server_input = QLineEdit()
        self.sql_server_input.setPlaceholderText("مثال: localhost یا 192.168.1.10,1433 یا SERVER\\INSTANCE")
        self.sql_database_input = QLineEdit()
        self.sql_database_input.setPlaceholderText("نام پایگاه داده")
        self.sql_driver_input = QLineEdit()
        self.sql_driver_input.setPlaceholderText("نام درایور ODBC نصب‌شده")
        self.sql_auth_combo = QComboBox()
        self.sql_auth_combo.addItem(rtl_text("احراز هویت ویندوز (Windows Authentication)"), "windows")
        self.sql_auth_combo.addItem("نام کاربری و رمز SQL Server", "sql")
        self.sql_username_input = QLineEdit()
        self.sql_username_input.setPlaceholderText("نام کاربری SQL؛ برای ویندوز خالی بماند")
        self.sql_password_input = QLineEdit()
        self.sql_password_input.setEchoMode(QLineEdit.Password)
        self.sql_password_input.setPlaceholderText("برای تغییر رمز وارد کنید؛ خالی = استفاده از رمز ذخیره‌شده")
        self.sql_password_input.setToolTip("رمز در Windows Credential Manager ذخیره می‌شود، نه در config.json.")
        self._add_row(box, "آدرس سرور:", self.sql_server_input)
        self._add_row(box, "نام پایگاه داده:", self.sql_database_input)
        self._add_row(box, "درایور ODBC:", self.sql_driver_input)
        self._add_row(box, "روش ورود:", self.sql_auth_combo)
        self._add_row(box, "نام کاربری:", self.sql_username_input)
        self._add_row(box, "رمز عبور:", self.sql_password_input)
        buttons = QHBoxLayout()
        self.test_sql_button = QPushButton("آزمایش اتصال")
        self.test_sql_button.setObjectName("secondary")
        self.test_sql_button.clicked.connect(self._test_sql_connection)
        buttons.addWidget(self.test_sql_button)
        buttons.addStretch(1)
        box.addLayout(buttons)
        return page

    def _build_paths_page(self):
        page, box = self._page_shell("مسیرها", "folder")
        self.path_widgets = {
            "database_file_path": self._create_path_selector("فایل اکسل مهندسی (.xlsm):", "file"),
            "output_base_path": self._create_path_selector("پوشه ذخیره خروجی:", "folder"),
        }
        for layout, _ in self.path_widgets.values():
            box.addLayout(layout)
        return page

    def _build_options_page(self):
        page, box = self._page_shell("گزینه‌های پردازش", "document")
        # rtl_text() keeps labels that mix Persian and English (like "فایل PDF ...") in reading order.
        self.print_prep_pdf_checkbox = QCheckBox(rtl_text("فایل PDF آماده سازی چاپ شود"))
        self.print_timing_pdf_checkbox = QCheckBox(rtl_text("فایل PDF زمانسنجی چاپ شود"))
        self.create_prep_excel_checkbox = QCheckBox(rtl_text("فایل اکسل آماده سازی ایجاد شود"))
        self.delete_temp_checkbox = QCheckBox(rtl_text("پس از پایان پردازش، فایل‌های موقت پاک شوند"))
        # Two rows x two columns, grouped at the RIGHT edge. In this right-to-left window the first
        # item of a horizontal layout is on the right and Qt.AlignLeading is the right edge, so the
        # grid is added BEFORE the stretch.
        grid = QGridLayout()
        grid.setHorizontalSpacing(28)
        grid.setVerticalSpacing(14)
        alignment = Qt.AlignLeading | Qt.AlignVCenter
        grid.addWidget(self.print_prep_pdf_checkbox, 0, 0, alignment=alignment)
        grid.addWidget(self.print_timing_pdf_checkbox, 0, 1, alignment=alignment)
        grid.addWidget(self.create_prep_excel_checkbox, 1, 0, alignment=alignment)
        grid.addWidget(self.delete_temp_checkbox, 1, 1, alignment=alignment)
        row = QHBoxLayout()
        row.addLayout(grid)
        row.addStretch(1)
        box.addLayout(row)
        return page

    def _build_report_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        header = QHBoxLayout()
        header.setSpacing(10)
        back = QPushButton(rtl_text("→  بازگشت"))
        back.setObjectName("secondary")
        back.clicked.connect(self._go_home)
        title = QLabel("سطح جزئیات گزارش پردازش")
        title.setObjectName("pageTitle")
        header.addWidget(back)
        header.addWidget(IconTile("report", 36))
        header.addWidget(title)
        header.addStretch(1)
        layout.addLayout(header)

        row = QHBoxLayout()
        row.setSpacing(14)
        self.level_cards = {}
        for level in ("summary", "detailed"):                # right to left: summary, detailed
            name, tagline, items, footer = REPORT_LEVELS[level]
            card = ChoiceCard(name, tagline, items, footer)
            card.clicked.connect(lambda lv=level: self._set_report_level(lv))
            row.addWidget(card, 1)
            self.level_cards[level] = card
        layout.addLayout(row, 1)
        note = QLabel("خطاها و هشدارها در هر دو سطح نمایش داده می‌شوند. این تنظیم از پردازش بعدی اعمال می‌شود.")
        note.setObjectName("muted")
        note.setWordWrap(True)
        layout.addWidget(note)
        return page

    # --------------------------------------------------------------------------
    # Navigation, live status
    # --------------------------------------------------------------------------
    def _go_home(self):
        self.pages.setCurrentIndex(0)

    def keyPressEvent(self, event):
        """ Esc goes back to the home page first; on the home page it closes the window. """
        if event.key() == Qt.Key_Escape and self.pages.currentIndex() != 0:
            self._go_home()
            return
        super().keyPressEvent(event)

    def _set_report_level(self, level):
        self.report_level = level
        for key, card in self.level_cards.items():
            card.set_selected(key == level)
        self._refresh_card_status()

    def _refresh_card_status(self):
        """ Updates the status line of every card and remembers which sections are incomplete. """
        self.incomplete_sections = []

        missing = []
        if not self.sql_server_input.text().strip():
            missing.append("آدرس سرور")
        if not self.sql_database_input.text().strip():
            missing.append("نام پایگاه داده")
        if not self.sql_driver_input.text().strip():
            missing.append("درایور ODBC")
        if self.sql_auth_combo.currentData() == "sql" and not self.sql_username_input.text().strip():
            missing.append("نام کاربری")
        if missing:
            self.incomplete_sections.append("اتصال به پایگاه‌داده")
            self.cards["erp"].set_status("ناقص: " + "، ".join(missing), False)
        else:
            self.cards["erp"].set_status("تنظیم شده: " + self.sql_database_input.text().strip())

        missing = []
        if not self.path_widgets["database_file_path"][1].text().strip():
            missing.append("فایل اکسل مهندسی")
        if not self.path_widgets["output_base_path"][1].text().strip():
            missing.append("پوشه خروجی")
        if missing:
            self.incomplete_sections.append("مسیرها")
            self.cards["paths"].set_status("ناقص: " + "، ".join(missing), False)
        else:
            self.cards["paths"].set_status("هر دو مسیر تنظیم شده است")

        boxes = (self.print_prep_pdf_checkbox, self.print_timing_pdf_checkbox,
                 self.create_prep_excel_checkbox, self.delete_temp_checkbox)
        self.cards["options"].set_status(f"{fa_digits(sum(box.isChecked() for box in boxes))} از ۴ گزینه فعال")
        self.cards["report"].set_status("سطح فعلی: " + REPORT_LEVELS[self.report_level][0])

    # --------------------------------------------------------------------------
    # Fields
    # --------------------------------------------------------------------------
    def _create_path_selector(self, label_text, selection_mode):
        """ Creates a layout for path selection with a label, line edit, and browse button. """
        layout = QHBoxLayout()
        label = QLabel(rtl_text(label_text))
        label.setObjectName("formLabel")
        label.setFixedWidth(160)
        line_edit = QLineEdit()
        browse_button = QPushButton("انتخاب مسیر")
        browse_button.setObjectName("secondary")
        layout.addWidget(label)
        layout.addWidget(line_edit, 1)
        layout.addWidget(browse_button)
        if selection_mode == "file":
            browse_button.clicked.connect(lambda: self._browse_file(line_edit))
        else:
            browse_button.clicked.connect(lambda: self._browse_folder(line_edit))
        return layout, line_edit

    def _browse_file(self, line_edit):
        """ Opens a file dialog to select an Excel file. """
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل 📄", "", "Excel Files (*.xlsx *.xlsm)")
        if path:
            line_edit.setText(path)

    def _browse_folder(self, line_edit):
        """ Opens a dialog to select a folder. """
        path = QFileDialog.getExistingDirectory(self, "انتخاب پوشه 📁")
        if path:
            line_edit.setText(path)

    def _populate_fields(self):
        """ Fills the settings fields with values from the config manager. """
        settings = self.config_manager.settings
        for key, widget_tuple in self.path_widgets.items():
            widget_tuple[1].setText(settings.get(key, ""))

        self.sql_server_input.setText(settings.get("sql_server", "localhost"))
        self.sql_database_input.setText(settings.get("sql_database", "FARDAN_OLTP"))
        self.sql_driver_input.setText(settings.get("sql_driver", "ODBC Driver 18 for SQL Server"))
        auth_index = self.sql_auth_combo.findData(settings.get("sql_authentication", "windows"))
        self.sql_auth_combo.setCurrentIndex(max(auth_index, 0))
        self.sql_username_input.setText(settings.get("sql_username", ""))
        self.sql_password_input.clear()

        self.print_prep_pdf_checkbox.setChecked(settings.get("print_preparation_pdf", True))
        self.print_timing_pdf_checkbox.setChecked(settings.get("print_timing_pdf", True))
        self.create_prep_excel_checkbox.setChecked(settings.get("create_preparation_excel", True))
        self.delete_temp_checkbox.setChecked(settings.get("delete_temp_files", True))

        level = settings.get("report_level", "detailed")
        self._set_report_level(level if level in REPORT_LEVELS else "detailed")

    def _connect_signals(self):
        """ Connects widget signals to their respective slots. """
        self.save_button.clicked.connect(self.accept)
        self.about_button.clicked.connect(self._show_about)
        for edit in (self.sql_server_input, self.sql_database_input, self.sql_driver_input, self.sql_username_input,
                     self.path_widgets["database_file_path"][1], self.path_widgets["output_base_path"][1]):
            edit.textChanged.connect(self._refresh_card_status)
        self.sql_auth_combo.currentIndexChanged.connect(self._refresh_card_status)
        for box in (self.print_prep_pdf_checkbox, self.print_timing_pdf_checkbox,
                    self.create_prep_excel_checkbox, self.delete_temp_checkbox):
            box.toggled.connect(self._refresh_card_status)

    def _save_settings(self):
        """ Saves the current settings from the dialog to the config manager. """
        for key, widget_tuple in self.path_widgets.items():
            self.config_manager.settings[key] = widget_tuple[1].text()

        settings = self.config_manager.settings
        settings["sql_server"] = self.sql_server_input.text().strip()
        settings["sql_database"] = self.sql_database_input.text().strip()
        settings["sql_driver"] = self.sql_driver_input.text().strip()
        settings["sql_authentication"] = self.sql_auth_combo.currentData()
        settings["sql_username"] = self.sql_username_input.text().strip() if settings["sql_authentication"] == "sql" else ""
        entered_password = self.sql_password_input.text()
        if entered_password and settings["sql_authentication"] == "sql":
            keyring.set_password(
                ErpRepository.credential_service(settings),
                settings["sql_username"],
                entered_password,
            )

        settings["print_preparation_pdf"] = self.print_prep_pdf_checkbox.isChecked()
        settings["print_timing_pdf"] = self.print_timing_pdf_checkbox.isChecked()
        settings["create_preparation_excel"] = self.create_prep_excel_checkbox.isChecked()
        settings["delete_temp_files"] = self.delete_temp_checkbox.isChecked()
        settings["report_level"] = self.report_level

        self.config_manager.save()

    def _test_sql_connection(self):
        """Validate the currently entered SQL settings without saving them."""
        settings = dict(self.config_manager.settings)
        settings.update({
            "sql_server": self.sql_server_input.text().strip(),
            "sql_database": self.sql_database_input.text().strip(),
            "sql_driver": self.sql_driver_input.text().strip(),
            "sql_authentication": self.sql_auth_combo.currentData(),
            "sql_username": self.sql_username_input.text().strip() if self.sql_auth_combo.currentData() == "sql" else "",
        })
        try:
            ErpRepository(settings).test_connection(password_override=self.sql_password_input.text() or None)
        except Exception as exc:
            QMessageBox.warning(
                self, "اتصال ناموفق",
                "اتصال به SQL Server برقرار نشد. آدرس، پورت، نام دیتابیس، درایور و مجوز ورود را بررسی کنید.\n\n"
                f"جزئیات: {exc}",
            )
            return
        QMessageBox.information(self, "اتصال موفق", "اتصال به پایگاه داده با موفقیت آزمایش شد.")

    def accept(self):
        """Save settings and keep credentials out of the JSON configuration."""
        try:
            self._save_settings()
        except Exception as exc:
            QMessageBox.critical(
                self, "ذخیره تنظیمات ناموفق",
                "تنظیمات ذخیره نشد. امکان ذخیره امن رمز در Windows Credential Manager وجود نداشت.\n\n"
                f"جزئیات: {exc}",
            )
            return
        message = "تنظیمات با موفقیت ذخیره شد."
        if self.incomplete_sections:
            message += "\n\nتوجه: این بخش‌ها هنوز ناقص‌اند: " + "، ".join(self.incomplete_sections)
        QMessageBox.information(self, "ذخیره شد", message)
        super().accept()

    def _show_about(self):
        """ Shows the 'About' dialog. """
        dlg = QDialog(self)
        dlg.setStyleSheet(ABOUT_STYLESHEET)
        dlg.setWindowTitle("About")
        dlg.setFixedSize(500, 450)

        main_layout = QVBoxLayout(dlg)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        intro_layout = QVBoxLayout()
        intro_layout.setSpacing(0)
        intro_layout.setContentsMargins(0, 0, 0, 0)
        
        lbl_intro_text = (
            "<h3><b>Fardan Apex — ProdPlanGenerator</b></h3>"
            "<h4>Order PDF Generator Application</h4><br>"
            "Automates the generation of production order documents "
            "from Excel data into consolidated PDFs.<br>"
            "Version: 2.6.19 — © 2025 All Rights Reserved<br>"
            "Developed exclusively for:<br>"
            "Garma Gostar Fardan Co."
        )
        lbl_intro = QLabel(lbl_intro_text)

        lbl_intro.setWordWrap(True)
        lbl_intro.setAlignment(Qt.AlignLeft)
        intro_layout.addWidget(lbl_intro)

        logo = QLabel()
        logo_pix = QPixmap(resource_path("FardanLogoEN.png"))
        if logo_pix.isNull():
            logo_pix = QPixmap(resource_path("FardanLogoFA.png"))
        
        if not logo_pix.isNull():
            logo.setPixmap(logo_pix.scaledToWidth(175, Qt.SmoothTransformation))
        else:
            logo.setText("Fardan Apex")
        logo.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        logo.setContentsMargins(35, 10, 0, 0)
        intro_layout.addWidget(logo)

        main_layout.addLayout(intro_layout)

        dev_layout = QVBoxLayout()
        dev_layout.setSpacing(0)
        dev_layout.setContentsMargins(5, 0, 0, 5)
        font_id = QFontDatabase.addApplicationFont(resource_path("BrittanySignature.ttf"))
        if font_id != -1:
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        else:
            font_family = "Sans Serif"

        lbl_dev_text = (
            f"<b>Design & Development:</b><br>"
            f"<span style='font-family:\"{font_family}\"; font-size:20pt; color:#4169E1;'>"
            f"&nbsp;&nbsp;&nbsp;&nbsp;Behnam Rabieyan</span><br>"
            "website: behnamrabieyan.ir | E-mail: info@behnamrabieyan.ir"
        )
        lbl_dev = QLabel(lbl_dev_text)

        lbl_dev.setWordWrap(True)
        lbl_dev.setAlignment(Qt.AlignLeft)
        dev_layout.addWidget(lbl_dev)

        main_layout.addLayout(dev_layout)
        dlg.exec_()

# ==============================================================================
# Main Application Window
# ==============================================================================
# ==============================================================================
# Main Window (modern card design)
# ==============================================================================
_DIGITS_TO_LATIN = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
_DIGITS_TO_PERSIAN = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
_INPUT_SEPARATORS = " ,،;؛\t-"          # typing one of these after a number adds " - "


def fa_digits(value):
    """ 12 -> ۱۲ """
    return str(value).translate(_DIGITS_TO_PERSIAN)


class ProdPlanApp(QWidget):
    """ Main application window (GUI). """
    def __init__(self):
        super().__init__()
        self.worker = None
        self.thread = None
        self._syncing = False          # True while the table and the order box update each other
        self._last_input_text = ""
        self.initUI()

    # --------------------------------------------------------------------------
    # UI construction
    # --------------------------------------------------------------------------
    def initUI(self):
        """ Initializes the UI components. """
        self.setWindowTitle("سازنده برگه‌های سفارش - ProdPlanGenerator - FardanApex")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.resize(WINDOW_DEFAULT_WIDTH, 600)       # the height is set in fit_to_screen()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.content = QWidget()
        self.content.setObjectName("content")        # the style sheet only targets this container
        # Right-to-left window: in every row the FIRST widget added sits on the RIGHT.
        # To flip the whole window back to left-to-right, change this to Qt.LeftToRight.
        self.content.setLayoutDirection(Qt.RightToLeft)
        outer.addWidget(self.content)

        root = QVBoxLayout(self.content)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)
        body = QHBoxLayout()
        body.setSpacing(10)
        body.addWidget(self._build_orders_card(), 5)   # right: ready orders (first = right-most)
        body.addWidget(self._build_log_card(), 6)      # left:  report + system status
        root.addLayout(body, 1)
        root.addWidget(self._build_input_bar())      # bottom row: label, box, process, settings (right to left)

        self.process_button.clicked.connect(self.start_processing)
        self.settings_button.clicked.connect(self.show_settings)
        self.refresh_button.clicked.connect(self.scan_order_directory)
        self.clear_log_button.clicked.connect(self.status_box.clear)
        self.order_search.textChanged.connect(self._filter_orders)
        self.order_input.textEdited.connect(self._input_edited)
        self.order_input.editingFinished.connect(self._normalize_input)
        self.orders_table.itemChanged.connect(self._on_table_item_changed)
        self.orders_table.cellClicked.connect(self._on_cell_clicked)
        self.orders_table.horizontalHeader().sectionClicked.connect(self._toggle_all_visible)

        self.apply_stylesheet()

    @staticmethod
    def _make_card(object_name="card"):
        card = QFrame()
        card.setObjectName(object_name)
        return card

    def _build_orders_card(self):
        card = self._make_card()
        card.setLayoutDirection(Qt.RightToLeft)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("سفارش‌های آماده برنامه‌ریزی")
        title.setObjectName("cardTitle")
        self.orders_badge = QLabel("۰ سفارش")
        self.orders_badge.setObjectName("badge")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.orders_badge)
        layout.addLayout(header)

        search_row = QHBoxLayout()
        search_row.setSpacing(8)
        self.refresh_button = QPushButton("⟳  بروزرسانی")
        self.refresh_button.setObjectName("primary")
        self.order_search = QLineEdit()
        self.order_search.setPlaceholderText("جستجوی شماره سفارش یا مشتری...")
        self.order_search.setClearButtonEnabled(True)
        search_row.addWidget(self.refresh_button)
        search_row.addWidget(self.order_search, 1)
        layout.addLayout(search_row)

        self.orders_table = self._build_orders_table()
        layout.addWidget(self.orders_table, 1)

        db_card = self._make_card("innerCard")
        db_layout = QHBoxLayout(db_card)
        db_layout.setContentsMargins(10, 5, 10, 5)
        db_title = QLabel("وضعیت سیستم")
        self.db_status_label = QLabel()
        db_layout.addWidget(db_title)
        db_layout.addStretch(1)
        db_layout.addWidget(self.db_status_label)
        layout.addWidget(db_card)
        self._set_db_status("idle")
        return card

    def _build_orders_table(self):
        """ Scrollable table of ready orders: it never grows taller than the window gives it. """
        table = QTableWidget(0, 5)
        table.setLayoutDirection(Qt.RightToLeft)
        table.setHorizontalHeaderLabels(["✓", "شماره سفارش", "مشتری", "تاریخ سفارش", "اقلام"])
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(ORDER_ROW_HEIGHT)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.NoSelection)
        table.setFocusPolicy(Qt.NoFocus)
        table.setShowGrid(False)
        table.setWordWrap(False)
        table.setTextElideMode(Qt.ElideRight)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        header = table.horizontalHeader()
        header.setFixedHeight(ORDER_ROW_HEIGHT)
        header.setHighlightSections(False)
        header.setSectionsClickable(True)            # clicking the "✓" header selects/clears all
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.resizeSection(0, 40)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        # The table is only this tall at minimum (header + N rows); extra window height is added to it.
        table.setMinimumHeight(ORDER_ROW_HEIGHT * (ORDER_TABLE_VISIBLE_ROWS + 1) + 4)
        return table

    def _build_log_card(self):
        card = self._make_card()
        card.setLayoutDirection(Qt.RightToLeft)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("گزارش پردازش")
        title.setObjectName("cardTitle")
        self.clear_log_button = QPushButton("پاک کردن")
        self.clear_log_button.setObjectName("secondary")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.clear_log_button)
        layout.addLayout(header)

        self.status_box = QTextEdit()                  # scrolls inside its own box
        self.status_box.setObjectName("logBox")
        self.status_box.setReadOnly(True)
        self.status_box.setMinimumHeight(LOG_BOX_MIN_HEIGHT)
        layout.addWidget(self.status_box, 1)

        status_card = self._make_card("innerCard")
        status_layout = QVBoxLayout(status_card)
        status_layout.setContentsMargins(10, 6, 10, 8)
        status_layout.setSpacing(4)
        top_row = QHBoxLayout()
        status_title = QLabel("وضعیت کلی سیستم")
        status_title.setObjectName("boldLabel")
        self.state_label = QLabel()
        self.progress_text_label = QLabel()
        self.progress_text_label.setObjectName("muted")
        top_row.addWidget(status_title)
        top_row.addWidget(self.state_label)
        top_row.addStretch(1)
        top_row.addWidget(self.progress_text_label)
        self.busy_bar = StripedProgressBar()
        status_layout.addLayout(top_row)
        status_layout.addWidget(self.busy_bar)
        layout.addWidget(status_card)
        self._set_state("ready")
        self._set_progress_text(0, 0)
        return card

    def _build_input_bar(self):
        card = self._make_card()
        layout = QHBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(10)
        label = QLabel("شماره سفارش:")
        label.setObjectName("boldLabel")
        self.order_input = QLineEdit()
        self.order_input.setLayoutDirection(Qt.LeftToRight)
        self.order_input.setPlaceholderText("شماره سفارش را وارد کنید...  (جداکننده: خط تیره)")
        self.process_button = QPushButton("▶  پردازش سفارش‌ها")
        self.process_button.setObjectName("primary")
        self.process_button.setFixedHeight(40)
        self.process_button.setMinimumWidth(170)
        self.settings_button = QPushButton("⚙")
        self.settings_button.setObjectName("iconButton")
        self.settings_button.setFixedSize(40, 40)
        self.settings_button.setToolTip("تنظیمات")
        layout.addWidget(label)
        layout.addWidget(self.order_input, 1)
        layout.addWidget(self.process_button)
        layout.addWidget(self.settings_button)
        return card

    def apply_stylesheet(self):
        """ Modern flat card style. Every rule starts with #content so the settings dialog is unaffected. """
        self.setStyleSheet("""
            #content { background-color: #eef3fb; }
            #content QLabel { background: transparent; color: #1e2a44; font-size: 10pt; }
            #content QLabel#cardTitle { font-size: 12pt; font-weight: bold; color: #14213d; }
            #content QLabel#boldLabel { font-weight: bold; }
            #content QLabel#muted { color: #7a879e; font-size: 9pt; }
            #content QLabel#badge {
                background: #e6eefc; color: #2c5bb8; border-radius: 10px;
                padding: 2px 12px; font-size: 9pt;
            }
            #content QFrame#card { background-color: #ffffff; border: 1px solid #dce5f3; border-radius: 14px; }
            #content QFrame#innerCard { background-color: #f5f8fe; border: 1px solid #e3eaf6; border-radius: 10px; }
            #content QLineEdit {
                background: #ffffff; border: 1px solid #d3dcec; border-radius: 10px;
                padding: 6px 12px; font-size: 10pt; selection-background-color: #2f7dff;
            }
            #content QLineEdit:focus { border: 1px solid #2f7dff; }
            #content QTextEdit#logBox {
                background: #fbfcfe; border: 1px solid #e3eaf6; border-radius: 10px;
                padding: 6px; font-size: 10pt;
            }
            #content QTableWidget { background: #ffffff; border: none; outline: none; font-size: 10pt; }
            #content QTableWidget::item { border-bottom: 1px solid #eef2f9; padding: 0px 6px; }
            #content QHeaderView::section {
                background: #f0f5fd; color: #52627f; border: none; padding: 0px 6px; font-size: 9pt;
            }
            #content QPushButton { font-size: 10pt; border-radius: 10px; padding: 6px 16px; border: none; }
            #content QPushButton#primary {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3f8dff, stop:1 #1e68eb);
                color: white; font-weight: bold;
            }
            #content QPushButton#primary:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5a9fff, stop:1 #2f78f5);
            }
            #content QPushButton#primary:pressed { background: #1a5bd4; }
            #content QPushButton#secondary { background: #ffffff; color: #2a3a5c; border: 1px solid #d3dcec; }
            #content QPushButton#secondary:hover { background: #f1f5fc; }
            #content QPushButton#iconButton {
                background: #ffffff; color: #2f7dff; border: 1px solid #d3dcec;
                font-size: 15pt; padding: 0px;
            }
            #content QPushButton#iconButton:hover { background: #f1f5fc; }
            #content QPushButton:disabled { background: #c9d3e5; color: #f3f6fb; border: none; }
            #content QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
            #content QScrollBar::handle:vertical { background: #c5d1e8; border-radius: 4px; min-height: 28px; }
            #content QScrollBar::handle:vertical:hover { background: #a9b9da; }
            #content QScrollBar::add-line:vertical, #content QScrollBar::sub-line:vertical { height: 0px; }
            #content QScrollBar::add-page:vertical, #content QScrollBar::sub-page:vertical { background: transparent; }
        """)

    # --------------------------------------------------------------------------
    # Small status helpers
    # --------------------------------------------------------------------------
    def _set_db_status(self, state):
        text, color = {
            "idle": ("برای دریافت فهرست، بروزرسانی را بزنید", "#8a96ad"),
            "loading": ("در حال دریافت از ERP ...", "#2f7dff"),
            "ok": ("متصل به دیتابیس", "#1e9e5a"),
            "error": ("اتصال به دیتابیس ناموفق بود", "#d64545"),
        }[state]
        self.db_status_label.setText(f'<span style="color:{color}">●&nbsp;&nbsp;{text}</span>')

    def _set_state(self, state):
        text, color = {
            "ready": ("آماده به کار", "#1e9e5a"),
            "running": ("در حال پردازش ...", "#2f7dff"),
        }[state]
        self.state_label.setText(f'<span style="color:{color}">●&nbsp;&nbsp;{text}</span>')

    def _set_progress_text(self, done, total):
        self.progress_text_label.setText(f"{fa_digits(done)} از {fa_digits(total)} سفارش پردازش شده")

    # --------------------------------------------------------------------------
    # Order-number box <-> table (they always show the same selection)
    # --------------------------------------------------------------------------
    @staticmethod
    def _parse_numbers(text):
        """ Order numbers found in free text: keeps the order, drops duplicates ("ORD-5368" -> 5368). """
        return list(dict.fromkeys(re.findall(r"\d+", text.translate(_DIGITS_TO_LATIN))))

    def _set_input_numbers(self, numbers):
        text = " - ".join(numbers)
        self.order_input.setText(text)
        self._last_input_text = text

    def _input_edited(self, text):
        """ Typing a separator after a number turns it into ' - ' (deleting is never touched). """
        if self._syncing:
            return
        grew = len(text) > len(self._last_input_text)
        self._last_input_text = text
        if grew and text and text[-1] in _INPUT_SEPARATORS and self.order_input.cursorPosition() == len(text):
            stripped = text.rstrip(_INPUT_SEPARATORS)
            if stripped and stripped[-1].isdigit():
                fixed = stripped + " - "
                if fixed != text:
                    self.order_input.setText(fixed)
                    self._last_input_text = fixed
        self._sync_table_from_input()

    def _normalize_input(self):
        """ Enter / leaving the box: numbers become 'a - b - c' without duplicates or stray characters. """
        self._set_input_numbers(self._parse_numbers(self.order_input.text()))
        self._sync_table_from_input()

    def _sync_table_from_input(self):
        wanted = set(self._parse_numbers(self.order_input.text()))
        self._syncing = True
        try:
            for row in range(self.orders_table.rowCount()):
                item = self.orders_table.item(row, 0)
                state = Qt.Checked if item.data(Qt.UserRole) in wanted else Qt.Unchecked
                if item.checkState() != state:
                    item.setCheckState(state)
        finally:
            self._syncing = False

    def _on_table_item_changed(self, item):
        if self._syncing or item.column() != 0:
            return
        number = item.data(Qt.UserRole)
        numbers = self._parse_numbers(self.order_input.text())
        if item.checkState() == Qt.Checked:
            if number not in numbers:
                numbers.append(number)
        elif number in numbers:
            numbers.remove(number)
        self._set_input_numbers(numbers)

    def _on_cell_clicked(self, row, column):
        """ Clicking anywhere on a row (not only the tick box) toggles it. """
        if column == 0:
            return
        item = self.orders_table.item(row, 0)
        item.setCheckState(Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked)

    def _toggle_all_visible(self, section):
        """ Clicking the '✓' header ticks every visible row, or clears them if all are already ticked. """
        if section != 0:
            return
        rows = [r for r in range(self.orders_table.rowCount()) if not self.orders_table.isRowHidden(r)]
        if not rows:
            return
        visible = [self.orders_table.item(r, 0).data(Qt.UserRole) for r in rows]
        numbers = self._parse_numbers(self.order_input.text())
        if all(number in numbers for number in visible):
            numbers = [number for number in numbers if number not in visible]
        else:
            numbers += [number for number in visible if number not in numbers]
        self._set_input_numbers(numbers)
        self._sync_table_from_input()

    def _filter_orders(self, query):
        query = query.strip().translate(_DIGITS_TO_LATIN).casefold()
        for row in range(self.orders_table.rowCount()):
            haystack = " ".join(self.orders_table.item(row, col).text() for col in range(1, 5)).casefold()
            self.orders_table.setRowHidden(row, bool(query) and query not in haystack)

    # --------------------------------------------------------------------------
    # Ready-orders list (ERP, read-only)
    # --------------------------------------------------------------------------
    def scan_order_directory(self):
        """Refresh ready-to-plan orders from SQL Server in read-only mode."""
        self.refresh_button.setDisabled(True)
        self._set_db_status("loading")
        QApplication.processEvents()
        try:
            orders, has_serial_timestamp = ErpRepository(CONFIG.settings).list_ready_orders()
            self._fill_orders_table(orders)
            self._set_db_status("ok")
            if not has_serial_timestamp:
                self.update_status("⚠️ در شِمای فعلی ERP تاریخ ایجاد سریال وجود ندارد؛ تاریخ تأیید نمایش‌داده نمی‌شود.")
            self.update_status(f"فهرست ERP بروزرسانی شد؛ {len(orders)} سفارش آماده یافت شد.")
        except Exception as error:
            self._set_db_status("error")
            self.update_status(f"خطا در دریافت فهرست سفارش‌ها: {error}")
            QMessageBox.warning(self, "خطای اتصال به ERP", str(error))
        finally:
            self.refresh_button.setDisabled(False)

    def _fill_orders_table(self, orders):
        table = self.orders_table
        self._syncing = True
        try:
            table.setRowCount(0)
            table.setRowCount(len(orders))
            for row, order in enumerate(orders):
                number = re.sub(r"^ORD-", "", str(order.get("order_number", "")).strip(), flags=re.IGNORECASE)
                check = QTableWidgetItem()
                check.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                check.setCheckState(Qt.Unchecked)
                check.setData(Qt.UserRole, number)
                serial_date = order.get("serial_approved_at")
                if hasattr(serial_date, "strftime"):
                    check.setToolTip(f"تأیید سریال: {jalali_date_text(serial_date)}")
                table.setItem(row, 0, check)
                values = [number, order.get("company_name") or "",
                          jalali_date_text(order.get("order_date")) or "-", str(order.get("item_count", 0))]
                for col, text in enumerate(values, start=1):
                    item = QTableWidgetItem(text)
                    item.setFlags(Qt.ItemIsEnabled)
                    if col == 2:
                        item.setTextAlignment(Qt.AlignLeading | Qt.AlignVCenter)
                    else:
                        item.setTextAlignment(Qt.AlignCenter)
                    if col == 1:
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)
                    table.setItem(row, col, item)
        finally:
            self._syncing = False
        self._sync_table_from_input()
        self._filter_orders(self.order_search.text())
        self.orders_badge.setText(f"{fa_digits(len(orders))} سفارش")

    # --------------------------------------------------------------------------
    # Processing
    # --------------------------------------------------------------------------
    def start_processing(self):
        """ Starts the worker thread to process orders. """
        if not self.process_button.isEnabled():
            return
        numbers = self._parse_numbers(self.order_input.text())
        if not numbers:
            QMessageBox.warning(self, "ورودی خالی", "لطفا حداقل یک شماره سفارش وارد کنید.")
            return
        self._set_input_numbers(numbers)
        self._sync_table_from_input()

        self.process_button.setDisabled(True)
        self.settings_button.setDisabled(True)
        self.refresh_button.setDisabled(True)
        self.status_box.clear()
        self.busy_bar.set_progress(0)
        self._set_progress_text(0, len(numbers))
        self._set_state("running")

        self.thread = QThread()
        self.worker = Worker("\n".join(numbers), CONFIG.settings)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.worker.status_update.connect(self.update_status)
        self.worker.error_signal.connect(self.show_error_message)
        self.worker.warning_signal.connect(self.show_warning_message)
        self.worker.info_signal.connect(self.show_info_message)
        self.worker.busy_changed.connect(self.set_busy)
        self.worker.progress_changed.connect(self.set_progress)

        self.thread.finished.connect(self._processing_finished)
        self.thread.start()

    def _processing_finished(self):
        self.busy_bar.stop()
        self._set_state("ready")
        self.process_button.setDisabled(False)
        self.settings_button.setDisabled(False)
        self.refresh_button.setDisabled(False)

    def set_busy(self, busy):
        """ Starts/stops the moving stripes of the status bar. """
        if busy:
            self.busy_bar.start()
        else:
            self.busy_bar.stop()

    def set_progress(self, done, total):
        """ 'x of y orders processed': fills the bar as orders are completed. """
        self.busy_bar.set_progress(done / total if total else 0)
        self._set_progress_text(done, total)

    # --------------------------------------------------------------------------
    # Window size / position
    # --------------------------------------------------------------------------
    def showEvent(self, event):
        """ On the first show, size and place the window (see fit_to_screen). """
        super().showEvent(event)
        if not getattr(self, "_screen_fitted", False):
            self._screen_fitted = True
            # Run after the window frame exists, so the real title-bar/border size is known.
            QTimer.singleShot(0, self.fit_to_screen)

    def fit_to_screen(self):
        """ Opens the window at the shortest height its boxes allow (see ORDER_TABLE_VISIBLE_ROWS and
        LOG_BOX_MIN_HEIGHT), never taller than WINDOW_MAX_HEIGHT_RATIO of the usable screen height,
        centered on screen. """
        screen = self.screen() if hasattr(self, "screen") else QApplication.primaryScreen()
        available = screen.availableGeometry()           # usable area: the taskbar is excluded
        frame_h = self.frameGeometry().height() - self.geometry().height()   # title bar + borders
        frame_w = self.frameGeometry().width() - self.geometry().width()
        if frame_h <= 0:   # window frame not measured yet: estimate title bar + borders
            frame_h = self.style().pixelMetric(QStyle.PM_TitleBarHeight) + 8
            frame_w = max(frame_w, 16)
        wanted_h = WINDOW_DEFAULT_HEIGHT or self.minimumSizeHint().height()
        total_h = min(wanted_h + frame_h, int(available.height() * WINDOW_MAX_HEIGHT_RATIO))
        total_w = min(self.width() + frame_w, available.width())
        self.resize(total_w - frame_w, total_h - frame_h)
        real_h = self.frameGeometry().height()           # Qt never goes below the minimum height
        self.move(available.x() + (available.width() - self.frameGeometry().width()) // 2,
                  max(available.y(), available.y() + (available.height() - real_h) // 2))

    # --------------------------------------------------------------------------
    # Report box and message boxes
    # --------------------------------------------------------------------------
    @staticmethod
    def _log_color(line):
        stripped = line.strip()
        if stripped.startswith("====="):
            return "#1f5fd8", True
        if any(mark in stripped for mark in ("✘", "❌", "❗", "خطا")):
            return "#d64545", False
        if "⚠" in stripped:
            return "#d98200", False
        if any(mark in stripped for mark in ("✔", "✅", "💯")):
            return "#1e9e5a", False
        if any(mark in stripped for mark in ("🌀", "⏳", "🔎")):
            return "#3b6fd4", False
        return "#2b3650", False

    def update_status(self, message):
        """ Adds a message to the report box: time stamp + color by message type (scrolls to the end). """
        lines = message.split("\n")
        while lines and not lines[-1].strip():
            lines.pop()
        stamp = time.strftime("%H:%M:%S")
        stamped = False
        for line in lines:
            if not line.strip():
                continue
            color, bold = self._log_color(line)
            text = html_lib.escape(line.strip("\r")).replace("  ", "&nbsp;&nbsp;")
            if bold:
                text = f"<b>{text}</b>"
            prefix = "" if stamped else f'<span style="color:#98a2b8; font-size:9pt">{stamp}</span>&nbsp;&nbsp;'
            stamped = True
            self.status_box.append(f'{prefix}<span style="color:{color}">{text}</span>')
        self.status_box.verticalScrollBar().setValue(self.status_box.verticalScrollBar().maximum())

    def show_error_message(self, title, message):
        """ Shows a critical error message box. """
        QMessageBox.critical(self, title, message)

    def show_warning_message(self, title, message):
        """ Shows a warning message box. """
        QMessageBox.warning(self, title, message)

    def show_info_message(self, title, message):
        """ Shows an informational message box. """
        QMessageBox.information(self, title, message)

    def show_settings(self):
        """ Opens the settings dialog. """
        dialog = SettingsDialog(CONFIG, self)
        if dialog.exec_() == QDialog.Accepted:
            self.scan_order_directory()


# ==============================================================================
# Main Execution Block
# ==============================================================================
def main():
    """ Main function to run the application. """
    app = QApplication(sys.argv)
    
    font_path = resource_path("IRAN.ttf")
    font_id = QFontDatabase.addApplicationFont(font_path)
    if font_id != -1:
        font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        app_font = QFont(font_family, 10)
        app.setFont(app_font)
    else:
        print("Warning: Font 'IRAN.ttf' could not be loaded. Using default font.")
        app.setFont(QFont("Tahoma", 10))

    splash_pix = QPixmap(resource_path("ProdPlanGenerator.png"))
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())
    progress = QProgressBar(splash)
    progress.setGeometry(
        90, splash_pix.height() - 100, splash_pix.width() - 180, 20
    )
    progress.setMaximum(100)
    progress.setValue(0)
    progress.setStyleSheet("""
        QProgressBar { 
            border: 1px solid grey; border-radius: 5px; 
            text-align: center; 
        }
        QProgressBar::chunk { background-color: #2e7dff; width: 1px; }
    """)
    splash.show()

    main_window = ProdPlanApp()
    
    counter = 0
    def update_progress():
        nonlocal counter
        counter += 1
        progress.setValue(counter)
        if counter >= 100:
            timer.stop()
            splash.close()
            main_window.show()

    timer = QTimer()
    timer.timeout.connect(update_progress)
    timer.start(25)

    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
