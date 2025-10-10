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
import shutil

# === Third-Party Libraries ===
import pandas as pd
import xlwings as xw
from PyPDF2 import PdfWriter
from PyQt5.QtCore import QObject, QThread, pyqtSignal, Qt, QTimer
from PyQt5.QtGui import QFont, QIcon, QPixmap, QFontDatabase
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTextEdit, QLabel, QMessageBox, QSplashScreen, QProgressBar, QStyle,
    QGroupBox, QDialog, QLineEdit, QFileDialog, QCheckBox, QRadioButton
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
            "order_file_path": "",
            "database_file_path": "",
            "output_base_path": "",
            "order_pdf_source_path": "",
            "delete_temp_files": True,
            "file_operation": "copy",
            "create_preparation_excel": True,
            "print_preparation_pdf": True,
            "print_timing_pdf": True
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
COL_QUANTITY = " تعداد"

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
# Core Application Logic (Worker Thread)
# ==============================================================================
class Worker(QObject):
    """ Handles the core data processing in a separate thread. """
    status_update = pyqtSignal(str)
    finished = pyqtSignal()
    error_signal = pyqtSignal(str, str)
    warning_signal = pyqtSignal(str, str)
    info_signal = pyqtSignal(str, str)

    def __init__(self, order_numbers_str, config_settings):
        super().__init__()
        self.order_numbers_str = order_numbers_str
        self.config = config_settings
        self.db_wb = None # To hold the workbook object
    
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
            self.status_update.emit(
                f"    ✔ چاپ {sheet.name} انجام شد:\n"
            )
            return True
        except Exception as e:
            self.status_update.emit(
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
            self.status_update.emit(
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
                self.status_update.emit(
                    f"    ✘ {suffix} اطلاعاتی برای پردازش ندارد.\n"
                    )
                continue

            print_this_pdf = True
            if suffix == "زمانسنجی" and not self.config.get('print_timing_pdf', True):
                print_this_pdf = False
                self.status_update.emit(
                    f"    - چاپ {suffix} بر اساس تنظیمات غیرفعال است.\n"
                )
            if suffix == "آماده سازی" and not self.config.get('print_preparation_pdf', True):
                print_this_pdf = False
                self.status_update.emit(
                    f"    - چاپ {suffix} بر اساس تنظیمات غیرفعال است.\n"
                )
            
            if print_this_pdf:
                pdf_filepath = os.path.join(order_folder, f"{product_code}_{suffix}.pdf")
                db_sheet.range(print_range).api.ExportAsFixedFormat(0, pdf_filepath)
                files_to_delete.append(pdf_filepath)
                if job_type == 'main': main_merger.append(pdf_filepath)
                elif job_type == 'preparation': preparation_merger.append(pdf_filepath)
                elif job_type == 'timing': timing_merger.append(pdf_filepath)
                self.status_update.emit(
                    f"    ✔ پرینت {suffix} برای {product_code} ذخیره شد.\n"
                )
            
            if suffix == "آماده سازی" and self.config.get('create_preparation_excel', True):
                try:
                    prep_data_range = db_sheet.range('T5:V24').options(ndim=2).value
                    for row_data in prep_data_range:
                        if row_data[0]:
                            new_row = {"شماره سفارش": order_num, "کد محصول": product_code, "شرح کالا": row_data[0], "تعداد": row_data[1], "اندازه برش": row_data[2]}
                            preparation_excel_data.append(new_row)
                    self.status_update.emit(
                        f"    ✔ داده‌های اکسل آماده‌سازی برای {product_code} استخراج شد.\n"
                    )
                except Exception as e:
                    self.status_update.emit(
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
            self.status_update.emit(
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
                self.status_update.emit(f"    ✔ نقشه فنی برای {product_code} کپی شد.\n")
            else:
                self.status_update.emit(f"    ✘ هشدار: نقشه فنی {os.path.basename(source_drawing_path)} یافت نشد.\n")

        return (main_merger, preparation_merger, timing_merger), files_to_delete, preparation_excel_data

    def run(self):
        """ Main processing logic. """
        try:
            order_file_path = os.path.normpath(self.config['order_file_path'])
            database_file_path = os.path.normpath(self.config['database_file_path'])
            output_base_path = os.path.normpath(self.config['output_base_path'])
            order_pdf_source_path = os.path.normpath(self.config['order_pdf_source_path'])
            
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

            if not all([order_file_path, database_file_path, output_base_path, order_pdf_source_path]):
                self.error_signal.emit(
                    "مسیرها تنظیم نشده",
                    "لطفا از بخش تنظیمات، تمام مسیرهای اصلی را مشخص کنید."
                )
                self.finished.emit()
                return

            order_numbers_list = [num.strip() for num in self.order_numbers_str.strip().split('\n') if num.strip()]
            if not order_numbers_list:
                self.error_signal.emit(
                    "ورودی خالی", 
                    "هیچ شماره سفارشی برای پردازش وارد نشده است."
                )
                self.finished.emit()
                return

            self.status_update.emit(
                f"شماره‌های سفارش برای پردازش:\n{order_numbers_list}\n"
            )
            self.status_update.emit(
                f"⏳ در حال خواندن فایل سفارش‌ها . . .\n"
            )
            df = pd.read_excel(order_file_path, sheet_name=ORDER_SHEET_NAME, engine='openpyxl')
            df.columns = df.columns.str.strip()
            # Ensure comparison is done between strings to avoid data type issues
            filtered_df = df[df[COL_ORDER_NUM.strip()].astype(str).isin(order_numbers_list)]

            if filtered_df.empty:
                self.warning_signal.emit(
                    "یافت نشد", "هیچ آیتمی مطابق با شماره سفارش‌های وارد شده در اکسل سفارش‌ها یافت نشد."
                )
                self.finished.emit()
                return
            self.status_update.emit(
                f"   🔎 تعداد {len(filtered_df)} آیتم برای پردازش یافت شد.\n"
            )

            self.status_update.emit(
                "   🌀 شروع پردازش اکسل دیتابیس محصولات و چاپ برگه‌ها . . .\n"
            )
            with xw.App(visible=False) as app:
                self.db_wb = app.books.open(database_file_path, read_only=True)
                db_sheet = self.db_wb.sheets[DATABASE_SHEET_NAME]

                for order_num, group in filtered_df.groupby(COL_ORDER_NUM.strip()):
                    order_num_str = str(order_num)
                    self.status_update.emit(
                        f"=======   شروع پردازش سفارش شماره {order_num_str}   =======\n"
                    )
                    order_folder = os.path.join(output_base_path, order_num_str)
                    os.makedirs(order_folder, exist_ok=True)
                    main_merger, preparation_merger, timing_merger = PdfWriter(), PdfWriter(), PdfWriter()
                    files_to_delete, original_order_filename = [], None
                    
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

                    try:
                        search_pattern = f"({order_num_str})"
                        for filename in os.listdir(order_pdf_source_path):
                            if search_pattern in filename and filename.lower().endswith('.pdf'):
                                source_filepath = os.path.join(order_pdf_source_path, filename)
                                dest_filepath = os.path.join(order_folder, filename)
                                
                                if self.config['file_operation'] == 'cut':
                                    shutil.move(source_filepath, dest_filepath)
                                    self.status_update.emit(
                                        f"  ✔ فایل اصلی سفارش {order_num_str} منتقل شد:\n"
                                    )
                                else:
                                    shutil.copy(source_filepath, dest_filepath)
                                    self.status_update.emit(
                                        f"  ✔ فایل اصلی سفارش {order_num_str} کپی شد:\n"
                                    )
                                
                                main_merger.append(dest_filepath)
                                files_to_delete.append(dest_filepath)
                                original_order_filename = filename
                                break
                        if not original_order_filename:
                            self.status_update.emit(
                                f"  - هشدار: فایل سفارش {order_num_str} یافت نشد.\n"
                            )
                    except Exception as e:
                        self.status_update.emit(
                            f"  - خطا در انتقال فایل اصلی سفارش: {e}\n"
                        )

                    for _, row in group.iterrows():
                        original_product_code = str(row[COL_PRODUCT_CODE.strip()])
                        quantity = row[COL_QUANTITY.strip()]
                        self.status_update.emit(
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
                            self.status_update.emit(
                                f"  ❗ هشدار: محصول {original_product_code} نامعتبر است. این آیتم نادیده گرفته شد.\n"
                            )
                            continue
                        
                        for final_code in valid_product_codes:
                            self.status_update.emit(
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
                                self.status_update.emit(
                                    f"  ✘ خطا در جستجوی قطعات جانبی: {e}\n"
                                )

                            if sub_components:
                                self.status_update.emit(
                                    f"  🔍 قطعات جانبی یافت شد: {', '.join(sub_components)}\n"
                                )
                                for sub_code in sub_components:
                                    self.status_update.emit(
                                        f"    🚀 شروع فرآیند چاپ برای {sub_code}\n"
                                    )
                                    mergers, files_to_delete, preparation_excel_data = self._process_product(
                                        sub_code, order_num_str, quantity, order_folder, db_sheet, mergers,
                                        files_to_delete, preparation_excel_data, technical_drawing_paths
                                    )
                                    main_merger, preparation_merger, timing_merger = mergers
                    
                    final_main_pdf_path = None
                    if len(main_merger.pages) > 0:
                        clean_name = order_num_str 
                        if original_order_filename: 
                            base_name = os.path.splitext(original_order_filename)[0]
                            clean_name = re.sub(r'\s*ok$', '', base_name, flags=re.IGNORECASE).strip()
                        final_main_pdf_path = os.path.join(order_folder, f"{clean_name}.pdf")
                        with open(final_main_pdf_path, "wb") as f: main_merger.write(f)
                        self.status_update.emit(f"  ✔ فایل اصلی ادغام شده برای سفارش {order_num_str} ذخیره شد.\n")

                    if len(preparation_merger.pages) > 0:
                        with open(os.path.join(order_folder, f"آماده سازی({order_num_str}).pdf"), "wb") as f: preparation_merger.write(f)
                        self.status_update.emit(f"  ✔ فایل آماده سازی ادغام شده برای سفارش {order_num_str} ذخیره شد.\n")
                    
                    if len(timing_merger.pages) > 0:
                        with open(os.path.join(order_folder, f"زمانسنجی({order_num_str}).pdf"), "wb") as f: timing_merger.write(f)
                        self.status_update.emit(f"  ✔ فایل زمانسنجی ادغام شده برای سفارش {order_num_str} ذخیره شد.\n")

                    if self.config.get('create_preparation_excel', False) and preparation_excel_data:
                        try:
                            prep_excel_path = os.path.join(order_folder, f"آماده سازی({order_num_str}).xlsx")
                            df_prep = pd.DataFrame(preparation_excel_data)
                            df_prep.insert(0, 'ردیف', range(1, len(df_prep) + 1))
                            df_prep = df_prep.reindex(columns=preparation_excel_headers)
                            df_prep.to_excel(prep_excel_path, index=False, engine='openpyxl')
                            self.status_update.emit(f"  ✔ فایل اکسل آماده سازی برای سفارش {order_num_str} ذخیره شد.\n")
                        except Exception as e:
                            self.status_update.emit(f"  ✘ خطا در ذخیره فایل اکسل آماده سازی: {e}\n")
                    
                    main_merger.close(); preparation_merger.close(); timing_merger.close()

                    if self.config['delete_temp_files'] and files_to_delete:
                        self.status_update.emit(
                            f"\n  ⏳ شروع پاکسازی فایل‌های موقت برای سفارش {order_num_str} . . .\n"
                        )
                        if final_main_pdf_path and final_main_pdf_path in files_to_delete:
                            files_to_delete.remove(final_main_pdf_path)
                            self.status_update.emit(
                                f"    - فایل نهایی {os.path.basename(final_main_pdf_path)} از لیست حذف خارج شد.\n"
                            )
                        deleted_count = 0
                        for file_path in files_to_delete:
                            try:
                                if os.path.exists(file_path): os.remove(file_path); deleted_count += 1
                            except Exception as e:
                                self.status_update.emit(
                                    f"    ❗ خطا در حذف فایل {os.path.basename(file_path)}: {e}\n"
                                )
                        self.status_update.emit(
                            f"    ✔ {deleted_count} فایل موقت با موفقیت حذف شد.\n"
                        )
                
                self.db_wb.close()
                self.db_wb = None
                self.status_update.emit(
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
            self.status_update.emit(
                f"خطا در یافتن فایل: {e}\n"
            )
        except Exception as e:
            self.error_signal.emit(
                "خطای کلی", f"یک خطای ناشناخته در برنامه رخ داد:\n{e}"
            )
            self.status_update.emit(
                f"خطای بحرانی: {e}\n"
            )
        finally:
            if self.db_wb:
                self.db_wb.close()
            self.finished.emit()

# ==============================================================================
# Settings Dialog Window
# ==============================================================================
class SettingsDialog(QDialog):
    """ Dialog window for application settings. """
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.setWindowTitle("ویژگی‌ها و تنظیمات")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setMinimumWidth(700)

        self.layout = QVBoxLayout(self)

        paths_group = QGroupBox("تنظیمات مسیرها")
        paths_layout = QVBoxLayout()
        self.path_widgets = {
            "order_file_path": self._create_path_selector("فایل اکسل سفارش‌ها:", "file"),
            "database_file_path": self._create_path_selector("فایل اکسل دیتابیس:", "file"),
            "order_pdf_source_path": self._create_path_selector("پوشه فایل‌های سفارش:", "folder"),
            "output_base_path": self._create_path_selector("پوشه ذخیره خروجی:", "folder"),
        }
        for _, widget_tuple in self.path_widgets.items():
            paths_layout.addLayout(widget_tuple[0])
        paths_group.setLayout(paths_layout)
        self.layout.addWidget(paths_group)

        options_group = QGroupBox("گزینه‌های پردازش")
        options_layout = QVBoxLayout()
        
        self.print_prep_pdf_checkbox = QCheckBox("آماده‌سازی چاپ شود PDF فایل")
        self.print_timing_pdf_checkbox = QCheckBox("زمانسنجی چاپ شود PDF فایل")
        self.create_prep_excel_checkbox = QCheckBox("فایل اکسل آماده‌سازی ایجاد شود")
        self.delete_temp_checkbox = QCheckBox("پس از پایان پردازش، فایل‌های موقت پاک شوند")
        
        op_layout = QHBoxLayout()
        op_label = QLabel("عملیات انتقال فایل اصلی سفارش:")
        self.copy_radio = QRadioButton("Copy")
        self.cut_radio = QRadioButton("Cut")
        op_layout.addWidget(op_label)
        op_layout.addWidget(self.copy_radio)
        op_layout.addWidget(self.cut_radio)
        op_layout.addStretch()

        options_layout.addWidget(self.print_prep_pdf_checkbox)
        options_layout.addWidget(self.print_timing_pdf_checkbox)
        options_layout.addWidget(self.create_prep_excel_checkbox)
        options_layout.addWidget(self.delete_temp_checkbox)
        options_layout.addLayout(op_layout)
        options_group.setLayout(options_layout)
        self.layout.addWidget(options_group)

        bottom_buttons_layout = QHBoxLayout()
        self.about_button = QPushButton("درباره برنامه")
        self.about_button.setObjectName("secondary")
        self.save_button = QPushButton("ذخیره تنظیمات")
        
        bottom_buttons_layout.addWidget(self.about_button)
        bottom_buttons_layout.addStretch()
        bottom_buttons_layout.addWidget(self.save_button)
        self.layout.addLayout(bottom_buttons_layout)

        self._connect_signals()
        self._populate_fields()
        self.setStyleSheet(parent.styleSheet())

    def _create_path_selector(self, label_text, selection_mode):
        """ Creates a layout for path selection with a label, line edit, and browse button. """
        layout = QHBoxLayout()
        label = QLabel(label_text)
        label.setFixedWidth(150)
        line_edit = QLineEdit()

        browse_button = QPushButton("انتخاب مسیر ⤷")
        browse_button.setObjectName("actionButton")
        browse_button.setFixedWidth(90)
        
        layout.addWidget(label)
        layout.addWidget(line_edit)
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
        
        self.print_prep_pdf_checkbox.setChecked(settings.get("print_preparation_pdf", True))
        self.print_timing_pdf_checkbox.setChecked(settings.get("print_timing_pdf", True))
        self.create_prep_excel_checkbox.setChecked(settings.get("create_preparation_excel", True))
        self.delete_temp_checkbox.setChecked(settings.get("delete_temp_files", True))
        
        if settings.get("file_operation", "copy") == "cut":
            self.cut_radio.setChecked(True)
        else:
            self.copy_radio.setChecked(True)

    def _connect_signals(self):
        """ Connects widget signals to their respective slots. """
        self.save_button.clicked.connect(self.accept)
        self.about_button.clicked.connect(self._show_about)

    def _save_settings(self):
        """ Saves the current settings from the dialog to the config manager. """
        for key, widget_tuple in self.path_widgets.items():
            self.config_manager.settings[key] = widget_tuple[1].text()
        
        self.config_manager.settings["print_preparation_pdf"] = self.print_prep_pdf_checkbox.isChecked()
        self.config_manager.settings["print_timing_pdf"] = self.print_timing_pdf_checkbox.isChecked()
        self.config_manager.settings["create_preparation_excel"] = self.create_prep_excel_checkbox.isChecked()
        self.config_manager.settings["delete_temp_files"] = self.delete_temp_checkbox.isChecked()
        self.config_manager.settings["file_operation"] = "cut" if self.cut_radio.isChecked() else "copy"
        
        self.config_manager.save()

    def accept(self):
        """ Overrides the default accept to save settings before closing. """
        self._save_settings()
        QMessageBox.information(self, "ذخیره شد", "تنظیمات با موفقیت ذخیره شد.")
        super().accept()

    def _show_about(self):
        """ Shows the 'About' dialog. """
        dlg = QDialog(self)
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
class ProdPlanApp(QWidget):
    """ Main application window (GUI). """
    def __init__(self):
        super().__init__()
        self.worker = None
        self.thread = None
        self.scanner_worker = None
        self.scanner_thread = None

        self.initUI()

    def initUI(self):
        """ Initializes the UI components. """
        self.setWindowTitle("سازنده برگه‌های سفارش - ProdPlanGenerator - FardanApex")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.setGeometry(250, 100, 900, 600)

        main_layout = QVBoxLayout(self)
        top_layout = QHBoxLayout()
        bottom_layout = QHBoxLayout()
        
        # --- Left Pane ---
        left_pane_layout = QVBoxLayout()
        status_group_box = QGroupBox("لیست سفارش‌ها")
        status_group_box_layout = QVBoxLayout()
        self.refresh_button = QPushButton("بروزرسانی لیست")
        self.refresh_button.setObjectName("actionButton")
        refresh_icon = self.style().standardIcon(QStyle.SP_BrowserReload)
        self.refresh_button.setIcon(refresh_icon)
        
        self.confirmed_title_label = QLabel("سفارش‌های تایید شده:")
        self.confirmed_orders_label = QLabel("برای مشاهده لیست، روی بروزرسانی کلیک کنید.")
        self.confirmed_orders_label.setObjectName("confirmedOrders")
        self.confirmed_orders_label.setWordWrap(True)
        
        self.pending_title_label = QLabel("سفارش‌های در انتظار تایید:")
        self.pending_orders_label = QLabel("برای مشاهده لیست، روی بروزرسانی کلیک کنید.")
        self.pending_orders_label.setObjectName("pendingOrders")
        self.pending_orders_label.setWordWrap(True)
        
        status_group_box_layout.addWidget(self.confirmed_title_label)
        status_group_box_layout.addWidget(self.confirmed_orders_label)
        status_group_box_layout.addSpacing(10)
        status_group_box_layout.addWidget(self.pending_title_label)
        status_group_box_layout.addWidget(self.pending_orders_label)
        status_group_box_layout.addWidget(self.refresh_button)
        status_group_box.setLayout(status_group_box_layout)
        status_group_box.setFixedHeight(status_group_box.sizeHint().height() + 20)
        
        input_group_box = QGroupBox("ورود شماره سفارش‌ها")
        input_group_box_layout = QVBoxLayout()
        self.order_input = QTextEdit()
        self.order_input.setPlaceholderText("هر شماره سفارش در یک خط جدید...")
        self.process_button = QPushButton("پردازش سفارش‌ها")
        self.process_button.setObjectName("actionButton")
        process_icon = self.style().standardIcon(QStyle.SP_DialogApplyButton)
        self.process_button.setIcon(process_icon)
        input_group_box_layout.addWidget(self.order_input)
        input_group_box_layout.addWidget(self.process_button)
        input_group_box.setLayout(input_group_box_layout)
        
        left_pane_layout.addWidget(status_group_box)
        left_pane_layout.addWidget(input_group_box)
        
        # --- Right Pane ---
        right_pane_layout = QVBoxLayout()
        processing_status_group_box = QGroupBox("گزارش وضعیت پردازش")
        processing_status_layout = QVBoxLayout()
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        processing_status_layout.addWidget(self.status_box)
        processing_status_group_box.setLayout(processing_status_layout)
        right_pane_layout.addWidget(processing_status_group_box)
        
        top_layout.addLayout(right_pane_layout, 65)
        top_layout.addLayout(left_pane_layout, 35)
        
        # --- Bottom Bar ---
        self.settings_button = QPushButton("تنظیمات 🛠️")
        self.settings_button.setFixedHeight(45)
        bottom_layout.addWidget(self.settings_button)
        bottom_layout.addStretch()
        
        # --- Final Layout ---
        main_layout.addLayout(top_layout)
        main_layout.addLayout(bottom_layout)
        self.setLayout(main_layout)
        
        self.process_button.clicked.connect(self.start_processing)
        self.settings_button.clicked.connect(self.show_settings)
        self.refresh_button.clicked.connect(self.scan_order_directory)
        
        self.apply_stylesheet()

    def scan_order_directory(self):
        """ Scans the source directory in a background thread. """
        path = CONFIG.settings.get('order_pdf_source_path')
        if not path or not os.path.isdir(path):
            msg = "مسیر پوشه سفارش‌ها تنظیم نشده یا نامعتبر است."
            self.update_status(f"راهنما: {msg}")
            self.confirmed_orders_label.setText("-")
            self.pending_orders_label.setText("-")
            return

        self.refresh_button.setDisabled(True)
        self.confirmed_orders_label.setText("در حال اسکن پوشه . . .")
        self.pending_orders_label.setText("در حال اسکن پوشه . . .")

        self.scanner_thread = QThread()
        self.scanner_worker = DirectoryScannerWorker(path)
        self.scanner_worker.moveToThread(self.scanner_thread)

        self.scanner_thread.started.connect(self.scanner_worker.run)
        self.scanner_worker.scan_complete.connect(self.update_order_lists)
        
        self.scanner_worker.finished.connect(self.scanner_thread.quit)
        self.scanner_worker.finished.connect(self.scanner_worker.deleteLater)
        self.scanner_thread.finished.connect(self.scanner_thread.deleteLater)
        self.scanner_thread.finished.connect(lambda: self.refresh_button.setDisabled(False))
        
        self.scanner_thread.start()

    def update_order_lists(self, confirmed_orders, pending_orders):
        """ Receives the results from the scanner worker and updates the UI. """
        self.confirmed_orders_label.setText(" - ".join(sorted(confirmed_orders)) if confirmed_orders else "سفارش تایید شده‌ای یافت نشد")
        self.pending_orders_label.setText(" - ".join(sorted(pending_orders)) if pending_orders else "سفارش تایید نشده‌ای یافت نشد")
        self.update_status("لیست سفارش‌ها بروزرسانی شد.")

    def apply_stylesheet(self):
        """ Applies a custom stylesheet to the application. """
        self.setStyleSheet("""
            QWidget { background-color: #f5f7fb; }
            QLabel { font-size: 10pt; color: #333; }
            QTextEdit { 
                background-color: white; border: 1px solid #d0d7df; 
                border-radius: 6px; padding: 6px; font-size: 10pt; 
            }
            QGroupBox { 
                border: 1px solid #d0d7df; border-radius: 6px; 
                margin-top: 10px; padding: 10px; 
            }
            QGroupBox::title { 
                subcontrol-origin: margin; subcontrol-position: top center; 
                padding: 0 5px; 
            }
            QLabel#confirmedOrders { color: #28a745; font-size: 10pt; }
            QLabel#pendingOrders { color: #dc3545; font-size: 10pt; }
            QPushButton { 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5aa9ff, stop:1 #2e7dff); 
                color: white; border: none; padding: 8px 10px; border-radius: 8px; 
            }
            QPushButton:hover { 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6bb8ff, stop:1 #3b8bff); 
            }
            QPushButton#secondary { 
                background: #eef4ff; color: #1a3b6e; border: 1px solid #d0dbff; 
            }
            QPushButton#secondary:hover { background: #e0e9ff; }
            QPushButton#actionButton { 
                background-color: #f0f0f0; color: #333; border: 1px solid #ccc; 
                text-align: Center; padding: 5px; font-size: 9pt; 
            }
            QPushButton#actionButton:hover { background-color: #e9e9e9; border-color: #bbb; }
            QPushButton:disabled { background-color: #bdc3c7; color: #7f8c8d; }
        """)

    def start_processing(self):
        """ Starts the worker thread to process orders. """
        order_numbers = self.order_input.toPlainText()
        if not order_numbers.strip():
            QMessageBox.warning(self, "ورودی خالی", "لطفا حداقل یک شماره سفارش وارد کنید.")
            return
        
        self.process_button.setDisabled(True)
        self.settings_button.setDisabled(True)
        self.status_box.clear()
        
        self.thread = QThread()
        self.worker = Worker(order_numbers, CONFIG.settings)
        self.worker.moveToThread(self.thread)
        
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        self.worker.status_update.connect(self.update_status)
        self.worker.error_signal.connect(self.show_error_message)
        self.worker.warning_signal.connect(self.show_warning_message)
        self.worker.info_signal.connect(self.show_info_message)
        
        self.thread.finished.connect(lambda: self.process_button.setDisabled(False))
        self.thread.finished.connect(lambda: self.settings_button.setDisabled(False))
        
        self.thread.start()

    def update_status(self, message):
        """ Appends a message to the status box. """
        self.status_box.append(message)
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



