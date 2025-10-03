"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Fardan Apex --- ProdPlanGenerator ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
This software automates processing factory orders by generating, organizing,
and exporting production documents into structured PDF files.

Author: Behnam Rabieyan
Company: Garma Gostar Fardan
Created: 2025
"""

# Standard library imports
import json
import os
import re
import shutil
import sys

# Third-party imports
import pandas as pd
import xlwings as xw
from PyPDF2 import PdfWriter
from PyQt5.QtCore import QObject, QThread, pyqtSignal, Qt, QTimer
from PyQt5.QtGui import QFont, QIcon, QPixmap, QFontDatabase
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit,
    QLabel, QMessageBox, QSplashScreen, QProgressBar, QStyle, QGroupBox,
    QDialog, QLineEdit, QFileDialog, QCheckBox, QRadioButton
)


# ==============================================================================
# Helper Function for File Paths
# ==============================================================================

def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller.

    Args:
        relative_path (str): The relative path to the resource file.

    Returns:
        str: The absolute path to the resource file.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# ==============================================================================
# Configuration Management
# ==============================================================================

class ConfigManager:
    """Manages loading and saving of application settings from a JSON file."""

    def __init__(self, config_file='config.json'):
        """
        Initializes the ConfigManager.

        Args:
            config_file (str): The name of the configuration file.
        """
        self.config_file = resource_path(config_file)
        self.settings = {}
        self.load()

    def load(self):
        """Loads settings from the JSON file."""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self._create_default_config()

    def save(self):
        """Saves the current settings to the JSON file."""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.settings, f, indent=4, ensure_ascii=False)

    def _create_default_config(self):
        """Creates a default configuration file if one doesn't exist."""
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
    "print_range": "B2:Y54", "cell_product": "K5", "cell_order": "W5",
    "cell_flag": "Z47", "check_cell": "P49"
}
ST_CONFIG = {
    "print_range": "B2:Y68", "cell_product": "G7",
    "cell_flag": "Z63", "check_cell": "P64"
}
KL_CONFIG = {
    "print_range": "B1:L41", "cell_product": "E2"
}


# ==============================================================================
# Core Application Logic (Worker Thread)
# ==============================================================================

class Worker(QObject):
    """Handles the core data processing in a separate thread."""
    status_update = pyqtSignal(str)
    finished = pyqtSignal()
    error_signal = pyqtSignal(str, str)
    warning_signal = pyqtSignal(str, str)
    info_signal = pyqtSignal(str, str)

    def __init__(self, order_numbers_str, config_settings):
        super().__init__()
        self.order_numbers_str = order_numbers_str
        self.config = config_settings
        self.db_wb = None  # To hold the workbook object

    def find_last_numeric_row(self, sheet, search_range):
        """Finds the last row with a numeric value in a given range."""
        values = sheet.range(search_range).options(ndim=1).value
        start_row = sheet.range(search_range).row
        for i in range(len(values) - 1, -1, -1):
            if isinstance(values[i], (int, float)) and values[i] is not None:
                return start_row + i
        return 0

    def print_conditional_sheet(self, sheet, product_code, pdf_filepath, config, order_num=None):
        """Prints a conditional sheet to PDF after setting necessary values."""
        try:
            sheet.range(config['cell_product']).value = product_code
            if order_num and 'cell_order' in config:
                sheet.range(config['cell_order']).value = order_num
            if 'check_cell' in config and 'cell_flag' in config:
                check_val = str(sheet.range(config['check_cell']).value).strip().upper()
                sheet.range(config['cell_flag']).value = (check_val == 'FALSE')
            sheet.range(config['print_range']).api.ExportAsFixedFormat(0, pdf_filepath)
            self.status_update.emit(
                f"    ✔ چاپ {sheet.name} انجام شد:\n        - نام فایل: {os.path.basename(pdf_filepath)}\n")
            return True
        except Exception as e:
            self.status_update.emit(f"    ✘ خطا در چاپ {sheet.name}: {e}\n")
            return False

    def _process_product(
            self, product_code, order_num, quantity, order_folder, db_sheet,
            mergers, files_to_delete, preparation_excel_data, technical_drawing_paths
    ):
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
                f"  ❗ هشدار: کد محصول {product_code} در دیتابیس نامعتبر است. این آیتم نادیده گرفته شد.\n")
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
                if db_sheet.range('P9').value is None:
                    has_data = False
                else:
                    last_row = self.find_last_numeric_row(db_sheet, 'N9:N47')
                    print_range = f"N4:Q{last_row}" if last_row > 0 else ""
            elif suffix == "آماده سازی":
                if db_sheet.range('U5').value is None:
                    has_data = False
                else:
                    last_row = self.find_last_numeric_row(db_sheet, 'S5:S24')
                    print_range = f"S1:Y{last_row}" if last_row > 0 else ""

            if not print_range:
                has_data = False

            if not has_data:
                self.status_update.emit(f"    ✘ {suffix} اطلاعاتی برای پردازش ندارد.\n")
                continue

            # Check settings before printing optional PDFs
            print_this_pdf = True
            if suffix == "زمانسنجی" and not self.config.get('print_timing_pdf', True):
                print_this_pdf = False
                self.status_update.emit(f"    - چاپ PDF {suffix} بر اساس تنظیمات غیرفعال است.\n")
            if suffix == "آماده سازی" and not self.config.get('print_preparation_pdf', True):
                print_this_pdf = False
                self.status_update.emit(f"    - چاپ PDF {suffix} بر اساس تنظیمات غیرفعال است.\n")

            # Print the PDF if enabled and data exists
            if print_this_pdf:
                pdf_filepath = os.path.join(order_folder, f"{product_code}_{suffix}.pdf")
                db_sheet.range(print_range).api.ExportAsFixedFormat(0, pdf_filepath)
                files_to_delete.append(pdf_filepath)
                if job_type == 'main':
                    main_merger.append(pdf_filepath)
                elif job_type == 'preparation':
                    preparation_merger.append(pdf_filepath)
                elif job_type == 'timing':
                    timing_merger.append(pdf_filepath)
                self.status_update.emit(f"    ✔ فایل PDF {suffix} برای {product_code} ذخیره شد.\n")

            # Extract data for preparation Excel file if enabled
            if suffix == "آماده سازی" and self.config.get('create_preparation_excel', True):
                try:
                    prep_data_range = db_sheet.range('T5:V24').options(ndim=2).value
                    for row_data in prep_data_range:
                        if row_data[0]:  # Check if the item description exists
                            new_row = {
                                "شماره سفارش": order_num,
                                "کد محصول": product_code,
                                "شرح کالا": row_data[0],
                                "تعداد": row_data[1],
                                "اندازه برش": row_data[2]
                            }
                            preparation_excel_data.append(new_row)
                    self.status_update.emit(f"    ✔ داده‌های اکسل آماده‌سازی برای {product_code} استخراج شد.\n")
                except Exception as e:
                    self.status_update.emit(f"    ✘ خطا در استخراج داده‌های اکسل آماده‌سازی: {e}\n")

        # --- Process conditional sheets ---
        process_code = str(db_sheet.range(CONDITIONAL_CHECK_CELL).value)[:2].upper()
        if process_code == 'MF':
            mf_sheet = self.db_wb.sheets[MF_SHEET_NAME]
            pdf_filepath = os.path.join(order_folder, f"{product_code}_{MF_SHEET_NAME}.pdf")
            if self.print_conditional_sheet(mf_sheet, product_code, pdf_filepath, MF_CONFIG, order_num=order_num):
                main_merger.append(pdf_filepath)
                files_to_delete.append(pdf_filepath)
        if process_code in ('DS', 'DF', 'NL', 'DL'):
            st_sheet = self.db_wb.sheets[ST_SHEET_NAME]
            pdf_filepath = os.path.join(order_folder, f"{product_code}_{ST_SHEET_NAME}.pdf")
            if self.print_conditional_sheet(st_sheet, product_code, pdf_filepath, ST_CONFIG):
                main_merger.append(pdf_filepath)
                files_to_delete.append(pdf_filepath)
        if process_code in ('NL', 'DL'):
            kl_sheet = self.db_wb.sheets[KL_SHEET_NAME]
            pdf_filepath = os.path.join(order_folder, f"{product_code}_{KL_SHEET_NAME}.pdf")
            if self.print_conditional_sheet(kl_sheet, product_code, pdf_filepath, KL_CONFIG):
                main_merger.append(pdf_filepath)
                files_to_delete.append(pdf_filepath)

        # --- Copy technical drawing ---
        drawing_source_folder = technical_drawing_paths.get(process_code)
        if drawing_source_folder:
            search_code = product_code
            if '-' in product_code:  # For TF and TS codes
                search_code = product_code.split('-')[0]

            source_drawing_path = os.path.join(drawing_source_folder, f"{search_code[:6]}.pdf")
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
        """Main processing logic."""
        try:
            # Normalize paths from config
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

            # Validate essential paths
            if not all([order_file_path, database_file_path, output_base_path, order_pdf_source_path]):
                self.error_signal.emit("مسیرها تنظیم نشده", "لطفا از بخش تنظیمات، تمام مسیرهای اصلی را مشخص کنید.")
                return

            # Read and filter order numbers
            order_numbers_list = [num.strip() for num in self.order_numbers_str.strip().split('\n') if num.strip()]
            if not order_numbers_list:
                self.error_signal.emit("ورودی خالی", "هیچ شماره سفارشی برای پردازش وارد نشده است.")
                return
            order_numbers_int = [int(num) for num in order_numbers_list]

            self.status_update.emit(f"شماره‌های سفارش برای پردازش:\n{order_numbers_list}\n")
            self.status_update.emit("⏳ در حال خواندن فایل سفارش‌ها...\n")
            df = pd.read_excel(order_file_path, sheet_name=ORDER_SHEET_NAME, engine='openpyxl')
            df.columns = df.columns.str.strip()
            filtered_df = df[df[COL_ORDER_NUM.strip()].isin(order_numbers_int)]

            if filtered_df.empty:
                self.warning_signal.emit("یافت نشد", "هیچ آیتمی مطابق با شماره سفارش‌های وارد شده یافت نشد.")
                return
            self.status_update.emit(f"   🔎 تعداد {len(filtered_df)} آیتم برای پردازش یافت شد.\n")

            self.status_update.emit("   🌀 شروع پردازش اکسل دیتابیس و چاپ برگه‌ها . . .\n")
            with xw.App(visible=False) as app:
                self.db_wb = app.books.open(database_file_path, read_only=True)
                db_sheet = self.db_wb.sheets[DATABASE_SHEET_NAME]

                for order_num, group in filtered_df.groupby(COL_ORDER_NUM.strip()):
                    self.status_update.emit(f"=====   شروع پردازش سفارش شماره: {order_num}   =====\n")
                    order_folder = os.path.join(output_base_path, str(order_num))
                    os.makedirs(order_folder, exist_ok=True)

                    # Initialize mergers and lists for this order
                    main_merger, prep_merger, time_merger = PdfWriter(), PdfWriter(), PdfWriter()
                    files_to_delete, original_order_filename = [], None
                    prep_excel_data = []
                    prep_excel_headers = [
                        "ردیف", "شماره سفارش", "کد محصول", "شرح کالا", "تعداد",
                        "اندازه برش", "تاریخ نیاز", "توضیحات", "امضای تحویل گیرنده"
                    ]

                    # Find and copy/move the main order PDF
                    try:
                        search_pattern = f"({order_num})"
                        for filename in os.listdir(order_pdf_source_path):
                            if search_pattern in filename and filename.lower().endswith('.pdf'):
                                src_path = os.path.join(order_pdf_source_path, filename)
                                dst_path = os.path.join(order_folder, filename)

                                if self.config['file_operation'] == 'cut':
                                    shutil.move(src_path, dst_path)
                                    self.status_update.emit(f"  ✔ فایل اصلی سفارش {order_num} منتقل شد.\n")
                                else:
                                    shutil.copy(src_path, dst_path)
                                    self.status_update.emit(f"  ✔ فایل اصلی سفارش {order_num} کپی شد.\n")

                                main_merger.append(dst_path)
                                files_to_delete.append(dst_path)
                                original_order_filename = filename
                                break
                        if not original_order_filename:
                            self.status_update.emit(f"  - هشدار: فایل سفارش {order_num} یافت نشد.\n")
                    except Exception as e:
                        self.status_update.emit(f"  - خطا در انتقال فایل اصلی سفارش: {e}\n")

                    # Process each product in the order group
                    for _, row in group.iterrows():
                        orig_code = str(row[COL_PRODUCT_CODE.strip()])
                        quantity = row[COL_QUANTITY.strip()]
                        self.status_update.emit(f"✨ بررسی محصول اصلی کد {orig_code}\n")

                        # Find valid product codes (main product or variants)
                        valid_codes = []
                        db_sheet.range(CELL_PRODUCT_CODE).value = orig_code
                        if str(db_sheet.range(CELL_CHECK).value).strip().lower() != 'empty':
                            valid_codes.append(orig_code)
                        else:
                            for suffix in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                                variant_code = f"{orig_code}{suffix}"
                                db_sheet.range(CELL_PRODUCT_CODE).value = variant_code
                                if str(db_sheet.range(CELL_CHECK).value).strip().lower() != 'empty':
                                    valid_codes.append(variant_code)
                                else:
                                    break

                        if not valid_codes:
                            self.status_update.emit(
                                f"  ❗ هشدار: محصول {orig_code} نامعتبر است. نادیده گرفته شد.\n")
                            continue

                        for final_code in valid_codes:
                            self.status_update.emit(f"  🚀 شروع فرآیند چاپ برای محصول: {final_code}\n")

                            mergers = (main_merger, prep_merger, time_merger)
                            mergers, files_to_delete, prep_excel_data = self._process_product(
                                final_code, order_num, quantity, order_folder, db_sheet, mergers,
                                files_to_delete, prep_excel_data, technical_drawing_paths
                            )
                            main_merger, prep_merger, time_merger = mergers

                            # Find and process sub-components (TF and TS)
                            sub_components = []
                            try:
                                bom_range = db_sheet.range('C5:C64').options(ndim=1).value
                                for cell_value in bom_range:
                                    if isinstance(cell_value, str):
                                        found = re.findall(r'(TS-\d+|TF-\d+)', cell_value, re.IGNORECASE)
                                        if found:
                                            sub_components.extend(found)
                                sub_components = sorted(list(set(sub_components)))
                            except Exception as e:
                                self.status_update.emit(f"  ✘ خطا در جستجوی قطعات جانبی: {e}\n")

                            if sub_components:
                                self.status_update.emit(f"  🔍 قطعات جانبی یافت شد: {', '.join(sub_components)}\n")
                                for sub_code in sub_components:
                                    self.status_update.emit(f"    🚀 چاپ برای قطعه جانبی: {sub_code}\n")
                                    mergers = (main_merger, prep_merger, time_merger)
                                    mergers, files_to_delete, prep_excel_data = self._process_product(
                                        sub_code, order_num, quantity, order_folder, db_sheet, mergers,
                                        files_to_delete, prep_excel_data, technical_drawing_paths
                                    )
                                    main_merger, prep_merger, time_merger = mergers

                    # --- Finalize and save merged files for the current order ---
                    final_main_pdf_path = None
                    if len(main_merger.pages) > 0:
                        base_name = os.path.splitext(original_order_filename)[0] if original_order_filename else str(order_num)
                        clean_name = re.sub(r'\s*ok$', '', base_name, flags=re.IGNORECASE).strip()
                        final_main_pdf_path = os.path.join(order_folder, f"{clean_name}.pdf")
                        with open(final_main_pdf_path, "wb") as f:
                            main_merger.write(f)
                        self.status_update.emit(f"  ✔ فایل اصلی ادغام شده برای سفارش {order_num} ذخیره شد.\n")

                    if len(prep_merger.pages) > 0:
                        path = os.path.join(order_folder, f"آماده سازی({order_num}).pdf")
                        with open(path, "wb") as f:
                            prep_merger.write(f)
                        self.status_update.emit(f"  ✔ فایل آماده سازی ادغام شده ذخیره شد.\n")

                    if len(time_merger.pages) > 0:
                        path = os.path.join(order_folder, f"زمانسنجی({order_num}).pdf")
                        with open(path, "wb") as f:
                            time_merger.write(f)
                        self.status_update.emit(f"  ✔ فایل زمانسنجی ادغام شده ذخیره شد.\n")

                    if self.config.get('create_preparation_excel', False) and prep_excel_data:
                        try:
                            path = os.path.join(order_folder, f"آماده سازی({order_num}).xlsx")
                            df_prep = pd.DataFrame(prep_excel_data)
                            df_prep.insert(0, 'ردیف', range(1, len(df_prep) + 1))
                            df_prep = df_prep.reindex(columns=prep_excel_headers)
                            df_prep.to_excel(path, index=False, engine='openpyxl')
                            self.status_update.emit(f"  ✔ فایل اکسل آماده سازی ذخیره شد.\n")
                        except Exception as e:
                            self.status_update.emit(f"  ✘ خطا در ذخیره اکسل آماده سازی: {e}\n")

                    main_merger.close()
                    prep_merger.close()
                    time_merger.close()

                    # --- Clean up temporary files ---
                    if self.config['delete_temp_files'] and files_to_delete:
                        self.status_update.emit(f"\n  ⏳ شروع پاکسازی فایل‌های موقت...\n")
                        if final_main_pdf_path and final_main_pdf_path in files_to_delete:
                            files_to_delete.remove(final_main_pdf_path)
                            self.status_update.emit(
                                f"    - فایل نهایی از لیست حذف خارج شد.\n")
                        deleted_count = 0
                        for file_path in files_to_delete:
                            try:
                                if os.path.exists(file_path):
                                    os.remove(file_path)
                                    deleted_count += 1
                            except Exception as e:
                                self.status_update.emit(f"    ❗ خطا در حذف فایل {os.path.basename(file_path)}: {e}\n")
                        self.status_update.emit(f"    ✔ {deleted_count} فایل موقت حذف شد.\n")

                self.db_wb.close()
                self.db_wb = None
                self.status_update.emit("\n💯 عملیات پردازش با موفقیت به پایان رسید.\n")
                self.info_signal.emit("اتمام عملیات", "تمام سفارش‌ها با موفقیت پردازش شدند.")
        except FileNotFoundError as e:
            msg = (f"فایل یا مسیر یافت نشد:\n{e.filename}\n\n"
                   "لطفا مسیرها را در تنظیمات بررسی کنید.")
            self.error_signal.emit("خطای فایل", msg)
            self.status_update.emit(f"خطای FileNotFoundError: {e}\n")
        except Exception as e:
            self.error_signal.emit("خطای کلی", f"یک خطای ناشناخته رخ داد:\n{e}")
            self.status_update.emit(f"خطای بحرانی: {e}\n")
        finally:
            if self.db_wb:
                self.db_wb.close()
            self.finished.emit()


# ==============================================================================
# Settings Dialog Window
# ==============================================================================

class SettingsDialog(QDialog):
    """Dialog window for application settings."""

    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.setWindowTitle("ویژگی‌ها و تنظیمات")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setMinimumWidth(600)

        self.layout = QVBoxLayout(self)
        self._create_path_settings_group()
        self._create_options_group()
        self._create_bottom_buttons()

        self._connect_signals()
        self._populate_fields()
        self.setStyleSheet(parent.styleSheet())

    def _create_path_settings_group(self):
        """Creates the group box for path settings."""
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

    def _create_options_group(self):
        """Creates the group box for processing options."""
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

    def _create_bottom_buttons(self):
        """Creates the 'About' and 'Save' buttons at the bottom."""
        bottom_buttons_layout = QHBoxLayout()
        self.about_button = QPushButton("درباره برنامه")
        self.about_button.setObjectName("secondary")
        self.save_button = QPushButton("ذخیره تنظیمات")

        bottom_buttons_layout.addWidget(self.about_button)
        bottom_buttons_layout.addStretch()
        bottom_buttons_layout.addWidget(self.save_button)
        self.layout.addLayout(bottom_buttons_layout)

    def _create_path_selector(self, label_text, selection_mode):
        """ Creates a layout for path selection. """
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
        """Opens a file dialog to select an Excel file."""
        path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل 📄", "", "Excel Files (*.xlsx *.xlsm)")
        if path:
            line_edit.setText(path)

    def _browse_folder(self, line_edit):
        """Opens a dialog to select a folder."""
        path = QFileDialog.getExistingDirectory(self, "انتخاب پوشه 📁")
        if path:
            line_edit.setText(path)

    def _populate_fields(self):
        """Fills the settings fields with values from the config manager."""
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
        """Connects widget signals to their respective slots."""
        self.save_button.clicked.connect(self.accept)
        self.about_button.clicked.connect(self._show_about_dialog)

    def _save_settings(self):
        """Saves the current settings from the dialog to the config manager."""
        for key, widget_tuple in self.path_widgets.items():
            self.config_manager.settings[key] = widget_tuple[1].text()

        self.config_manager.settings["print_preparation_pdf"] = self.print_prep_pdf_checkbox.isChecked()
        self.config_manager.settings["print_timing_pdf"] = self.print_timing_pdf_checkbox.isChecked()
        self.config_manager.settings["create_preparation_excel"] = self.create_prep_excel_checkbox.isChecked()
        self.config_manager.settings["delete_temp_files"] = self.delete_temp_checkbox.isChecked()
        self.config_manager.settings["file_operation"] = "cut" if self.cut_radio.isChecked() else "copy"

        self.config_manager.save()

    def accept(self):
        """Overrides the default accept to save settings before closing."""
        self._save_settings()
        QMessageBox.information(self, "ذخیره شد", "تنظیمات با موفقیت ذخیره شد.")
        super().accept()

    def _show_about_dialog(self):
        """Shows the 'About' dialog."""
        dlg = QDialog(self)
        dlg.setWindowTitle("درباره برنامه")
        dlg.setFixedSize(500, 450)
        main_layout = QVBoxLayout(dlg)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # --- Intro Section ---
        intro_layout = QVBoxLayout()
        intro_layout.setSpacing(0)
        intro_layout.setContentsMargins(0, 0, 0, 0)
        lbl_intro_text = (
            "<h3><b>Fardan Apex — ProdPlanGenerator</b></h3>"
            "<h4>Order PDF Generator Application</h4><br>"
            "Automates the generation of production order documents "
            "from Excel data into consolidated PDFs.<br>"
            "Version: 1.2 — © 2025 All Rights Reserved<br>"
            "Developed exclusively for:<br>"
            "Garma Gostar Fardan Co."
        )
        lbl_intro = QLabel(lbl_intro_text)
        lbl_intro.setWordWrap(True)
        lbl_intro.setAlignment(Qt.AlignLeft)
        intro_layout.addWidget(lbl_intro)

        logo = QLabel()
        logo_pix = QPixmap(resource_path("FardanLogo.jpg"))
        if logo_pix.isNull():
            logo_pix = QPixmap(resource_path("FardanLogoEN.png"))
        if not logo_pix.isNull():
            logo.setPixmap(logo_pix.scaledToWidth(175, Qt.SmoothTransformation))
        else:
            logo.setText("Fardan Apex")
        logo.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        logo.setContentsMargins(35, 10, 0, 0)
        intro_layout.addWidget(logo)
        main_layout.addLayout(intro_layout)

        # --- Developer Section ---
        dev_layout = QVBoxLayout()
        dev_layout.setSpacing(0)
        dev_layout.setContentsMargins(5, 0, 0, 5)
        font_id = QFontDatabase.addApplicationFont(resource_path("BrittanySignature.ttf"))
        font_family = "Sans Serif"
        if font_id != -1:
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]

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
    """Main application window (GUI)."""

    def __init__(self):
        super().__init__()
        self.worker_thread = None
        self.init_ui()
        self.scan_order_directory()

    def init_ui(self):
        """Initializes the UI components."""
        self.setWindowTitle("سازنده برگه‌های سفارش - ProdPlanGenerator - FardanApex")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
        self.setGeometry(250, 100, 900, 600)

        main_layout = QVBoxLayout(self)
        top_layout = QHBoxLayout()
        bottom_layout = QHBoxLayout()

        # --- Left Pane (Inputs and Status) ---
        left_pane = self._create_left_pane()
        # --- Right Pane (Log Output) ---
        right_pane = self._create_right_pane()
        top_layout.addLayout(right_pane, 65)
        top_layout.addLayout(left_pane, 35)

        # --- Bottom Bar (Settings) ---
        self.settings_button = QPushButton("تنظیمات 🛠️")
        self.settings_button.setFixedHeight(45)
        bottom_layout.addWidget(self.settings_button)
        bottom_layout.addStretch()

        # --- Final Layout ---
        main_layout.addLayout(top_layout)
        main_layout.addLayout(bottom_layout)
        self.setLayout(main_layout)

        # --- Connect Signals ---
        self.process_button.clicked.connect(self.start_processing)
        self.settings_button.clicked.connect(self.show_settings)
        self.refresh_button.clicked.connect(self.scan_order_directory)

        self.apply_stylesheet()

    def _create_left_pane(self):
        """Creates the left pane of the main window."""
        layout = QVBoxLayout()
        status_group_box = QGroupBox("لیست سفارش‌ها")
        status_group_layout = QVBoxLayout()
        self.refresh_button = QPushButton("بروزرسانی لیست")
        self.refresh_button.setObjectName("actionButton")
        self.refresh_button.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))

        self.confirmed_orders_label = QLabel("...")
        self.confirmed_orders_label.setObjectName("confirmedOrders")
        self.confirmed_orders_label.setWordWrap(True)
        self.pending_orders_label = QLabel("...")
        self.pending_orders_label.setObjectName("pendingOrders")
        self.pending_orders_label.setWordWrap(True)

        status_group_layout.addWidget(QLabel("سفارش‌های تایید شده:"))
        status_group_layout.addWidget(self.confirmed_orders_label)
        status_group_layout.addSpacing(10)
        status_group_layout.addWidget(QLabel("سفارش‌های در انتظار تایید:"))
        status_group_layout.addWidget(self.pending_orders_label)
        status_group_layout.addWidget(self.refresh_button)
        status_group_box.setLayout(status_group_layout)
        status_group_box.setFixedHeight(status_group_box.sizeHint().height() + 20)

        input_group_box = QGroupBox("ورود شماره سفارش‌ها")
        input_group_layout = QVBoxLayout()
        self.order_input = QTextEdit()
        self.order_input.setPlaceholderText("هر شماره سفارش در یک خط جدید...")
        self.process_button = QPushButton("پردازش سفارش‌ها")
        self.process_button.setObjectName("actionButton")
        self.process_button.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        input_group_layout.addWidget(self.order_input)
        input_group_layout.addWidget(self.process_button)
        input_group_box.setLayout(input_group_layout)

        layout.addWidget(status_group_box)
        layout.addWidget(input_group_box)
        return layout

    def _create_right_pane(self):
        """Creates the right pane of the main window."""
        layout = QVBoxLayout()
        status_group_box = QGroupBox("گزارش وضعیت پردازش")
        status_group_layout = QVBoxLayout()
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        status_group_layout.addWidget(self.status_box)
        status_group_box.setLayout(status_group_layout)
        layout.addWidget(status_group_box)
        return layout

    def scan_order_directory(self):
        """Scans the source directory for confirmed and pending order PDFs."""
        confirmed, pending = [], []
        path = CONFIG.settings.get('order_pdf_source_path')

        if not path or not os.path.isdir(path):
            msg = "مسیر پوشه سفارش‌ها تنظیم نشده یا نامعتبر است."
            self.update_status(f"راهنما: {msg}")
            self.confirmed_orders_label.setText("-")
            self.pending_orders_label.setText("-")
            return

        for filename in os.listdir(path):
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

        self.confirmed_orders_label.setText(" - ".join(sorted(confirmed)) or "سفارش تایید شده‌ای یافت نشد.")
        self.pending_orders_label.setText(" - ".join(sorted(pending)) or "سفارش در انتظار تاییدی یافت نشد.")
        self.update_status("لیست سفارش‌ها از پوشه بروزرسانی شد.")

    def apply_stylesheet(self):
        """Applies a custom stylesheet to the application."""
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
        """Starts the worker thread to process orders."""
        order_numbers = self.order_input.toPlainText()
        if not order_numbers.strip():
            QMessageBox.warning(self, "ورودی خالی", "لطفا حداقل یک شماره سفارش وارد کنید.")
            return

        self.process_button.setDisabled(True)
        self.settings_button.setDisabled(True)
        self.status_box.clear()

        self.worker_thread = QThread()
        worker = Worker(order_numbers, CONFIG.settings)
        worker.moveToThread(self.worker_thread)

        # Connect worker signals
        self.worker_thread.started.connect(worker.run)
        worker.finished.connect(self.worker_thread.quit)
        worker.finished.connect(worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)

        worker.status_update.connect(self.update_status)
        worker.error_signal.connect(self.show_error_message)
        worker.warning_signal.connect(self.show_warning_message)
        worker.info_signal.connect(self.show_info_message)

        # Re-enable buttons on finish
        self.worker_thread.finished.connect(lambda: self.process_button.setDisabled(False))
        self.worker_thread.finished.connect(lambda: self.settings_button.setDisabled(False))

        self.worker_thread.start()

    def update_status(self, message):
        """Appends a message to the status box."""
        self.status_box.append(message)
        self.status_box.verticalScrollBar().setValue(self.status_box.verticalScrollBar().maximum())

    def show_error_message(self, title, message):
        """Shows a critical error message box."""
        QMessageBox.critical(self, title, message)

    def show_warning_message(self, title, message):
        """Shows a warning message box."""
        QMessageBox.warning(self, title, message)

    def show_info_message(self, title, message):
        """Shows an informational message box."""
        QMessageBox.information(self, title, message)

    def show_settings(self):
        """Opens the settings dialog."""
        dialog = SettingsDialog(CONFIG, self)
        if dialog.exec_() == QDialog.Accepted:
            self.scan_order_directory()


# ==============================================================================
# Main Execution Block
# ==============================================================================

def main():
    """Main function to run the application."""
    app = QApplication(sys.argv)

    # --- Font Setup ---
    font_path = resource_path("IRAN.ttf")
    font_id = QFontDatabase.addApplicationFont(font_path)
    if font_id != -1:
        font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        app.setFont(QFont(font_family, 10))
    else:
        print("Warning: Font 'IRAN.ttf' could not be loaded. Using default font.")
        app.setFont(QFont("Tahoma", 10))

    # --- Splash Screen ---
    splash_pix = QPixmap(resource_path("ProdPlanGenerator.png"))
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())
    progress = QProgressBar(splash)
    progress.setGeometry(90, splash_pix.height() - 100, splash_pix.width() - 180, 20)
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

    # --- Animate Splash Screen Progress ---
    counter = 0
    timer = QTimer()

    def update_progress():
        nonlocal counter
        counter += 1
        progress.setValue(counter)
        if counter >= 100:
            timer.stop()
            splash.close()
            main_window.show()

    timer.timeout.connect(update_progress)
    timer.start(25)  # Update progress every 25ms

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()


