"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Fardan Apex --- ProdPlanGenerator ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
This software automates processing factory orders by generating, organizing,
and exporting production documents into structured PDF files.

Author: Behnam Rabieyan
Company: Garma Gostar Fardan
Created: 2025
"""

import sys
import os
import shutil
import re
import pandas as pd
import xlwings as xw
from PyPDF2 import PdfWriter
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTextEdit, QLabel, QMessageBox, QFrame)
from PyQt5.QtCore import QObject, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon

# ==============================================================================
# بخش تنظیمات: تمام مسیرها، نام‌ها و آدرس‌ها از اینجا قابل ویرایش هستند
# ==============================================================================
# --- دیکشنری فایل‌های اصلی ---
ORDER_FILE_PATH = r"D:\MyWork\G.G.Fardan\row\order.xlsx"
DATABASE_FILE_PATH = r"D:\MyWork\G.G.Fardan\row\Database.xlsm"
OUTPUT_BASE_PATH = r"D:\MyWork\G.G.Fardan\row"
ORDER_PDF_SOURCE_PATH = r"D:\MyWork\G.G.Fardan"

# --- دیکشنری مسیر نقشه‌های فنی ---
TECHNICAL_DRAWING_PATHS = {
    "TS": r"\\fileserver\mohandesi\PDF Plan\ترموسوئیچ",
    "TF": r"\\fileserver\mohandesi\PDF Plan\ترموفیوز\نقشه های معتبر",
    "DS": r"\\fileserver\mohandesi\PDF Plan\هیتر سیمی\نقشه های معتبر",
    "DF": r"\\fileserver\mohandesi\PDF Plan\فویلی\نقشه های معتبر",
    "NL": r"\\fileserver\mohandesi\PDF Plan\لوله ای\نقشه های معتبر",
    "DL": r"\\fileserver\mohandesi\PDF Plan\لوله ای\نقشه های معتبر",
    "MF": r"\\fileserver\mohandesi\PDF Plan\میله ای\نقشه های معتبر",
    "MR": r"\\fileserver\mohandesi\PDF Plan\میله ای\نقشه های معتبر"
}

# --- نام شیت‌ و ستون‌های اکسل سفارش ---
ORDER_SHEET_NAME = "OrderList"
COL_ORDER_NUM = "شماره سفارش"
COL_PRODUCT_CODE = "کد محصول"
COL_QUANTITY = "تعداد"

# --- نام شیت و سلول‌های برگه اصلی ---
DATABASE_SHEET_NAME = "LOM"
CELL_PRODUCT_CODE = "I4"
CELL_CHECK = "D1"
CELL_QUANTITY = "J6"
CELL_ORDER_NUM_DB = "W3"

# --- لیست کارهای چاپ از شیت اصلی (LOM) ---
LOM_PRINT_JOBS = [
    {"suffix": "LOM", "type": "main"},
    {"suffix": "زمانسنجی", "type": "timing"},
    {"suffix": "آماده سازی", "type": "preparation"}
]

# --- پارامتر کنترل شرطی ---
CONDITIONAL_CHECK_CELL = 'D3'

# --- برگه‌های شرطی ---
MF_SHEET_NAME = "برنامه فنرپیچ"
ST_SHEET_NAME = "برنامه سیم تابنده"
KL_SHEET_NAME = "برنامه خم لوله‌ای"

# --- تنظیمات چاپ برگه‌های شرطی ---
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
# بخش منطق اصلی برنامه (بدون تغییر)
# ==============================================================================
class Worker(QObject):
    """
    این کلاس مسئول اجرای فرآیندهای سنگین در یک ترد جداگانه است
    تا از قفل شدن رابط کاربری جلوگیری شود.
    """
    status_update = pyqtSignal(str)
    finished = pyqtSignal()
    error_signal = pyqtSignal(str, str)
    warning_signal = pyqtSignal(str, str)
    info_signal = pyqtSignal(str, str)


    def __init__(self, order_numbers_str):
        super().__init__()
        self.order_numbers_str = order_numbers_str


    def find_last_numeric_row(self, sheet, search_range):
        """آخرین ردیف حاوی مقدار عددی را در یک محدوده مشخص پیدا می‌کند."""
        values = sheet.range(search_range).options(ndim=1).value
        start_row = sheet.range(search_range).row
        for i in range(len(values) - 1, -1, -1):
            if isinstance(values[i], (int, float)) and values[i] is not None:
                return start_row + i
        return 0


    def print_conditional_sheet(self, sheet, product_code, pdf_filepath, config, order_num=None):
        """یک برگه شرطی را بر اساس تنظیمات داده شده، پردازش و چاپ می‌کند."""
        try:
            sheet.range(config['cell_product']).value = product_code
            if order_num and 'cell_order' in config:
                sheet.range(config['cell_order']).value = order_num
            if 'check_cell' in config and 'cell_flag' in config:
                check_val = str(sheet.range(config['check_cell']).value).strip().upper()
                sheet.range(config['cell_flag']).value = (check_val == 'FALSE')
            sheet.range(config['print_range']).api.ExportAsFixedFormat(0, pdf_filepath)
            self.status_update.emit(f"    - چاپ شرطی ({sheet.name}): {os.path.basename(pdf_filepath)}\n")
            return True
        except Exception as e:
            self.status_update.emit(f"    - خطا در چاپ شیت '{sheet.name}': {e}\n")
            return False


    def run(self):
        """تابع اصلی پردازش که تمام منطق برنامه در آن قرار دارد."""
        try:
            order_numbers_list = [num.strip() for num in self.order_numbers_str.strip().split('\n') if num.strip()]
            if not order_numbers_list:
                self.error_signal.emit("ورودی خالی", "هیچ شماره سفارشی برای پردازش وارد نشده است.")
                self.finished.emit()
                return
            order_numbers_int = [int(num) for num in order_numbers_list]

            self.status_update.emit(f"شماره‌های سفارش برای پردازش: {order_numbers_list}\n")
            self.status_update.emit(f"در حال خواندن فایل سفارش‌ها: {ORDER_FILE_PATH}...\n")
            df = pd.read_excel(ORDER_FILE_PATH, sheet_name=ORDER_SHEET_NAME, engine='openpyxl')
            df.columns = df.columns.str.strip()
            filtered_df = df[df[COL_ORDER_NUM.strip()].isin(order_numbers_int)]

            if filtered_df.empty:
                self.warning_signal.emit("یافت نشد", "هیچ آیتمی مطابق با شماره سفارش‌های وارد شده در اکسل سفارش‌ها یافت نشد.")
                self.finished.emit()
                return
            self.status_update.emit(f"تعداد {len(filtered_df)} آیتم برای پردازش یافت شد.\n")

            self.status_update.emit("شروع پردازش اکسل دیتابیس و چاپ برگه‌ها...\n")
            with xw.App(visible=False) as app:
                db_wb = app.books.open(DATABASE_FILE_PATH, read_only=True)
                db_sheet = db_wb.sheets[DATABASE_SHEET_NAME]

                for order_num, group in filtered_df.groupby(COL_ORDER_NUM.strip()):
                    self.status_update.emit(f"\n===== شروع پردازش سفارش شماره: {order_num} =====\n")
                    order_folder = os.path.join(OUTPUT_BASE_PATH, str(order_num))
                    os.makedirs(order_folder, exist_ok=True)

                    main_merger, preparation_merger, timing_merger = PdfWriter(), PdfWriter(), PdfWriter()
                    files_to_delete = []
                    original_order_filename = None

                    try:
                        search_pattern = f"({order_num})"
                        for filename in os.listdir(ORDER_PDF_SOURCE_PATH):
                            if search_pattern in filename and filename.lower().endswith('.pdf'):
                                source_filepath = os.path.join(ORDER_PDF_SOURCE_PATH, filename)
                                dest_filepath = os.path.join(order_folder, filename)
                                shutil.copy(source_filepath, dest_filepath)
                                main_merger.append(dest_filepath)
                                files_to_delete.append(dest_filepath)
                                original_order_filename = filename
                                self.status_update.emit(f"  - فایل اصلی سفارش '{filename}' کپی و به لیست اصلی اضافه شد.\n")
                                break
                        if not original_order_filename:
                            self.status_update.emit(f"  - هشدار: فایل PDF اصلی برای سفارش {order_num} یافت نشد.\n")
                    except Exception as e:
                        self.status_update.emit(f"  - خطا در کپی فایل اصلی سفارش: {e}\n")

                    for _, row in group.iterrows():
                        original_product_code = str(row[COL_PRODUCT_CODE.strip()])
                        quantity = row[COL_QUANTITY.strip()]
                        self.status_update.emit(f"-> بررسی محصول: {original_product_code}\n")
                        
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
                            self.status_update.emit(f"  - هشدار: محصول {original_product_code} نامعتبر است. این آیتم نادیده گرفته شد.\n")
                            continue

                        for final_code in valid_product_codes:
                            self.status_update.emit(f"  * شروع فرآیند چاپ برای محصول: {final_code}\n")
                            db_sheet.range(CELL_ORDER_NUM_DB).value = order_num
                            db_sheet.range(CELL_QUANTITY).value = quantity
                            db_sheet.range(CELL_PRODUCT_CODE).value = final_code

                            # مرحله ۱: چاپ برگه‌های LOM، آماده‌سازی و زمانسنجی
                            for job in LOM_PRINT_JOBS:
                                suffix, job_type = job['suffix'], job['type']
                                should_print, print_range = True, ""
                                if suffix == "LOM":
                                    last_row = self.find_last_numeric_row(db_sheet, 'B5:B65'); print_range = f"B1:G{last_row}" if last_row > 0 else ""
                                elif suffix == "زمانسنجی":
                                    if db_sheet.range('P9').value is None: should_print = False
                                    else: last_row = self.find_last_numeric_row(db_sheet, 'N9:N47'); print_range = f"N4:Q{last_row}" if last_row > 0 else ""
                                elif suffix == "آماده سازی":
                                    if db_sheet.range('U5').value is None: should_print = False
                                    else: last_row = self.find_last_numeric_row(db_sheet, 'S5:S24'); print_range = f"S1:Y{last_row}" if last_row > 0 else ""
                               
                                if not print_range: should_print = False

                                if should_print:
                                    pdf_filepath = os.path.join(order_folder, f"{final_code}_{suffix}.pdf")
                                    db_sheet.range(print_range).api.ExportAsFixedFormat(0, pdf_filepath)
                                    files_to_delete.append(pdf_filepath)
                                    if job_type == 'main': main_merger.append(pdf_filepath)
                                    elif job_type == 'preparation': preparation_merger.append(pdf_filepath)
                                    elif job_type == 'timing': timing_merger.append(pdf_filepath)
                                    self.status_update.emit(f"    - PDF ({suffix}) ذخیره و به لیست مربوطه اضافه شد.\n")
                                else:
                                    self.status_update.emit(f"    - چاپ '{suffix}' به دلیل نبود اطلاعات لغو شد.\n")

                            # مرحله ۲: چاپ برگه‌های شرطی
                            process_code = str(db_sheet.range(CONDITIONAL_CHECK_CELL).value)[:2].upper()
                            if process_code == 'MF':
                                mf_sheet = db_wb.sheets[MF_SHEET_NAME]
                                pdf_filepath = os.path.join(order_folder, f"{final_code}_{MF_SHEET_NAME}.pdf")
                                if self.print_conditional_sheet(mf_sheet, final_code, pdf_filepath, MF_CONFIG, order_num=order_num):
                                    main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)
                            if process_code in ('DS', 'DF', 'NL', 'DL'):
                                st_sheet = db_wb.sheets[ST_SHEET_NAME]
                                pdf_filepath = os.path.join(order_folder, f"{final_code}_{ST_SHEET_NAME}.pdf")
                                if self.print_conditional_sheet(st_sheet, final_code, pdf_filepath, ST_CONFIG):
                                    main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)
                            if process_code in ('NL', 'DL'):
                                kl_sheet = db_wb.sheets[KL_SHEET_NAME]
                                pdf_filepath = os.path.join(order_folder, f"{final_code}_{KL_SHEET_NAME}.pdf")
                                if self.print_conditional_sheet(kl_sheet, final_code, pdf_filepath, KL_CONFIG):
                                    main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)

                            # مرحله ۳: کپی نقشه فنی
                            drawing_source_folder = TECHNICAL_DRAWING_PATHS.get(process_code)
                            if drawing_source_folder:
                                source_drawing_path = os.path.join(drawing_source_folder, f"{final_code[:6]}.pdf")
                                dest_drawing_path = os.path.join(order_folder, f"{final_code}_نقشه.pdf")
                                if os.path.exists(source_drawing_path):
                                    shutil.copy(source_drawing_path, dest_drawing_path)
                                    main_merger.append(dest_drawing_path)
                                    files_to_delete.append(dest_drawing_path)
                                    self.status_update.emit(f"    - نقشه فنی کپی و به لیست اصلی اضافه شد.\n")
                                else:
                                    self.status_update.emit(f"    - هشدار: نقشه فنی '{os.path.basename(source_drawing_path)}' یافت نشد.\n")

                    # ذخیره و پاکسازی نهایی برای سفارش فعلی
                    if len(main_merger.pages) > 0:
                        clean_name = str(order_num)
                        if original_order_filename: clean_name = re.sub(r'\s*ok', '', original_order_filename, flags=re.IGNORECASE).replace('.pdf', '').strip()
                        output_path = os.path.join(order_folder, f"{clean_name}.pdf")
                        with open(output_path, "wb") as f: main_merger.write(f)
                        self.status_update.emit(f"  -> فایل اصلی ادغام شده برای سفارش {order_num} ذخیره شد.\n")
                    if len(preparation_merger.pages) > 0:
                        with open(os.path.join(order_folder, f"آماده سازی({order_num}).pdf"), "wb") as f: preparation_merger.write(f)
                        self.status_update.emit(f"  -> فایل 'آماده سازی' ادغام شده برای سفارش {order_num} ذخیره شد.\n")
                    if len(timing_merger.pages) > 0:
                        with open(os.path.join(order_folder, f"زمانسنجی({order_num}).pdf"), "wb") as f: timing_merger.write(f)
                        self.status_update.emit(f"  -> فایل 'زمانسنجی' ادغام شده برای سفارش {order_num} ذخیره شد.\n")
                   
                    main_merger.close(); preparation_merger.close(); timing_merger.close()
                    
                    if files_to_delete:
                        self.status_update.emit(f"  -> شروع پاکسازی فایل‌های موقت برای سفارش {order_num}...\n")
                        deleted_count = 0
                        for file_path in files_to_delete:
                            try:
                                if os.path.exists(file_path): os.remove(file_path); deleted_count += 1
                            except Exception as e:
                                self.status_update.emit(f"    - خطا در حذف فایل {os.path.basename(file_path)}: {e}\n")
                        self.status_update.emit(f"    - {deleted_count} فایل موقت با موفقیت حذف شد.\n")

                db_wb.close()
                self.status_update.emit("\nعملیات پردازش، ساخت، ادغام و پاکسازی با موفقیت به پایان رسید.\n")
                self.info_signal.emit("اتمام عملیات", "تمام سفارش‌ها با موفقیت پردازش شدند.")

        except Exception as e:
            self.error_signal.emit("خطای کلی", f"یک خطای ناشناخته در برنامه رخ داد:\n{e}")
            self.status_update.emit(f"خطای بحرانی: {e}\n")
        finally:
            self.finished.emit()


# ==============================================================================
# بخش رابط کاربری (GUI) با PyQt5
# ==============================================================================
class ProdPlanApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()


    def initUI(self):
        self.setWindowTitle("پیشرفته از روی لیست سفارش PDF سازنده و ادغام کننده")
        self.setGeometry(300, 300, 850, 600)
        self.setFont(QFont("Tahoma", 10))

        # Layouts
        main_layout = QVBoxLayout(self)
        top_layout = QHBoxLayout() # For input and status boxes
        bottom_layout = QHBoxLayout() # For buttons

        # --- Top Section ---
        # Input Box (Right)
        input_v_box = QVBoxLayout()
        self.input_label = QLabel("شماره های سفارش را وارد کنید:")
        self.order_input = QTextEdit()
        self.order_input.setPlaceholderText("هر شماره سفارش در یک خط جدید...")
        input_v_box.addWidget(self.input_label)
        input_v_box.addWidget(self.order_input)

        # Status Box (Left)
        status_v_box = QVBoxLayout()
        self.status_label = QLabel("وضعیت پردازش:")
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        status_v_box.addWidget(self.status_label)
        status_v_box.addWidget(self.status_box)

        # Add input and status boxes to top layout
        top_layout.addLayout(status_v_box, 65) # 65% width
        top_layout.addLayout(input_v_box, 35) # 35% width

        # --- Bottom Section ---
        self.process_button = QPushButton("آغاز پردازش")
        self.process_button.setFixedHeight(45)
        self.settings_button = QPushButton("ویژگی‌ها")
        self.settings_button.setFixedHeight(45)
        
        # Add buttons to bottom layout
        bottom_layout.addWidget(self.settings_button)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.process_button)

        # --- Main Layout ---
        main_layout.addLayout(top_layout)
        main_layout.addLayout(bottom_layout)
        self.setLayout(main_layout)

        # Connections
        self.process_button.clicked.connect(self.start_processing)
        self.settings_button.clicked.connect(self.show_settings)

        # Apply Stylesheet
        self.apply_stylesheet()


    def apply_stylesheet(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f0f0;
            }
            QLabel {
                font-size: 11pt;
                font-weight: bold;
                color: #333;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #ccc;
                border-radius: 8px;
                padding: 5px;
                font-size: 10pt;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                border-radius: 8px;
                padding: 10px 25px;
                border: none;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
                color: #ecf0f1;
            }
        """)


    def start_processing(self):
        order_numbers = self.order_input.toPlainText()
        if not order_numbers.strip():
            QMessageBox.warning(self, "ورودی خالی", "لطفاً حداقل یک شماره سفارش وارد کنید.")
            return

        self.process_button.setDisabled(True)
        self.settings_button.setDisabled(True)
        self.status_box.clear()
        
        # Create worker thread
        self.thread = QThread()
        self.worker = Worker(order_numbers)
        self.worker.moveToThread(self.thread)

        # Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.status_update.connect(self.update_status)

        # Connect message box signals
        self.worker.error_signal.connect(self.show_error_message)
        self.worker.warning_signal.connect(self.show_warning_message)
        self.worker.info_signal.connect(self.show_info_message)

        # Final action after thread finishes
        self.thread.finished.connect(lambda: self.process_button.setDisabled(False))
        self.thread.finished.connect(lambda: self.settings_button.setDisabled(False))
        
        # Start the thread
        self.thread.start()


    def update_status(self, message):
        self.status_box.append(message)
        self.status_box.verticalScrollBar().setValue(self.status_box.verticalScrollBar().maximum())


    def show_error_message(self, title, message):
        QMessageBox.critical(self, title, message)


    def show_warning_message(self, title, message):
        QMessageBox.warning(self, title, message)


    def show_info_message(self, title, message):
        QMessageBox.information(self, title, message)


    def show_settings(self):
        QMessageBox.information(self, "ویژگی‌ها", "این بخش برای تنظیمات آینده در نظر گرفته شده است.")


def main():
    app = QApplication(sys.argv)
    # Set layout direction to Right-to-Left for Persian UI
    app.setLayoutDirection(Qt.RightToLeft)
    ex = ProdPlanApp()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()


