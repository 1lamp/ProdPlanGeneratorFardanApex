"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Fardan Apex --- ProdPlanGenerator ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
This software automates processing factory orders by generating, organizing,
and exporting production documents into structured PDF files.

Author: Behnam Rabieyan
Company: Garma Gostar Fardan
Created: 2025
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox
import pandas as pd
import xlwings as xw
import os
import threading
import shutil
import re
from PyPDF2 import PdfWriter


# ==============================================================================
# بخش تنظیمات: تمام مسیرها، نام‌ها و آدرس‌ها از اینجا قابل ویرایش هستند
# ==============================================================================
# --- فایل‌های اصلی ---
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


# --- نام شیت‌ها و ستون‌ها ---
ORDER_SHEET_NAME = "OrderList"
DATABASE_SHEET_NAME = "LOM"
COL_ORDER_NUM = "شماره سفارش"
COL_PRODUCT_CODE = "کد محصول"
COL_QUANTITY = "تعداد"


# --- سلول‌های ورودی در شیت LOM ---
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


def find_last_numeric_row(sheet, search_range):
    """آخرین ردیف حاوی مقدار عددی را در یک محدوده مشخص پیدا می‌کند."""
    values = sheet.range(search_range).options(ndim=1).value
    start_row = sheet.range(search_range).row
    for i in range(len(values) - 1, -1, -1):
        if isinstance(values[i], (int, float)) and values[i] is not None:
            return start_row + i
    return 0


def print_conditional_sheet(sheet, product_code, pdf_filepath, config, status_callback, order_num=None):
    """یک برگه شرطی را بر اساس تنظیمات داده شده، پردازش و چاپ می‌کند."""
    try:
        sheet.range(config['cell_product']).value = product_code
        if order_num and 'cell_order' in config:
            sheet.range(config['cell_order']).value = order_num
        if 'check_cell' in config and 'cell_flag' in config:
            check_val = str(sheet.range(config['check_cell']).value).strip().upper()
            sheet.range(config['cell_flag']).value = (check_val == 'FALSE')
        sheet.range(config['print_range']).api.ExportAsFixedFormat(0, pdf_filepath)
        status_callback(f"    - چاپ شرطی ({sheet.name}): {os.path.basename(pdf_filepath)}\n")
        return True
    except Exception as e:
        status_callback(f"    - خطا در چاپ شیت '{sheet.name}': {e}\n")
        return False


def process_orders_in_thread(order_numbers_str, status_update_callback):
    """تابع اصلی پردازش که تمام منطق برنامه در آن قرار دارد."""
    try:
        order_numbers_list = [num.strip() for num in order_numbers_str.strip().split('\n') if num.strip()]
        if not order_numbers_list:
            messagebox.showerror("ورودی خالی", "هیچ شماره سفارشی برای پردازش وارد نشده است.")
            return
        order_numbers_int = [int(num) for num in order_numbers_list]

        status_update_callback(f"شماره‌های سفارش برای پردازش: {order_numbers_list}\n")
        status_update_callback(f"در حال خواندن فایل سفارش‌ها: {ORDER_FILE_PATH}...\n")
        df = pd.read_excel(ORDER_FILE_PATH, sheet_name=ORDER_SHEET_NAME, engine='openpyxl')
        df.columns = df.columns.str.strip()
        filtered_df = df[df[COL_ORDER_NUM.strip()].isin(order_numbers_int)]

        if filtered_df.empty:
            messagebox.showwarning("یافت نشد", "هیچ آیتمی مطابق با شماره سفارش‌های وارد شده در اکسل سفارش‌ها یافت نشد.")
            return
        status_update_callback(f"تعداد {len(filtered_df)} آیتم برای پردازش یافت شد.\n")

        status_update_callback("شروع پردازش اکسل دیتابیس و چاپ برگه‌ها...\n")
        with xw.App(visible=False) as app:
            db_wb = app.books.open(DATABASE_FILE_PATH, read_only=True)
            db_sheet = db_wb.sheets[DATABASE_SHEET_NAME]

            for order_num, group in filtered_df.groupby(COL_ORDER_NUM.strip()):
                status_update_callback(f"\n===== شروع پردازش سفارش شماره: {order_num} =====\n")
                order_folder = os.path.join(OUTPUT_BASE_PATH, str(order_num))
                os.makedirs(order_folder, exist_ok=True)

                main_merger = PdfWriter()
                preparation_merger = PdfWriter()
                timing_merger = PdfWriter()
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
                            status_update_callback(f"  - فایل اصلی سفارش '{filename}' کپی و به لیست اصلی اضافه شد.\n")
                            break
                    if not original_order_filename:
                         status_update_callback(f"  - هشدار: فایل PDF اصلی برای سفارش {order_num} یافت نشد.\n")
                except Exception as e:
                    status_update_callback(f"  - خطا در کپی فایل اصلی سفارش: {e}\n")

                for index, row in group.iterrows():
                    original_product_code = str(row[COL_PRODUCT_CODE.strip()])
                    quantity = row[COL_QUANTITY.strip()]
                    status_update_callback(f"-> بررسی محصول: {original_product_code}\n")

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
                        status_update_callback(f"  - هشدار: محصول {original_product_code} نامعتبر است. این آیتم نادیده گرفته شد.\n")
                        continue

                    for final_code in valid_product_codes:
                        status_update_callback(f"  * شروع فرآیند چاپ برای محصول: {final_code}\n")
                        db_sheet.range(CELL_ORDER_NUM_DB).value = order_num
                        db_sheet.range(CELL_QUANTITY).value = quantity
                        db_sheet.range(CELL_PRODUCT_CODE).value = final_code

                        # مرحله ۱: چاپ برگه‌های LOM، آماده‌سازی و زمانسنجی
                        for job in LOM_PRINT_JOBS:
                            suffix = job['suffix']; job_type = job['type']
                            should_print = True; print_range = ""
                            if suffix == "LOM":
                                last_row = find_last_numeric_row(db_sheet, 'B5:B65'); print_range = f"B1:G{last_row}" if last_row > 0 else ""
                            elif suffix == "زمانسنجی":
                                if db_sheet.range('P9').value is None: should_print = False
                                else: last_row = find_last_numeric_row(db_sheet, 'N9:N47'); print_range = f"N4:Q{last_row}" if last_row > 0 else ""
                            elif suffix == "آماده سازی":
                                if db_sheet.range('U5').value is None: should_print = False
                                else: last_row = find_last_numeric_row(db_sheet, 'S5:S24'); print_range = f"S1:Y{last_row}" if last_row > 0 else ""
                            
                            if not print_range: should_print = False

                            if should_print:
                                pdf_filename = f"{final_code}_{suffix}.pdf"
                                pdf_filepath = os.path.join(order_folder, pdf_filename)
                                db_sheet.range(print_range).api.ExportAsFixedFormat(0, pdf_filepath)
                                files_to_delete.append(pdf_filepath)


                                if job_type == 'main': main_merger.append(pdf_filepath)
                                elif job_type == 'preparation': preparation_merger.append(pdf_filepath)
                                elif job_type == 'timing': timing_merger.append(pdf_filepath)
                                status_update_callback(f"    - PDF ({suffix}) ذخیره و به لیست مربوطه اضافه شد.\n")
                            else:
                                status_update_callback(f"    - چاپ '{suffix}' به دلیل نبود اطلاعات لغو شد.\n")

                        # مرحله ۲: چاپ برگه‌های شرطی
                        process_code = str(db_sheet.range(CONDITIONAL_CHECK_CELL).value)[:2].upper()
                        if process_code == 'MF':
                            mf_sheet = db_wb.sheets[MF_SHEET_NAME]
                            pdf_filepath = os.path.join(order_folder, f"{final_code}_{MF_SHEET_NAME}.pdf")
                            if print_conditional_sheet(mf_sheet, final_code, pdf_filepath, MF_CONFIG, status_update_callback, order_num=order_num):
                                main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)
                        if process_code in ('DS', 'DF', 'NL', 'DL'):
                            st_sheet = db_wb.sheets[ST_SHEET_NAME]
                            pdf_filepath = os.path.join(order_folder, f"{final_code}_{ST_SHEET_NAME}.pdf")
                            if print_conditional_sheet(st_sheet, final_code, pdf_filepath, ST_CONFIG, status_update_callback):
                                main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)
                        if process_code in ('NL', 'DL'):
                            kl_sheet = db_wb.sheets[KL_SHEET_NAME]
                            pdf_filepath = os.path.join(order_folder, f"{final_code}_{KL_SHEET_NAME}.pdf")
                            if print_conditional_sheet(kl_sheet, final_code, pdf_filepath, KL_CONFIG, status_update_callback):
                                main_merger.append(pdf_filepath); files_to_delete.append(pdf_filepath)

                        # مرحله ۳: کپی نقشه فنی (آخرین آیتم برای هر محصول)
                        drawing_source_folder = TECHNICAL_DRAWING_PATHS.get(process_code)
                        if drawing_source_folder:
                            source_drawing_name = f"{final_code[:6]}.pdf"
                            dest_drawing_name = f"{final_code}_نقشه.pdf"
                            source_drawing_path = os.path.join(drawing_source_folder, source_drawing_name)
                            dest_drawing_path = os.path.join(order_folder, dest_drawing_name)
                            if os.path.exists(source_drawing_path):
                                shutil.copy(source_drawing_path, dest_drawing_path)
                                main_merger.append(dest_drawing_path)
                                files_to_delete.append(dest_drawing_path)
                                status_update_callback(f"    - نقشه فنی '{dest_drawing_name}' کپی و به لیست اصلی اضافه شد.\n")
                            else:
                                status_update_callback(f"    - هشدار: نقشه فنی '{source_drawing_name}' یافت نشد.\n")

                # ذخیره و پاکسازی نهایی برای سفارش فعلی
                if len(main_merger.pages) > 0:
                    clean_name = str(order_num)
                    if original_order_filename: clean_name = re.sub(r'\s*ok', '', original_order_filename, flags=re.IGNORECASE).replace('.pdf', '').strip()
                    output_path = os.path.join(order_folder, f"{clean_name}.pdf")
                    with open(output_path, "wb") as f: main_merger.write(f)
                    status_update_callback(f"  -> فایل اصلی ادغام شده برای سفارش {order_num} ذخیره شد.\n")
                if len(preparation_merger.pages) > 0:
                    output_path = os.path.join(order_folder, f"آماده سازی({order_num}).pdf")
                    with open(output_path, "wb") as f: preparation_merger.write(f)
                    status_update_callback(f"  -> فایل 'آماده سازی' ادغام شده برای سفارش {order_num} ذخیره شد.\n")
                if len(timing_merger.pages) > 0:
                    output_path = os.path.join(order_folder, f"زمانسنجی({order_num}).pdf")
                    with open(output_path, "wb") as f: timing_merger.write(f)
                    status_update_callback(f"  -> فایل 'زمانسنجی' ادغام شده برای سفارش {order_num} ذخیره شد.\n")
                
                main_merger.close(); preparation_merger.close(); timing_merger.close()

                if files_to_delete:
                    status_update_callback(f"  -> شروع پاکسازی فایل‌های موقت برای سفارش {order_num}...\n")
                    deleted_count = 0
                    for file_path in files_to_delete:
                        try:
                            if os.path.exists(file_path): os.remove(file_path); deleted_count += 1
                        except Exception as e:
                            status_update_callback(f"    - خطا در حذف فایل {os.path.basename(file_path)}: {e}\n")
                    status_update_callback(f"    - {deleted_count} فایل موقت با موفقیت حذف شد.\n")

            db_wb.close()
            status_update_callback("\nعملیات پردازش، ساخت، ادغام و پاکسازی با موفقیت به پایان رسید.\n")
            messagebox.showinfo("اتمام عملیات", "تمام سفارش‌ها با موفقیت پردازش شدند.")

    except Exception as e:
        messagebox.showerror("خطای کلی", f"یک خطای ناشناخته در برنامه رخ داد:\n{e}")
        status_update_callback(f"خطای بحرانی: {e}\n")


def create_gui():
    """این تابع رابط کاربری گرافیکی برنامه را می‌سازد."""
    def start_processing_action():
        start_button.config(state=tk.DISABLED)
        status_box.config(state=tk.NORMAL); status_box.delete('1.0', tk.END); status_box.config(state=tk.DISABLED)
        order_numbers_str = order_input.get("1.0", tk.END)
        def update_gui_status(message):
            status_box.config(state=tk.NORMAL)
            status_box.insert(tk.END, message)
            status_box.see(tk.END)
            status_box.config(state=tk.DISABLED)
            if "خطا" in message or "پایان" in message or "اتمام" in message:
                start_button.config(state=tk.NORMAL)
        processing_thread = threading.Thread(target=process_orders_in_thread, args=(order_numbers_str, update_gui_status))
        processing_thread.daemon = True
        processing_thread.start()
    root = tk.Tk()
    root.title("سازنده و ادغام کننده PDF پیشرفته از روی لیست سفارش")
    root.geometry("700x650")
    main_frame = tk.Frame(root, padx=15, pady=15)
    main_frame.pack(fill=tk.BOTH, expand=True)
    input_label = tk.Label(main_frame, text="شماره‌های سفارش را وارد کنید (هر کدام در یک خط جدید):", justify=tk.RIGHT)
    input_label.pack(fill=tk.X, pady=(0, 5))
    order_input = scrolledtext.ScrolledText(main_frame, height=8, font=("Tahoma", 10))
    order_input.pack(fill=tk.BOTH, expand=True)
    order_input.focus()
    start_button = tk.Button(main_frame, text="شروع پردازش، ساخت و ادغام همه PDF ها", command=start_processing_action, font=('Tahoma', 11, 'bold'), bg="#4CAF50", fg="white")
    start_button.pack(pady=15, ipady=5, fill=tk.X)
    status_label = tk.Label(main_frame, text="وضعیت پردازش:", justify=tk.RIGHT)
    status_label.pack(fill=tk.X, pady=(10, 5))
    status_box = scrolledtext.ScrolledText(main_frame, height=20, state=tk.DISABLED, bg="#f0f0f0", font=("Tahoma", 9))
    status_box.pack(fill=tk.BOTH, expand=True)
    root.mainloop()


if __name__ == "__main__":
    create_gui()


