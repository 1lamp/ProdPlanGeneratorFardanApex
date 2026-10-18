"""Generate a current order sheet from ERP data."""
from datetime import date, datetime
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


def _persian(value):
    text = "" if value is None else str(value)
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(text))
    except ImportError:
        return text


def create_order_sheet_pdf(order, output_path, font_path):
    """Create an order sheet from ERP data; never use a legacy order PDF."""
    font_name = "FardanIRAN"
    try:
        pdfmetrics.getFont(font_name)
    except KeyError:
        pdfmetrics.registerFont(TTFont(font_name, font_path))
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("FardanNormal", parent=styles["Normal"], fontName=font_name,
                            fontSize=9, leading=14, alignment=TA_RIGHT, wordWrap="CJK")
    title = ParagraphStyle("FardanTitle", parent=normal, fontSize=15, leading=22,
                           alignment=TA_CENTER, spaceAfter=8)
    small = ParagraphStyle("FardanSmall", parent=normal, fontSize=8, leading=11)
    doc = SimpleDocTemplate(output_path, pagesize=landscape(A4),
                            rightMargin=12*mm, leftMargin=12*mm,
                            topMargin=10*mm, bottomMargin=10*mm)
    story = [Paragraph(_persian("برگه سفارش تولید"), title)]
    order_date = order.get("order_date")
    if isinstance(order_date, (date, datetime)):
        order_date = order_date.strftime("%Y-%m-%d")
    summary = [
        [Paragraph(_persian("شماره سفارش"), normal), Paragraph(_persian(order.get("order_number")), normal),
         Paragraph(_persian("تاریخ ثبت سفارش"), normal), Paragraph(_persian(order_date), normal)],
        [Paragraph(_persian("مشتری"), normal), Paragraph(_persian(order.get("company_name")), normal),
         Paragraph(_persian("زمان تحویل کلی"), normal), Paragraph(_persian(order.get("delivery_time")), normal)],
        [Paragraph(_persian("توضیحات سفارش"), normal), Paragraph(_persian(order.get("notes")), normal), "", ""],
    ]
    header = Table(summary, colWidths=[27*mm, 72*mm, 30*mm, 95*mm])
    header.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), .4, colors.grey),
        ("BACKGROUND", (0,0), (0,-1), colors.whitesmoke),
        ("BACKGROUND", (2,0), (2,1), colors.whitesmoke),
        ("SPAN", (1,2), (3,2)), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (-1,-1), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 5), ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story.extend([header, Spacer(1, 5*mm)])
    rows = [[Paragraph(_persian(label), normal) for label in
             ["ردیف", "کد محصول", "شرح کالا", "تعداد", "ملاحظات قلم", "زمان تحویل", "سریال تولید"]]]
    for index, item in enumerate(order.get("items", []), 1):
        rows.append([
            Paragraph(str(index), normal),
            Paragraph(_persian(item.get("technical_code")), normal),
            Paragraph(_persian(item.get("product_name")), normal),
            Paragraph(_persian(item.get("quantity")), normal),
            Paragraph(_persian(item.get("item_notes")), small),
            Paragraph(_persian(item.get("delivery_time")), small),
            Paragraph(_persian(item.get("serial")), normal),
        ])
    table = Table(rows, repeatRows=1, colWidths=[12*mm, 30*mm, 50*mm, 18*mm, 70*mm, 40*mm, 32*mm])
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), .4, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (0,0), (-1,-1), "RIGHT"),
        ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story.append(table)
    doc.build(story)