"""Generate a right-to-left (RTL) production order sheet from ERP data."""

import re
from datetime import date, datetime

import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Flowable, HRFlowable, KeepTogether, SimpleDocTemplate, Spacer, Table, TableStyle,
)

FONT_NAME = "FardanIRAN"

# Sizes (pt)
TITLE_SIZE = 15
BODY_SIZE = 11
SUBTITLE_SIZE = 9.5      # customer's product name: one step smaller than the technical code
LINE_FACTOR = 1.5

# In the approved sample the "prepared by" block sits at the left edge. Use "R" for the right edge.
PREPARER_ALIGN = "L"

# ---- Footer texts: edit these by hand when needed ---------------------------------------
PREPARER_UNIT = "واحد فروش"        # printed after "تهیه کننده:"
PREPARER_NAME = "فرنوش نصیری"      # <-- نام تهیه‌کننده؛ برای تغییر، فقط همین مقدار را عوض کنید
COPY_TO = ["واحد برنامه ریزی", "واحد مهندسی", "واحد انبار", "واحد تولید"]   # "رونوشت"

# The IRAN font has no "isolated" presentation forms for several letters (ا ت د ر ...);
# drawing them as the plain letters avoids the empty boxes seen in the old output.
_RESHAPER = arabic_reshaper.ArabicReshaper(configuration={"use_unshaped_instead_of_isolated": True})

_PERSIAN_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
_ARABIC_TO_PERSIAN = str.maketrans({"ي": "ی", "ك": "ک", "ى": "ی"})


# ------------------------------------------------------------------------------
# Text helpers
# ------------------------------------------------------------------------------
def _clean(value):
    return "" if value is None else str(value).strip()


def _visual(text, base_dir="R"):
    """Logical text -> shaped, visually ordered text ready to be drawn left-to-right."""
    text = _clean(text).translate(_ARABIC_TO_PERSIAN)
    if not text:
        return ""
    shaped = _RESHAPER.reshape(text).replace("\u200c", "")
    return get_display(shaped, base_dir=base_dir)


def _width(text, size, base_dir="R"):
    return pdfmetrics.stringWidth(_visual(text, base_dir), FONT_NAME, size)


def _hard_split(word, size, max_width, base_dir):
    """Break a single over-long word into pieces that fit."""
    pieces, current = [], ""
    for char in word:
        if current and _width(current + char, size, base_dir) > max_width:
            pieces.append(current)
            current = char
        else:
            current += char
    if current:
        pieces.append(current)
    return pieces


def _wrap(text, size, max_width, base_dir="R"):
    """Wrap *logical* text into lines that fit max_width (breaking before reordering)."""
    lines = []
    for paragraph in (_clean(text).splitlines() or [""]):
        current = ""
        for word in paragraph.split():
            if _width(word, size, base_dir) > max_width:
                if current:
                    lines.append(current)
                    current = ""
                *head, tail = _hard_split(word, size, max_width, base_dir)
                lines.extend(head)
                current = tail
                continue
            candidate = f"{current} {word}" if current else word
            if current and _width(candidate, size, base_dir) > max_width:
                lines.append(current)
                current = word
            else:
                current = candidate
        lines.append(current)
    return lines


def _to_jalali(gy, gm, gd):
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = gy + 1 if gm > 2 else gy
    days = (355666 + 365 * gy + (gy2 + 3) // 4 - (gy2 + 99) // 100 + (gy2 + 399) // 400
            + gd + g_d_m[gm - 1])
    jy = -1595 + 33 * (days // 12053)
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm, jd = 1 + days // 31, 1 + days % 31
    else:
        jm, jd = 7 + (days - 186) // 30, 1 + (days - 186) % 30
    return jy, jm, jd


def _format_date(value):
    """Gregorian date -> Jalali 'YYYY/MM/DD' with Persian digits."""
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        jy, jm, jd = _to_jalali(value.year, value.month, value.day)
        return f"{jy:04d}/{jm:02d}/{jd:02d}".translate(_PERSIAN_DIGITS)
    return _clean(value)


def _format_quantity(value):
    if value is None:
        return ""
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else str(value)
    except (TypeError, ValueError):
        return _clean(value)


def _join_notes(*values):
    """Item notes first, then delivery time; ' ، ' only between two non-empty values."""
    return "، ".join(text for text in (_clean(v) for v in values) if text)


def merged_pdf_filename(order, fallback_number=""):
    """Final merged file name: '(5370) استیل البرز.pdf' (order number without 'ORD', one space, customer)."""
    order = order or {}
    number = _clean(order.get("order_number")) or _clean(fallback_number)
    digits = re.findall(r"\d+", number)
    number = digits[-1] if digits else number
    company = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", _clean(order.get("company_name")))
    company = re.sub(r"\s+", " ", company).strip(" .")
    return (f"({number}) {company}" if company else f"({number})") + ".pdf"


# ------------------------------------------------------------------------------
# Flowable that draws right-to-left text (wrapping done on logical text)
# ------------------------------------------------------------------------------
class RtlText(Flowable):
    """parts: list of (text, font_size); align: 'R' | 'C' | 'L'; base_dir: 'R' | 'L'."""

    def __init__(self, parts, align="R", base_dir="R", bold=False):
        super().__init__()
        self.parts = [(text, size) for text, size in parts if _clean(text)]
        self.align = align
        self.base_dir = base_dir
        self.bold = bold
        self._lines = []

    def wrap(self, avail_width, avail_height):
        self.width = avail_width
        self._lines = []
        height = 0
        for text, size in self.parts:
            for line in _wrap(text, size, avail_width - 1, self.base_dir):
                self._lines.append((_visual(line, self.base_dir), size))
                height += size * LINE_FACTOR
        self.height = height or BODY_SIZE * LINE_FACTOR
        return self.width, self.height

    def draw(self):
        canv = self.canv
        y = self.height
        for visual, size in self._lines:
            lead = size * LINE_FACTOR
            y -= lead
            baseline = y + (lead - size) / 2 + size * 0.25
            text_width = pdfmetrics.stringWidth(visual, FONT_NAME, size)
            if self.align == "R":
                x = self.width - text_width
            elif self.align == "C":
                x = (self.width - text_width) / 2
            else:
                x = 0
            text = canv.beginText(x, baseline)
            text.setFont(FONT_NAME, size)
            if self.bold:  # IRAN has only a regular weight: fake bold with a thin outline
                text.setTextRenderMode(2)
                canv.setStrokeColor(colors.black)
                canv.setLineWidth(size * 0.035)
            text.textOut(visual)
            canv.drawText(text)


def _copy_to_table():
    """One row of unit names, first unit at the right edge, with wider gaps between them."""
    gap = 8 * mm
    units = list(reversed(COPY_TO))  # visual left-to-right order
    widths = [_width(unit, BODY_SIZE) + 2 + (gap if i else 0) for i, unit in enumerate(units)]
    table = Table([[RtlText([(unit, BODY_SIZE)]) for unit in units]],
                  colWidths=widths, hAlign="RIGHT")
    table.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), gap), ("LEFTPADDING", (0, 0), (0, 0), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return table


# ------------------------------------------------------------------------------
# Sheet
# ------------------------------------------------------------------------------
def create_order_sheet_pdf(order, output_path, font_path):
    """Create the RTL order sheet from ERP data; never uses a legacy order PDF."""
    try:
        pdfmetrics.getFont(FONT_NAME)
    except KeyError:
        pdfmetrics.registerFont(TTFont(FONT_NAME, font_path))

    margin = 18 * mm
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=margin, rightMargin=margin, topMargin=20 * mm, bottomMargin=15 * mm,
        title=f"برگه سفارش {_clean(order.get('order_number'))}",
    )
    content_width = A4[0] - 2 * margin
    half = content_width / 2
    pad = dict(LEFTPADDING=5, RIGHTPADDING=5, TOPPADDING=4, BOTTOMPADDING=4)

    def cell(text, **kwargs):
        return RtlText([(text, BODY_SIZE)], **kwargs)

    story = [RtlText([("برگه سفارش", TITLE_SIZE)], align="C", bold=True), Spacer(1, 6 * mm)]

    # ---- Summary block: starts at the right edge --------------------------------
    delivery_plan = re.sub(r"^\s*پلن\s+", "", _clean(order.get("delivery_time")))
    summary = [
        [cell(f"مشتری: {_clean(order.get('company_name'))}"),
         cell(f"شماره سفارش: {_clean(order.get('order_number'))}")],
        [cell(f"پلن تحویل: {delivery_plan}"),
         cell(f"تاریخ ثبت: {_format_date(order.get('order_date'))}")],
    ]
    style = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    if _clean(order.get("notes")):
        summary.append([cell(f"توضیحات: {_clean(order.get('notes'))}"), ""])
        style.append(("SPAN", (0, 2), (1, 2)))
    style.extend((name, (0, 0), (-1, -1), value) for name, value in pad.items())
    header = Table(summary, colWidths=[half, half])
    header.setStyle(TableStyle(style))
    story.extend([header, Spacer(1, 8 * mm)])

    # ---- Items table: columns run right-to-left (row number is the right-most) ---
    widths = [34 * mm, 36 * mm, 20 * mm, 0, 15 * mm]
    widths[3] = content_width - sum(widths)
    labels = ["سریال تولید", "ملاحظات", "تعداد", "کد / شرح محصول", "ردیف"]
    rows = [[cell(label, align="C") for label in labels]]
    for index, item in enumerate(order.get("items", []), 1):
        rows.append([
            cell(item.get("serial"), align="C", base_dir="L"),
            cell(_join_notes(item.get("item_notes"), item.get("delivery_time"))),
            cell(_format_quantity(item.get("quantity")), align="C"),
            RtlText([(item.get("technical_code"), BODY_SIZE),
                     (item.get("product_name"), SUBTITLE_SIZE)]),
            cell(str(index).translate(_PERSIAN_DIGITS), align="C"),
        ])
    style = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.91, 0.91, 0.91)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("VALIGN", (1, 1), (1, -1), "TOP"),   # notes
        ("VALIGN", (3, 1), (3, -1), "TOP"),   # code / description
    ]
    style.extend((name, (0, 0), (-1, -1), value) for name, value in pad.items())
    table = Table(rows, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle(style))
    story.append(table)

    # ---- Footer ----------------------------------------------------------------
    footer = [Spacer(1, 20 * mm),
              RtlText([(f"تهیه کننده: {PREPARER_UNIT}", BODY_SIZE)], align=PREPARER_ALIGN),
              Spacer(1, 3 * mm),
              RtlText([(PREPARER_NAME, BODY_SIZE)], align=PREPARER_ALIGN)]
    footer.extend([
        Spacer(1, 12 * mm),
        HRFlowable(width="100%", thickness=0.8, color=colors.black, spaceAfter=3 * mm),
        RtlText([("رونوشت:", BODY_SIZE)]),
        Spacer(1, 1 * mm),
        _copy_to_table(),
    ])
    story.append(KeepTogether(footer))

    doc.build(story)
