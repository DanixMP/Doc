#!/usr/bin/env python3
"""Generate RTL Persian proposal DOCX — YOLO-centric stack, no Latin in Persian text."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont

ROOT = Path("/workspace")
FONT_DIR = ROOT / "fonts" / "vazirmatn" / "fonts" / "ttf"
ASSETS = ROOT / "docs" / "assets"
DOC_FONTS = ROOT / "docs" / "fonts"
OUT_DOCX = ROOT / "docs" / "پیشنهادیه-سامانه-هوشمند-پاپ-اسمیر.docx"

FONT_REG = FONT_DIR / "Vazirmatn-Regular.ttf"
FONT_BOLD = FONT_DIR / "Vazirmatn-Bold.ttf"
FONT_MED = FONT_DIR / "Vazirmatn-Medium.ttf"
FONT_LIGHT = FONT_DIR / "Vazirmatn-Light.ttf"
FONT_NAME = "Vazirmatn"

C_TEAL = "#0F6B6B"
C_TEAL_DARK = "#0A4F4F"
C_TEAL_LIGHT = "#E6F3F3"
C_CORAL = "#C45C4A"
C_GOLD = "#C4A35A"
C_INK = "#1C2B2B"
C_MUTED = "#5A6B6B"
C_BG = "#F7FAFA"
C_WHITE = "#FFFFFF"
C_LINE = "#D5E3E3"

TABLE_COUNTER = {"n": 0}


def to_fa(s) -> str:
    return str(s).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


def load_font(path: Path, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(path), size=size)


def draw_rtl(draw, xy, text, font, fill, anchor="rt"):
    draw.text(xy, text, font=font, fill=fill, anchor=anchor, direction="rtl", language="fa")


def hex_rgb(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------------------
# Charts — Persian only + numbers
# ---------------------------------------------------------------------------

def make_cover_banner(path: Path):
    W, H = 1800, 980
    img = Image.new("RGB", (W, H), hex_rgb(C_TEAL_DARK))
    draw = ImageDraw.Draw(img)
    base, top = hex_rgb(C_TEAL_DARK), (18, 95, 95)
    for y in range(H):
        t = y / H
        col = tuple(int(base[i] * (1 - t) + top[i] * t) for i in range(3))
        draw.line([(0, y), (W, y)], fill=col)

    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.ellipse([-180, -120, 520, 520], fill=(255, 255, 255, 30))
    od.ellipse([1250, 420, 1900, 1100], fill=(196, 163, 90, 55))
    od.ellipse([980, -160, 1500, 320], fill=(255, 255, 255, 22))
    od.rounded_rectangle([90, 120, 1710, 860], radius=36, outline=(255, 255, 255, 55), width=2)
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)

    draw_rtl(draw, (W // 2, 175), "بسمه تعالی", load_font(FONT_BOLD, 44), hex_rgb(C_GOLD), "mm")
    draw_rtl(draw, (W // 2, 245), "فرم درخواست طرح توسعه فناوری", load_font(FONT_MED, 36), (230, 242, 242), "mm")
    draw.rectangle([620, 280, 1180, 286], fill=hex_rgb(C_GOLD))
    draw_rtl(draw, (W // 2, 340), "پیشنهادیه جامع طرح", load_font(FONT_BOLD, 52), (255, 255, 255), "mm")

    title_lines = [
        "توسعه سامانه هوشمند شناسایی و طبقه‌بندی ناهنجاری‌های",
        "سیتولوژیک در نمونه‌های پاپ اسمیر به‌منظور تشخیص",
        "سرطان دهانه رحم در بانوان با کمک هوش مصنوعی",
    ]
    y = 430
    for line in title_lines:
        draw_rtl(draw, (W // 2, y), line, load_font(FONT_BOLD, 40), (240, 250, 250), "mm")
        y += 58

    draw.rectangle([0, H - 22, W, H], fill=hex_rgb(C_GOLD))
    draw_rtl(
        draw, (W // 2, H - 70),
        "هسته یولو · آموزش محلی · محرمانگی داده · دوازده ماه",
        load_font(FONT_MED, 30), (220, 235, 235), "mm",
    )
    img.save(path, quality=95)


def chart_timeline(path: Path):
    W, H = 1600, 720
    img = Image.new("RGB", (W, H), hex_rgb(C_BG))
    draw = ImageDraw.Draw(img)
    draw_rtl(draw, (W - 60, 45), "زمان‌بندی دوازده‌ماهه طرح", load_font(FONT_BOLD, 36), hex_rgb(C_TEAL_DARK), "rt")

    phases = [
        ("۱. راه‌اندازی و اخلاق", 0, 2, C_TEAL),
        ("۲. جمع‌آوری و برچسب‌گذاری", 1.5, 5, C_TEAL_DARK),
        ("۳. آموزش مدل یولو", 3.5, 8, C_CORAL),
        ("۴. نمونه سامانه پشتیبان", 6.5, 10, C_GOLD),
        ("۵. پایلوت بیمارستانی", 8.5, 11, "#3D8B8B"),
        ("۶. مستندسازی و تجاری‌سازی", 10.5, 12, C_MUTED),
    ]
    left, right, top, row_h = 80, 1120, 110, 80
    scale = (right - left) / 12
    for m in range(0, 13):
        x = left + m * scale
        draw.line([(x, top - 10), (x, top + len(phases) * row_h)], fill=hex_rgb(C_LINE), width=1)
        draw_rtl(draw, (x, top + len(phases) * row_h + 18), to_fa(m), load_font(FONT_REG, 22), hex_rgb(C_MUTED), "mm")
    draw_rtl(draw, ((left + right) / 2, H - 40), "ماه", load_font(FONT_REG, 22), hex_rgb(C_MUTED), "mm")

    for i, (name, start, end, color) in enumerate(phases):
        y0 = top + i * row_h + 18
        x0 = left + start * scale
        x1 = left + end * scale
        draw.rounded_rectangle([x0, y0, x1, y0 + 44], radius=10, fill=hex_rgb(color))
        draw_rtl(draw, (right + 40, y0 + 22), name, load_font(FONT_MED, 26), hex_rgb(C_INK), "rt")
    img.save(path, quality=95)


def chart_cost(path: Path):
    W, H = 1400, 780
    img = Image.new("RGB", (W, H), hex_rgb(C_BG))
    draw = ImageDraw.Draw(img)
    draw_rtl(draw, (W - 50, 40), "مقایسه هزینه تقریبی آموزش محلی و ابری", load_font(FONT_BOLD, 34), hex_rgb(C_TEAL_DARK), "rt")

    items = [
        ("۱. پشته محلی\nپیشنهادی", 4, C_TEAL),
        ("۲. سکوی ابری پایه\nبا مصرف اعتبار", 6.5, C_GOLD),
        ("۳. سکوی ابری سازمانی", 18, C_CORAL),
    ]
    chart_top, chart_bottom = 120, 560
    chart_left, gap, bar_w = 180, 120, 180
    max_v = 22
    for v in [0, 5, 10, 15, 20]:
        y = chart_bottom - (v / max_v) * (chart_bottom - chart_top)
        draw.line([(140, y), (1200, y)], fill=hex_rgb(C_LINE), width=1)
        draw_rtl(draw, (125, y), to_fa(v), load_font(FONT_REG, 22), hex_rgb(C_MUTED), "rm")

    for i, (name, val, color) in enumerate(items):
        x0 = chart_left + i * (bar_w + gap)
        h = (val / max_v) * (chart_bottom - chart_top)
        y0 = chart_bottom - h
        draw.rounded_rectangle([x0, y0, x0 + bar_w, chart_bottom], radius=14, fill=hex_rgb(color))
        draw_rtl(draw, (x0 + bar_w / 2, y0 - 28), to_fa(f"حدود {val} هزار دلار"), load_font(FONT_MED, 22), hex_rgb(C_INK), "mm")
        yy = chart_bottom + 28
        for line in name.split("\n"):
            draw_rtl(draw, (x0 + bar_w / 2, yy), line, load_font(FONT_MED, 24), hex_rgb(C_INK), "mm")
            yy += 32
    draw_rtl(draw, (W / 2, H - 35), "هزینه تقریبی سال اول", load_font(FONT_REG, 22), hex_rgb(C_MUTED), "mm")
    img.save(path, quality=95)


def chart_stack(path: Path):
    """Numbered YOLO-centric stack — Persian only. Labeling feeds YOLO."""
    W, H = 1700, 920
    img = Image.new("RGB", (W, H), hex_rgb(C_BG))
    draw = ImageDraw.Draw(img)
    draw_rtl(draw, (W - 50, 40), "نمودار پشته فناوری با هسته یولو", load_font(FONT_BOLD, 34), hex_rgb(C_TEAL_DARK), "rt")

    def box(x, y, w, h, fill, num, title, sub):
        draw.rounded_rectangle([x, y, x + w, y + h], radius=18, fill=hex_rgb(fill))
        draw_rtl(draw, (x + w / 2, y + 28), to_fa(num), load_font(FONT_BOLD, 26), (255, 255, 255), "mm")
        draw_rtl(draw, (x + w / 2, y + h / 2 + 4), title, load_font(FONT_BOLD, 22), (255, 255, 255), "mm")
        draw_rtl(draw, (x + w / 2, y + h / 2 + 34), sub, load_font(FONT_REG, 17), (235, 245, 245), "mm")

    def arrow(x1, y1, x2, y2):
        draw.line([(x1, y1), (x2, y2)], fill=hex_rgb(C_MUTED), width=3)
        draw.polygon([(x2, y2), (x2 - 12, y2 - 7), (x2 - 12, y2 + 7)], fill=hex_rgb(C_MUTED))

    # 1 read → 2 label / 3 store → 4 YOLO core → 5 clinical → 6 service → 7 doctor
    box(40, 360, 200, 130, C_TEAL_DARK, "۱", "خواندن اسلاید", "برش قطعات تصویر")
    box(300, 180, 220, 130, "#3D8B8B", "۲", "برچسب‌گذاری", "سامانه محلی برچسب")
    box(300, 540, 220, 130, C_MUTED, "۳", "ذخیره امن", "مخزن محلی داده")
    box(600, 360, 260, 130, C_CORAL, "۴", "هسته یولو", "تشخیص و طبقه‌بندی")
    box(930, 360, 230, 130, C_GOLD, "۵", "ادغام بالینی", "سن و سابقه و علائم")
    box(1220, 360, 230, 130, "#2F6F8F", "۶", "خدمت بیمارستانی", "سامانه پشتیبان تصمیم")
    box(1510, 360, 150, 130, C_TEAL_DARK, "۷", "پزشک", "تأیید نهایی")

    # 1 → split to 2 and 3
    draw.line([(240, 425), (270, 425), (270, 245), (295, 245)], fill=hex_rgb(C_MUTED), width=3)
    draw.polygon([(295, 245), (283, 238), (283, 252)], fill=hex_rgb(C_MUTED))
    draw.line([(270, 425), (270, 605), (295, 605)], fill=hex_rgb(C_MUTED), width=3)
    draw.polygon([(295, 605), (283, 598), (283, 612)], fill=hex_rgb(C_MUTED))
    # 2 & 3 → 4 YOLO
    draw.line([(520, 245), (560, 245), (560, 425), (595, 425)], fill=hex_rgb(C_MUTED), width=3)
    draw.line([(520, 605), (560, 605), (560, 425)], fill=hex_rgb(C_MUTED), width=3)
    draw.polygon([(595, 425), (583, 418), (583, 432)], fill=hex_rgb(C_MUTED))
    arrow(860, 425, 925, 425)
    arrow(1160, 425, 1215, 425)
    arrow(1450, 425, 1505, 425)

    draw.rounded_rectangle([100, 780, 1600, 870], radius=16, fill=hex_rgb(C_TEAL_LIGHT), outline=hex_rgb(C_LINE), width=2)
    draw_rtl(
        draw, (W / 2, 825),
        "هسته اصلی جزء ۴ است؛ همه مراحل روی شبکه داخلی بیمارستان یا دانشگاه اجرا می‌شود",
        load_font(FONT_MED, 22), hex_rgb(C_TEAL_DARK), "mm",
    )
    img.save(path, quality=95)


def chart_privacy(path: Path):
    W, H = 1500, 780
    img = Image.new("RGB", (W, H), hex_rgb(C_BG))
    draw = ImageDraw.Draw(img)
    draw_rtl(draw, (W - 50, 40), "مقایسه کیفی آموزش محلی و ابری", load_font(FONT_BOLD, 34), hex_rgb(C_TEAL_DARK), "rt")

    cats = ["۱. حریم خصوصی", "۲. تناسب با اسلاید کامل", "۳. ادغام بالینی", "۴. کنترل هزینه", "۵. پذیرش بیمارستانی"]
    local = [5, 5, 5, 4, 5]
    online = [2, 2, 1, 2, 2]
    left, right, top, row_h = 450, 1280, 120, 100

    for i, cat in enumerate(cats):
        y = top + i * row_h
        draw_rtl(draw, (left - 30, y + 35), cat, load_font(FONT_MED, 24), hex_rgb(C_INK), "rt")
        for s in range(1, 6):
            x = left + s / 5 * (right - left)
            draw.line([(x, y), (x, y + 70)], fill=hex_rgb(C_LINE), width=1)
        lw = local[i] / 5 * (right - left)
        ow = online[i] / 5 * (right - left)
        draw.rounded_rectangle([left, y + 8, left + lw, y + 32], radius=8, fill=hex_rgb(C_TEAL))
        draw.rounded_rectangle([left, y + 40, left + ow, y + 64], radius=8, fill=hex_rgb(C_CORAL))

    draw.rounded_rectangle([80, 680, 160, 710], radius=6, fill=hex_rgb(C_TEAL))
    draw_rtl(draw, (180, 695), "پشته محلی با یولو", load_font(FONT_REG, 20), hex_rgb(C_INK), "rm")
    draw.rounded_rectangle([420, 680, 500, 710], radius=6, fill=hex_rgb(C_CORAL))
    draw_rtl(draw, (520, 695), "سکوی ابری خارجی", load_font(FONT_REG, 20), hex_rgb(C_INK), "rm")
    for s in range(1, 6):
        x = left + s / 5 * (right - left)
        draw_rtl(draw, (x, top + len(cats) * row_h + 10), to_fa(s), load_font(FONT_REG, 20), hex_rgb(C_MUTED), "mm")
    draw_rtl(draw, ((left + right) / 2, H - 35), "امتیاز تناسب از یک تا پنج", load_font(FONT_REG, 20), hex_rgb(C_MUTED), "mm")
    img.save(path, quality=95)


# ---------------------------------------------------------------------------
# DOCX RTL helpers — fixed
# ---------------------------------------------------------------------------

def ensure_cs_rtl_on_run(run, size_pt: float, bold: bool = False, color: str | None = None):
    """Force complex-script RTL on every run (critical for Word RTL)."""
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color.replace("#", ""))

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_NAME)

    # RTL mark on run
    for tag in ("w:rtl", "w:cs", "w:bCs", "w:sz", "w:szCs"):
        for old in rPr.findall(qn(tag)):
            rPr.remove(old)

    rtl = OxmlElement("w:rtl")
    rtl.set(qn("w:val"), "1")
    rPr.append(rtl)

    # complex script flag
    cs = OxmlElement("w:cs")
    rPr.append(cs)

    if bold:
        b = OxmlElement("w:b")
        rPr.append(b)
        bCs = OxmlElement("w:bCs")
        rPr.append(bCs)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(szCs)


def set_paragraph_rtl(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT):
    paragraph.alignment = align
    pPr = paragraph._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:bidi")):
        pPr.remove(old)
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)
    # text direction
    for old in pPr.findall(qn("w:textDirection")):
        pPr.remove(old)


def add_p(doc, text, size=11, bold=False, color=None, after=8, before=0,
          align=WD_ALIGN_PARAGRAPH.RIGHT, indent=None):
    p = doc.add_paragraph()
    set_paragraph_rtl(p, align=align)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if indent is not None:
        p.paragraph_format.first_line_indent = Cm(indent)
    run = p.add_run(text)
    ensure_cs_rtl_on_run(run, size, bold=bold, color=color)
    return p


def add_h(doc, text, level=1):
    sizes = {1: 16, 2: 13, 3: 12}
    colors = {1: C_TEAL_DARK, 2: C_TEAL, 3: C_TEAL}
    p = add_p(doc, text, size=sizes[level], bold=True, color=colors[level], before=14, after=6)
    if level == 1:
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "4")
        bottom.set(qn("w:color"), C_GOLD.replace("#", ""))
        pBdr.append(bottom)
        pPr.append(pBdr)
    return p


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color.replace("#", ""))
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell(cell, text, bold=False, size=10, color=C_INK, fill=None, center=True):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if fill:
        shade(cell, fill)
    p = cell.paragraphs[0]
    set_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.RIGHT)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    run = p.add_run(text)
    ensure_cs_rtl_on_run(run, size, bold=bold, color=color)


def set_table_rtl(table):
    """Mark table as RTL visually in Word."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:bidiVisual")):
        tblPr.remove(old)
    bidi = OxmlElement("w:bidiVisual")
    bidi.set(qn("w:val"), "1")
    tblPr.append(bidi)


def add_table(doc, title: str, headers: list[str], rows: list[list[str]], col_widths=None):
    """
    Numbered table caption + rows.
    Do NOT reverse columns: with bidiVisual, Word shows col0 on the right (RTL reading start).
    First header = rightmost column = first in Persian reading order.
    """
    TABLE_COUNTER["n"] += 1
    n = TABLE_COUNTER["n"]
    add_p(doc, f"جدول {to_fa(n)}. {title}", size=11, bold=True, color=C_TEAL_DARK, before=10, after=4)

    # Prepend row number column as first reading-order column
    headers_full = ["ردیف"] + headers
    rows_full = [[to_fa(i + 1)] + list(row) for i, row in enumerate(rows)]

    table = doc.add_table(rows=1 + len(rows_full), cols=len(headers_full))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_rtl(table)

    for j, h in enumerate(headers_full):
        set_cell(table.rows[0].cells[j], h, bold=True, size=10, color=C_WHITE, fill=C_TEAL_DARK)
    for i, row in enumerate(rows_full):
        fill = C_TEAL_LIGHT if i % 2 == 0 else C_WHITE
        for j, val in enumerate(row):
            set_cell(table.rows[i + 1].cells[j], str(val), size=9.5, fill=fill)

    if col_widths:
        widths = [1.4] + col_widths  # row number col
        # pad/truncate
        while len(widths) < len(headers_full):
            widths.append(3)
        widths = widths[: len(headers_full)]
        for row in table.rows:
            for idx, w in enumerate(widths):
                row.cells[idx].width = Cm(w)

    doc.add_paragraph()
    return table


def configure_document_rtl(doc: Document):
    """Document-level RTL + Persian language."""
    section = doc.sections[0]
    sectPr = section._sectPr
    for old in sectPr.findall(qn("w:bidi")):
        sectPr.remove(old)
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    sectPr.append(bidi)

    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.right_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.8)

    # settings: themeFontLang bidi fa-IR
    settings = doc.settings.element
    for old in settings.findall(qn("w:themeFontLang")):
        settings.remove(old)
    lang = OxmlElement("w:themeFontLang")
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "en-US")
    lang.set(qn("w:bidi"), "fa-IR")
    settings.append(lang)


def add_footer(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    set_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run("صفحه ")
    ensure_cs_rtl_on_run(run, 9, color=C_MUTED)
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
    run2 = p.add_run()
    run2._r.append(fld1); run2._r.append(instr); run2._r.append(fld2)
    ensure_cs_rtl_on_run(run2, 9, color=C_MUTED)


def add_pic(doc, path, width_cm=16):
    p = doc.add_paragraph()
    set_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(str(path), width=Cm(width_cm))
    p.paragraph_format.space_after = Pt(10)


def add_line(doc, color=C_GOLD):
    p = doc.add_paragraph()
    set_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.CENTER)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color.replace("#", ""))
    pBdr.append(bottom)
    pPr.append(pBdr)


def package_fonts():
    DOC_FONTS.mkdir(parents=True, exist_ok=True)
    for f in [FONT_REG, FONT_BOLD, FONT_MED, FONT_LIGHT]:
        shutil.copy2(f, DOC_FONTS / f.name)


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build():
    TABLE_COUNTER["n"] = 0
    ASSETS.mkdir(parents=True, exist_ok=True)
    package_fonts()

    banner = ASSETS / "cover_banner.png"
    c_timeline = ASSETS / "chart_timeline.png"
    c_cost = ASSETS / "chart_cost.png"
    c_stack = ASSETS / "chart_stack.png"
    c_privacy = ASSETS / "chart_privacy.png"

    make_cover_banner(banner)
    chart_timeline(c_timeline)
    chart_cost(c_cost)
    chart_stack(c_stack)
    chart_privacy(c_privacy)

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = Pt(11)
    rPr = style._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT_NAME)

    configure_document_rtl(doc)
    add_footer(doc.sections[0])

    # Cover
    add_pic(doc, banner, 17)
    meta_rows = [
        ["حوزه فناوری", "تجهیزات پزشکی و دارو و سلامت به‌همراه فناوری اطلاعات"],
        ["مدت اجرا", "دوازده ماه"],
        ["خروجی اصلی", "نمونه محصول و مجموعه‌داده بومی"],
        ["رویکرد فنی", "آموزش و استقرار کاملاً محلی"],
        ["هسته پشته", "مدل یولو مبتنی بر پایتون"],
    ]
    add_table(doc, "شناسنامه طرح روی جلد", ["عنوان", "شرح"], meta_rows, col_widths=[5, 10])
    add_p(doc, "نسخه فارسی رسمی با جداول شماره‌دار و نمودارها", size=10, color=C_MUTED,
          align=WD_ALIGN_PARAGRAPH.CENTER, before=8)
    add_p(doc, "قلم وزیرمتن · چینش راست‌به‌چپ · هسته یولو · محرمانگی داده بیمار",
          size=10, bold=True, color=C_TEAL, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
    doc.add_page_break()

    # TOC
    add_h(doc, "فهرست مطالب", 1)
    for item in [
        "۱. درک سند و الزامات طرح",
        "۲. مشخصات کلی طرح",
        "۳. تعریف مسأله و اهداف",
        "۴. مسیر اجرایی پیشنهادی",
        "۵. پشته فناوری با هسته یولو",
        "۶. روش آموزش مدل به‌صورت محلی",
        "۷. مقایسه آموزش محلی و ابری",
        "۸. ذخیره‌سازی و محرمانگی داده",
        "۹. شیوه همکاری با بیمارستان‌ها",
        "۱۰. زمان‌بندی دوازده‌ماهه",
        "۱۱. چالش‌ها، دستاوردها و تجاری‌سازی",
        "۱۲. جمع‌بندی و توصیه نهایی",
    ]:
        add_p(doc, item, after=4)
    doc.add_page_break()

    # 1
    add_h(doc, "۱. درک سند و الزامات طرح", 1)
    add_p(
        doc,
        "سند مبنا یک فرم درخواست طرح توسعه فناوری است. هدف، ساخت سامانه پشتیبان تصمیم‌گیری بالینی "
        "برای غربالگری سرطان دهانه رحم بر پایه تصاویر تمام‌اسلاید پاپ اسمیر و اطلاعات بالینی بیماران ایرانی است. "
        "در این نسخه، هسته یادگیری ماشین بر مدل یولو در محیط پایتون قرار گرفته تا تشخیص و طبقه‌بندی "
        "سلول‌های مشکوک در یک جریان واحد و قابل نگهداری انجام شود.",
        indent=0.5,
    )
    add_table(
        doc,
        "الزامات سند و برداشت اجرایی",
        ["الزام سند", "برداشت اجرایی"],
        [
            ["تحلیل مستقیم اسلاید کامل", "برش خودکار قطعات تصویر بدون برش دستی سلول"],
            ["طبقه‌بندی نرمال و غیرنرمال", "مدل یولو برای مکان‌یابی و طبقه سلول"],
            ["ادغام داده بالینی", "سن، سابقه درمان و علائم در تصمیم نهایی"],
            ["مجموعه‌داده بومی", "نمونه‌های واقعی ایرانی با استانداردسازی"],
            ["ارزیابی دقت و حساسیت", "شاخص‌های کمی به‌همراه پایلوت بالینی"],
            ["نمونه محصول و دادگان", "نمونه قابل استقرار و دیتاست نسخه‌دار"],
            ["چالش محرمانگی", "آموزش و استنتاج کاملاً محلی"],
        ],
        col_widths=[7, 8],
    )

    # 2
    add_h(doc, "۲. مشخصات کلی طرح", 1)
    add_h(doc, "عنوان طرح", 2)
    add_p(
        doc,
        "توسعه سامانه هوشمند شناسایی و طبقه‌بندی ناهنجاری‌های سیتولوژیک در نمونه‌های پاپ اسمیر "
        "به‌منظور تشخیص سرطان دهانه رحم در بانوان با کمک هوش مصنوعی",
        bold=True,
    )
    add_h(doc, "خلاصه دقیق طرح", 2)
    add_p(
        doc,
        "سرطان دهانه رحم از شایع‌ترین سرطان‌های قابل پیشگیری در زنان است و تشخیص زودهنگام "
        "نقش تعیین‌کننده‌ای در کاهش مرگ‌ومیر و هزینه درمان دارد. تفسیر پاپ اسمیر وابسته به نیروی "
        "متخصص است و با خطای انسانی، زمان‌بر بودن و کمبود پاتولوژیست همراه است. این طرح یک سامانه "
        "بومی می‌سازد که تصاویر تمام‌اسلاید را می‌خواند، با مدل یولو سلول‌ها را مکان‌یابی و "
        "به‌عنوان نرمال یا غیرنرمال طبقه‌بندی می‌کند، و اطلاعات بالینی را در امتیاز نهایی ادغام "
        "می‌نماید. آموزش روی زیرساخت داخلی انجام می‌شود تا داده بیمار از شبکه مرکز درمانی خارج نشود.",
        indent=0.5,
    )
    add_table(
        doc,
        "فیلدهای اصلی فرم",
        ["فیلد", "مقدار"],
        [
            ["کلمات کلیدی", "پاپ اسمیر دیجیتال، هوش مصنوعی، یولو، سرطان دهانه رحم، پاتولوژی دیجیتال"],
            ["مدت اجرا", "دوازده ماه"],
            ["حوزه فناوری", "تجهیزات پزشکی و سلامت به‌همراه فناوری اطلاعات"],
            ["کاربرد نتایج", "تولید نمونه محصول و تهیه دادگان"],
            ["شیوه دانش فنی", "توسعه و تحقیق داخلی"],
            ["بازار هدف", "ملی و بین‌المللی"],
        ],
        col_widths=[5, 10.5],
    )

    # 3
    add_h(doc, "۳. تعریف مسأله و اهداف", 1)
    add_p(
        doc,
        "مسأله اصلی، نبود سامانه بومی و چندوجهی است که بتواند تصویر پاپ اسمیر را با داده بالینی "
        "ترکیب کند، روی زیرساخت محلی آموزش ببیند و در مراکز درمانی به‌کار رود. بسیاری از سکوهای "
        "خارجی ابری‌اند، با شرایط جمعیتی کشور منطبق نیستند و محرمانگی را تهدید می‌کنند.",
        indent=0.5,
    )
    add_h(doc, "اهداف پنج‌گانه طرح", 2)
    for i, g in enumerate([
        "طراحی و پیاده‌سازی الگوریتم هوشمند مبتنی بر یولو برای تحلیل تصاویر پاپ اسمیر دیجیتال",
        "توسعه مدل تشخیص و طبقه‌بندی سلول‌های نرمال و غیرنرمال دهانه رحم",
        "ادغام داده‌های تصویری و اطلاعات بالینی بیماران در فرآیند تصمیم‌گیری",
        "ایجاد مجموعه‌داده بومی استاندارد از تصاویر پاپ اسمیر و داده‌های بالینی مرتبط",
        "ارزیابی عملکرد سامانه از نظر دقت، حساسیت و قابلیت استفاده بالینی",
    ], 1):
        add_p(doc, f"{to_fa(i)}. {g}", after=3)

    # 4
    add_h(doc, "۴. مسیر اجرایی پیشنهادی", 1)
    for i, s in enumerate([
        "انعقاد تفاهم با بیمارستان یا دانشگاه علوم پزشکی و اخذ مجوز اخلاق",
        "راه‌اندازی مخزن داده محلی و پروتکل بی‌نام‌سازی",
        "برچسب‌گذاری جعبه‌ای سلول‌ها در سامانه محلی برچسب و ساخت دیتاست بومی سازگار با یولو",
        "آموزش مدل یولو روی پردازنده گرافیکی داخلی و تنظیم آستانه حساسیت",
        "ادغام امتیاز یولو با متغیرهای بالینی در لایه تصمیم",
        "استقرار نمونه سامانه پشتیبان تصمیم روی شبکه داخلی بیمارستان",
        "پایلوت بالینی، بازخورد پاتولوژیست، مستندسازی و بسته تجاری‌سازی",
    ], 1):
        add_p(doc, f"{to_fa(i)}) {s}", after=3)

    # 5 Stack
    add_h(doc, "۵. پشته فناوری با هسته یولو", 1)
    add_p(
        doc,
        "پشته برای سادگی نگهداری فشرده شده و هسته آن جزء ۴ یعنی مدل یولو در پایتون است. "
        "اجزای خواندن اسلاید، برچسب‌گذاری و ذخیره امن ورودی مدل را آماده می‌کنند و ادغام بالینی "
        "و خدمت بیمارستانی خروجی را به پزشک می‌رسانند.",
        indent=0.5,
    )
    add_table(
        doc,
        "اجزای پشته فناوری فشرده",
        ["نام جزء", "نقش اصلی", "خروجی"],
        [
            ["خواندن اسلاید کامل", "بازکردن پرونده اسلاید و برش قطعات", "قطعات تصویر با مختصات"],
            ["سامانه برچسب‌گذاری محلی", "ترسیم جعبه توسط پاتولوژیست", "برچسب سازگار با آموزش یولو"],
            ["ذخیره امن محلی", "نگهداری اسلاید و برچسب و بالینی", "مخزن نسخه‌دار مرکز"],
            ["هسته یولو در پایتون", "تشخیص مکان سلول و طبقه نرمال یا غیرنرمال", "جعبه‌ها و احتمال طبقه"],
            ["ادغام اطلاعات بالینی", "ترکیب امتیاز تصویر با سن و سابقه و علائم", "امتیاز خطر کیس"],
            ["خدمت سامانه پشتیبان", "ارائه نتیجه روی شبکه داخلی", "پیشنهاد برای تأیید پزشک"],
        ],
        col_widths=[5, 6, 4.5],
    )
    add_p(doc, "نمودار جریان پشته:", bold=True, before=4)
    add_pic(doc, c_stack, 16.2)
    add_h(doc, "توضیح اجزا", 2)
    add_p(
        doc,
        "جزء ۱ اسلایدهای بسیار بزرگ را به قطعات قابل آموزش تبدیل می‌کند. جزء ۲ برچسب جعبه‌ای "
        "تولید می‌کند و جزء ۳ داده را امن نگه می‌دارد. جزء ۴ با یک مدل یولو هم‌زمان مکان سلول "
        "مشکوک و طبقه آن را می‌آموزد و این کار را ساده‌تر از زنجیره جداگانه تقسیم‌بندی و "
        "طبقه‌بندی نگه می‌دارد. جزء ۵ الزام فرم برای داده بالینی را پوشش می‌دهد و جزء ۶ نمونه "
        "محصول قابل استقرار در بیمارستان است.",
        indent=0.5,
    )

    # 6 Training
    add_h(doc, "۶. روش آموزش مدل به‌صورت محلی", 1)
    add_table(
        doc,
        "مراحل آموزش محلی",
        ["مرحله", "اقدام", "خروجی"],
        [
            ["آماده‌سازی قطعات", "برش اسلاید و پالایش کیفیت", "مجموعه قطعات آموزش"],
            ["برچسب جعبه‌ای", "ترسیم سلول نرمال و غیرنرمال", "پرونده برچسب یولو"],
            ["آموزش یولو", "اجرا روی پردازنده گرافیکی داخلی", "وزن مدل منتخب"],
            ["ادغام بالینی", "ترکیب امتیاز مدل با جدول بالینی", "امتیاز خطر بیمار"],
            ["ارزیابی", "سنجش دقت و حساسیت روی داده نگه‌داشته", "گزارش شاخص‌ها"],
        ],
        col_widths=[4, 6.5, 5],
    )
    add_p(doc, "اصول آموزش:", bold=True)
    for t in [
        "تقسیم داده بر اساس شناسه بیمار برای جلوگیری از نشت بین مجموعه‌ها",
        "افزایش داده با چرخش و تغییر ظاهر رنگ‌آمیزی برای مقاومت به تنوع اسکنر",
        "اولویت به حساسیت در غربالگری و تنظیم آستانه با نظر پاتولوژیست",
        "آموزش فقط روی ایستگاه داخلی؛ بدون ارسال تصویر بیمار به سکوی ابری",
        "نسخه‌گذاری وزن مدل و ثبت شاخص‌ها در گزارش‌های داخلی طرح",
    ]:
        add_p(doc, f"• {t}", after=3)

    # 7 Compare
    add_h(doc, "۷. مقایسه آموزش محلی و ابری", 1)
    add_p(
        doc,
        "سکوهای ابری خارجی برای آزمایش روی داده عمومی مناسب‌اند، اما برای تصاویر و اطلاعات "
        "بالینی بیماران این طرح توصیه نمی‌شوند. پشته محلی با هسته یولو از نظر محرمانگی، هزینه "
        "و تناسب با اسلاید کامل برتری دارد.",
        indent=0.5,
    )
    add_table(
        doc,
        "مقایسه رویکرد محلی و ابری",
        ["معیار", "پشته محلی با یولو", "سکوی ابری خارجی"],
        [
            ["حریم خصوصی", "داده در شبکه داخلی می‌ماند", "وابسته به ابر فروشنده"],
            ["تناسب با اسلاید کامل", "با برش محلی و یولو کنترل می‌شود", "غالباً برای تصویر معمولی طراحی شده"],
            ["ادغام بالینی", "لایه جدا و قابل سفارشی‌سازی", "پشتیبانی ضعیف یا ناممکن"],
            ["هزینه", "خرید یک‌باره سخت‌افزار", "اشتراک ماهانه و اعتبار مصرفی"],
            ["تکرار آزمایش", "نامحدود روی دستگاه داخلی", "محدود به اعتبار و سقف طرح"],
            ["پذیرش فناوری اطلاعات", "قابل ایزوله در بیمارستان", "اغلب برای داده بیمار مسدود"],
        ],
        col_widths=[4, 6, 6],
    )
    add_p(doc, "نمودار مقایسه کیفی:", bold=True, before=4)
    add_pic(doc, c_privacy, 15.5)
    add_p(doc, "نمودار هزینه تقریبی سال اول:", bold=True)
    add_pic(doc, c_cost, 14.5)
    add_p(
        doc,
        "نتیجه: آموزش و استنتاج محلی با هسته یولو انتخاب اصلی طرح است. استفاده از سکوی ابری "
        "فقط برای نمایش آموزشی روی داده عمومی مجاز است و داده بیمار نباید به آن وارد شود.",
        indent=0.5,
    )

    # 8 Storage
    add_h(doc, "۸. ذخیره‌سازی و محرمانگی داده", 1)
    add_table(
        doc,
        "ساختار مخزن داده محلی",
        ["سطل داده", "محتوا", "قالب پیشنهادی"],
        [
            ["اسلاید خام", "پرونده اصلی اسکنر", "قالب‌های رایج اسلاید دیجیتال"],
            ["قطعات تصویر", "برش‌ها و مختصات", "تصویر و جدول مختصات"],
            ["برچسب‌ها", "جعبه‌های سلول", "قالب آموزش یولو"],
            ["بالینی", "فیلدهای بی‌نام بیمار", "جدول ساخت‌یافته"],
            ["مدل‌ها", "وزن و نسخه", "بایگانی نسخه‌دار"],
            ["ممیزی", "لاگ مشاهده و تأیید", "پرونده الحاق‌شونده"],
        ],
        col_widths=[3.5, 6.5, 5.5],
    )
    add_p(doc, "قواعد محرمانگی:", bold=True)
    for t in [
        "عدم نگهداری نام و کد ملی و تلفن در پوشه آموزش",
        "استفاده از شناسه برگشت‌ناپذیر مطالعه",
        "جداسازی نگاشت شناسایی روی دیسک دسترسی‌محدود",
        "رمزنگاری مخزن و نقش‌بندی دسترسی",
        "ممنوعیت بارگذاری داده بیمار روی سکوهای ابری برچسب یا آموزش",
        "پشتیبان روزانه محلی و نسخه سرد هفتگی",
    ]:
        add_p(doc, f"• {t}", after=3)

    # 9 Hospitals
    add_h(doc, "۹. شیوه همکاری با بیمارستان‌ها", 1)
    add_p(
        doc,
        "اصل راهنما این است که نرم‌افزار، پروتکل، وزن مدل و بسته ارزیابی بی‌نام قابل اشتراک است؛ "
        "بایگانی خام قابل شناسایی بیماران به‌صورت عمومی منتقل نمی‌شود.",
        indent=0.5,
    )
    add_table(
        doc,
        "حالت‌های همکاری بیمارستانی",
        ["حالت همکاری", "دریافتی مرکز", "بازگشتی مرکز"],
        [
            ["شریک داده", "پروتکل برچسب و قالب پروژه", "اسلاید و بالینی بی‌نام روی مسیر امن"],
            ["سایت پایلوت", "بسته سامانه روی شبکه داخلی", "بازخورد پزشک و شاخص استفاده"],
            ["به‌روزرسانی مدل", "وزن مدل جدید", "شاخص‌های تجمیعی اختیاری"],
            ["طرح چندمرکزی آتی", "دستورعمل مشترک", "خروجی بی‌نام هماهنگ"],
        ],
        col_widths=[4, 6, 6],
    )
    add_p(doc, "پیش‌نیازهای حکمرانی:", bold=True)
    for t in [
        "تأیید کمیته اخلاق",
        "توافق‌نامه اشتراک داده",
        "راستی‌آزمایی بی‌نام‌سازی",
        "تعریف نقش دسترسی و مسئول پاسخ رخداد",
    ]:
        add_p(doc, f"• {t}", after=3)

    # 10 Timeline
    add_h(doc, "۱۰. زمان‌بندی دوازده‌ماهه", 1)
    add_pic(doc, c_timeline, 16.2)
    add_table(
        doc,
        "فازهای اجرایی",
        ["فاز", "بازه ماه", "خروجی کلیدی"],
        [
            ["راه‌اندازی", "۱ تا ۲", "مجوز اخلاق، تفاهم‌نامه، سخت‌افزار و مخزن"],
            ["داده", "۲ تا ۵", "دیتاست بومی نسخه یک با برچسب یولو"],
            ["مدل", "۴ تا ۸", "مدل یولو و لایه ادغام بالینی با شاخص قابل قبول"],
            ["سامانه", "۷ تا ۱۰", "نمونه سامانه پشتیبان روی شبکه داخلی"],
            ["پایلوت", "۹ تا ۱۱", "گزارش قابلیت استفاده بالینی"],
            ["اختتام", "۱۱ تا ۱۲", "مستندات، کارت دیتاست و بسته تجاری‌سازی"],
        ],
        col_widths=[3.5, 3.5, 9],
    )
    add_table(
        doc,
        "تناظر اهداف فرم با زمان تحقق",
        ["هدف فرم", "ماه تحقق"],
        [
            ["الگوریتم تحلیل تصویر مبتنی بر یولو", "ماه ۶"],
            ["مدل نرمال و غیرنرمال", "ماه ۸"],
            ["ادغام بالینی", "ماه ۹"],
            ["مجموعه‌داده بومی", "ماه ۵ و ماه ۱۱"],
            ["ارزیابی دقت و حساسیت و کاربری", "ماه‌های ۱۰ تا ۱۲"],
        ],
        col_widths=[10, 5.5],
    )

    # 11
    add_h(doc, "۱۱. چالش‌ها، دستاوردها و تجاری‌سازی", 1)
    add_table(
        doc,
        "چالش‌ها و راهکارها",
        ["چالش", "راهکار"],
        [
            ["کمبود دادگان بومی", "همکاری با دانشگاه علوم پزشکی و چند مرکز"],
            ["تنوع کیفیت لام", "پروتکل استاندارد و افزایش داده و مدل مقاوم"],
            ["محرمانگی بیمار", "بی‌نام‌سازی و آموزش محلی و کنترل دسترسی"],
            ["عدم توازن طبقه", "وزن‌دهی خطا و آستانه حساسیت‌محور در یولو"],
            ["پیچیدگی نگهداری", "پشته فشرده با هسته واحد یولو"],
        ],
        col_widths=[6, 10],
    )
    add_h(doc, "سایر دستاوردها", 2)
    for t in [
        "انتشار مقالات علمی در مجلات و کنفرانس‌های معتبر",
        "ایجاد مجموعه‌داده بومی قابل استفاده در پژوهش‌های آینده",
        "توسعه نمونه اولیه سامانه تشخیص هوشمند با استقرار محلی",
        "زمینه‌سازی برای ثبت اختراع یا تجاری‌سازی محصول",
        "ارتقاء کیفیت و هوشمندسازی فرآیندهای تشخیصی موجود در کشور",
    ]:
        add_p(doc, f"• {t}", after=3)

    add_h(doc, "آمادگی تجاری‌سازی", 2)
    add_table(
        doc,
        "وضعیت تجاری‌سازی",
        ["عنوان", "پاسخ"],
        [
            ["دستاورد نهایی", "تولید فناوری"],
            ["تمایل به تجاری‌سازی", "توسط تیم مجری"],
            ["بهره‌بردار اصلی", "تمامی بانوان در جمعیت غربالگری"],
            ["نهادهای استفاده‌کننده", "بیمارستان‌ها و مراکز درمانی دولتی و خصوصی"],
            ["حمایت مورد نیاز", "همکاری دانشگاه علوم پزشکی استان برای جمع‌آوری داده"],
            ["گام پس از اتمام", "تجاری‌سازی برای فراگیری در کشور"],
        ],
        col_widths=[5.5, 10.5],
    )

    # 12
    add_h(doc, "۱۲. جمع‌بندی و توصیه نهایی", 1)
    add_p(
        doc,
        "این پیشنهادیه در قالب فرم توسعه فناوری و با محوریت مدل یولو در پایتون تنظیم شده است: "
        "خواندن اسلاید کامل، برچسب محلی، آموزش یولو روی پردازنده گرافیکی داخلی، ادغام بالینی، "
        "و خدمت سامانه پشتیبان روی شبکه بیمارستان. این مسیر با پنج هدف سند هم‌راستاست و محرمانگی "
        "داده بیمار را حفظ می‌کند.",
        indent=0.5,
    )
    add_p(doc, "جریان نهایی پیشنهادی:", bold=True, before=8)
    add_p(
        doc,
        "اسلاید محلی ← برش تصویر ← برچسب جعبه‌ای ← آموزش یولو ← ادغام بالینی ← سامانه پشتیبان ← تأیید پاتولوژیست",
        align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, color=C_TEAL_DARK, size=10,
    )
    add_line(doc)
    add_p(
        doc,
        "پایان پیشنهادیه — نسخه فارسی با قلم وزیرمتن و چینش راست‌به‌چپ اصلاح‌شده",
        size=9, color=C_MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, before=10,
    )

    OUT_DOCX.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT_DOCX))
    print(f"Wrote {OUT_DOCX} with {TABLE_COUNTER['n']} tables")
    return OUT_DOCX


if __name__ == "__main__":
    build()
