import os
import sys
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_document_defaults(doc, font_name="Times New Roman", font_size_pt=10.0):
    """Устанавливает Times New Roman и базовый кегль как системный дефолт всего документа (docDefaults и стиль Normal)."""
    style = doc.styles['Normal']
    style.font.name = font_name
    style.font.size = Pt(font_size_pt)

    rPrDefault = doc.styles.element.xpath('w:docDefaults/w:rPrDefault/w:rPr')
    if rPrDefault:
        sz_val = str(int(round(font_size_pt * 2)))
        rFonts = parse_xml(r'<w:rFonts %s w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s"/>' % (nsdecls('w'), font_name, font_name, font_name, font_name))
        rPrDefault[0].append(rFonts)
        sz = parse_xml(r'<w:sz %s w:val="%s"/>' % (nsdecls('w'), sz_val))
        szCs = parse_xml(r'<w:szCs %s w:val="%s"/>' % (nsdecls('w'), sz_val))
        rPrDefault[0].append(sz)
        rPrDefault[0].append(szCs)
        lang = parse_xml(r'<w:lang %s w:val="ru-RU"/>' % nsdecls('w'))
        rPrDefault[0].append(lang)

def add_styled_run(p, text, font_size=10.0, bold=False, italic=False, font_name="Times New Roman"):
    """Добавляет run со СТРОГО фиксированным шрифтом, кеглем и языком для всех скриптов (ASCII, Cyrillic CS, EastAsia)."""
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.name = font_name
    r.font.size = Pt(font_size)

    sz_val = str(int(round(font_size * 2)))
    rPr = r._r.get_or_add_rPr()

    for tag in ['w:rFonts', 'w:sz', 'w:szCs', 'w:lang']:
        for el in rPr.xpath(tag):
            rPr.remove(el)

    rFonts = parse_xml(r'<w:rFonts %s w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s"/>' % (nsdecls('w'), font_name, font_name, font_name, font_name))
    sz = parse_xml(r'<w:sz %s w:val="%s"/>' % (nsdecls('w'), sz_val))
    szCs = parse_xml(r'<w:szCs %s w:val="%s"/>' % (nsdecls('w'), sz_val))
    lang = parse_xml(r'<w:lang %s w:val="ru-RU"/>' % nsdecls('w'))

    rPr.append(rFonts)
    rPr.append(sz)
    rPr.append(szCs)
    rPr.append(lang)
    return r

def create_document_base(font_size=10.0):
    doc = docx.Document()
    set_document_defaults(doc, font_name="Times New Roman", font_size_pt=font_size)

    s = doc.sections[0]
    s.top_margin = Inches(0.71)    # 18 mm (51.1 pt)
    s.bottom_margin = Inches(0.71) # 18 mm (51.1 pt)
    s.left_margin = Inches(0.79)   # 20 mm (56.9 pt)
    s.right_margin = Inches(0.59)  # 15 mm (42.5 pt)
    s.page_width = Inches(8.27)    # 210 mm
    s.page_height = Inches(11.69)  # 297 mm
    return doc, s

def add_full_width_signature(doc, body_width, role_title, printed_name, font_size=10.0, line_spacing=1.06):
    """
    Создает эталонный блок подписи во всю ширину страницы с нижней границей w:pBdr
    и подстрочным текстом 'ФИО полностью' / 'подпись' по правому краю без мягких переносов.
    """
    p_t = doc.add_paragraph()
    p_t.paragraph_format.space_before = Pt(2.5)
    p_t.paragraph_format.space_after = Pt(1.0)
    p_t.paragraph_format.line_spacing = line_spacing
    p_t.paragraph_format.first_line_indent = Pt(0)
    p_t.paragraph_format.keep_with_next = True
    add_styled_run(p_t, role_title, font_size=font_size, bold=True)

    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(0)
    p_line.paragraph_format.space_after = Pt(1.0)
    p_line.paragraph_format.first_line_indent = Pt(0)
    p_line.paragraph_format.keep_with_next = True
    pPr = p_line._p.get_or_add_pPr()
    pBdr = parse_xml(r"""
        <w:pBdr %s>
            <w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>
        </w:pBdr>
    """ % nsdecls("w"))
    pPr.append(pBdr)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(1.0)
    p_sub.paragraph_format.first_line_indent = Pt(0)
    p_sub.paragraph_format.keep_with_next = True
    p_sub.paragraph_format.tab_stops.add_tab_stop(body_width, WD_TAB_ALIGNMENT.RIGHT)
    add_styled_run(p_sub, "ФИО полностью", font_size=8.0, italic=True)
    add_styled_run(p_sub, "\t", font_size=8.0)
    add_styled_run(p_sub, "подпись", font_size=8.0, italic=True)

    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(3.0)
    p_name.paragraph_format.first_line_indent = Pt(0)
    p_name.paragraph_format.keep_with_next = True
    add_styled_run(p_name, printed_name, font_size=font_size, bold=False)
