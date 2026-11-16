# -*- coding: utf-8 -*-
"""
Обновление отдельного файла расписок по первоначальному взносу:
Файл: Расписка_281400_Первоначальный_взнос_Коняев_Ситников.docx
Лист 1: Расписка в получении денежных средств (первоначальный взнос) на сумму 281 400 руб. от 12.11.2026 г.
Лист 2: Возвратная расписка на эту же сумму (281 400 руб.) с открытой датой «___» ____________ 2026 г.
(Коняев вернул Ситникову 281 400 руб., Ситников деньги получил, претензий нет).
Сохраняется на Рабочий стол.
"""

import os
import sys
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.stdout.reconfigure(encoding="utf-8")

def set_run_ru(run, font_name="Times New Roman", size_pt=10.5, bold=False, italic=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    if not rPr.xpath("./w:lang"):
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="ru-RU" w:bidi="ru-RU"/>'))

def add_p(doc, text="", bold_prefix="", size_pt=10.5, bold=False, italic=False,
          align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0.0, space_after=1.5,
          line_spacing=1.08, indent=21.25, keep_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent and align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        p.paragraph_format.first_line_indent = Pt(indent)
    else:
        p.paragraph_format.first_line_indent = Pt(0)
    if keep_next:
        p.paragraph_format.keep_with_next = True

    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        set_run_ru(r_pre, font_name="Times New Roman", size_pt=size_pt, bold=True)
    if text:
        r_txt = p.add_run(text)
        set_run_ru(r_txt, font_name="Times New Roman", size_pt=size_pt, bold=bold, italic=italic)
    return p

def add_p_parts(doc, parts, size_pt=10.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                space_before=0.0, space_after=1.5, line_spacing=1.08, indent=21.25, keep_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent and align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        p.paragraph_format.first_line_indent = Pt(indent)
    else:
        p.paragraph_format.first_line_indent = Pt(0)
    if keep_next:
        p.paragraph_format.keep_with_next = True

    for text, is_bold in parts:
        r = p.add_run(text)
        set_run_ru(r, font_name="Times New Roman", size_pt=size_pt, bold=is_bold)
    return p

def add_h(doc, title_text, space_before=2.0, space_after=1.5, size_pt=11.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.08
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(title_text)
    set_run_ru(r, font_name="Times New Roman", size_pt=size_pt, bold=True)
    return p

def add_city_date(doc, city_text="г. Пенза", date_text="12 ноября 2026 г.", size_pt=10.5):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    c1, c2 = tbl.cell(0, 0), tbl.cell(0, 1)
    c1.width = Pt(240)
    c2.width = Pt(255)
    for cell in (c1, c2):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(f"""
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:left w:val="none"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        """)
        tcPr.append(tcBorders)

    p_l = c1.paragraphs[0]
    p_l.paragraph_format.space_before = Pt(2)
    p_l.paragraph_format.space_after = Pt(3)
    r_l = p_l.add_run(city_text)
    set_run_ru(r_l, font_name="Times New Roman", size_pt=size_pt, bold=True)

    p_r = c2.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(2)
    p_r.paragraph_format.space_after = Pt(3)
    r_r = p_r.add_run(date_text)
    set_run_ru(r_r, font_name="Times New Roman", size_pt=size_pt, bold=True)

def add_sig_block(doc, title_text, keep_all=True):
    add_p(doc, title_text, size_pt=10.5, bold=True, indent=0, space_after=1.5, keep_next=True)

    p_line = doc.add_paragraph()
    pPr = p_line._p.get_or_add_pPr()
    pBdr = parse_xml(f"""
    <w:pBdr {nsdecls("w")}>
        <w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>
    </w:pBdr>
    """)
    pPr.append(pBdr)
    p_line.paragraph_format.space_after = Pt(1.0)
    p_line.paragraph_format.keep_with_next = True
    r_empty = p_line.add_run(" ")
    set_run_ru(r_empty, size_pt=9.0)

    p_sub = doc.add_paragraph()
    pPr_sub = p_sub._p.get_or_add_pPr()
    tabs = parse_xml(f"""
    <w:tabs {nsdecls("w")}>
        <w:tab w:val="right" w:pos="9921"/>
    </w:tabs>
    """)
    pPr_sub.append(tabs)
    p_sub.paragraph_format.space_after = Pt(4.0 if keep_all else 1.0)
    p_sub.paragraph_format.keep_with_next = True

    r1 = p_sub.add_run("ФИО полностью")
    set_run_ru(r1, font_name="Times New Roman", size_pt=8.5, italic=False)
    r2 = p_sub.add_run("\tподпись")
    set_run_ru(r2, font_name="Times New Roman", size_pt=8.5, italic=False)

def generate_receipts_file(desktop_dir):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.65)
        s.bottom_margin = Inches(0.65)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.55)
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)

    # =========================================================================
    # ЛИСТ 1: РАСПИСКА В ПОЛУЧЕНИИ ДЕНЕЖНЫХ СРЕДСТВ (ПЕРВОНАЧАЛЬНЫЙ ВЗНОС)
    # =========================================================================
    add_h(doc, "РАСПИСКА", space_before=2.0, space_after=2.0, size_pt=12.0)
    add_h(doc, "в получении денежных средств (первоначальный взнос)", space_before=0, space_after=6.0, size_pt=10.5)
    add_city_date(doc, "г. Пенза", "12 ноября 2026 г.", size_pt=10.5)

    add_p_parts(doc, [
        ("Я, ", False),
        ("гр. РФ Коняев Петр Николаевич", True),
        (", 10.04.1960 года рождения, место рождения: село Мичкасс, Никольского р-на, Пензенской обл., "
         "паспорт гражданина Российской Федерации 5605 № 497454, выдан ОТДЕЛОМ ВНУТРЕННИХ ДЕЛ ЖЕЛЕЗНОДОРОЖНОГО РАЙОНА Г. ПЕНЗЫ, "
         "дата выдачи 29.06.2005, код подразделения 582-004, зарегистрированный по адресу: г. Пенза, ул. Кл. Цеткин, д. 23, кв. 125, "
         "настоящей распиской подтверждаю, что лично получил от", False)
    ], size_pt=11.0, space_after=2.5)

    add_p_parts(doc, [
        ("гр. РФ Ситникова Алексея Сергеевича", True),
        (", 26.04.1995 года рождения, место рождения: с. Дубровское, Лопатинского р-на, Пензенской обл., "
         "паспорт гражданина Российской Федерации 5614 № 264654, выдан МЕЖРАЙОННЫМ ОТДЕЛЕНИЕМ УФМС РОССИИ ПО ПЕНЗЕНСКОЙ ОБЛАСТИ "
         "В ШЕМЫШЕЙСКОМ РАЙОНЕ, дата выдачи 27.05.2015, код подразделения 580-033, зарегистрированного по адресу: Пензенская обл., "
         "Лопатинский р-н, ул. Южная, д. 6, кв. 5,", False)
    ], size_pt=11.0, space_after=2.5)

    add_p_parts(doc, [
        ("денежные средства в размере ", False),
        ("281 400 (Двести восемьдесят одна тысяча четыреста) рублей 00 копеек", True),
        (" в качестве первоначального взноса (оплаты части стоимости за счет собственных денежных средств Покупателя) "
         "за продаваемую комнату по Договору купли-продажи комнаты от 12 ноября 2026 года, имеющую следующие технические характеристики:", False)
    ], size_pt=11.0, space_after=2.0)

    add_p(doc,
          "Комната, этаж № 4, назначение: жилое, общая площадь 12,7 кв.м., местоположение: Пензенская область, г. Пенза, ул. Медицинская, д. 6, "
          "кадастровый номер 58:29:2014005:5301.",
          size_pt=11.0, bold=True, space_after=3.5)

    add_p(doc,
          "Денежные средства в размере 281 400 (Двести восемьдесят одна тысяча четыреста) рублей переданы Покупателем и получены Продавцом лично в полном объеме в момент подписания Договора купли-продажи. "
          "Претензий по сумме и порядку передачи указанного первоначального взноса к Ситникову Алексею Сергеевичу не имею.",
          size_pt=11.0, space_after=3.0)

    add_p(doc,
          "Настоящая расписка составлена в двух, юридически равных экземплярах, по одному для каждой из сторон.",
          size_pt=11.0, space_after=6.0)

    add_sig_block(doc, "Деньги получил (Продавец):", keep_all=True)
    add_sig_block(doc, "Деньги передал (Покупатель):", keep_all=False)

    # =========================================================================
    # ЛИСТ 2: ВОЗВРАТНАЯ РАСПИСКА НА ЭТУ ЖЕ СУММУ (281 400 РУБ.) С ОТКРЫТОЙ ДАТОЙ
    # =========================================================================
    doc.add_page_break()

    add_h(doc, "ВОЗВРАТНАЯ РАСПИСКА", space_before=2.0, space_after=6.0, size_pt=12.0)
    add_city_date(doc, "г. Пенза", "«___» _______________ 2026 г.", size_pt=10.5)

    add_p_parts(doc, [
        ("Я, ", False),
        ("гр. РФ Ситников Алексей Сергеевич", True),
        (", 26.04.1995 года рождения, место рождения: с. Дубровское, Лопатинского р-на, Пензенской обл., "
         "паспорт гражданина РФ 5614 № 264654, выдан МЕЖРАЙОННЫМ ОТДЕЛЕНИЕМ УФМС РОССИИ ПО ПЕНЗЕНСКОЙ ОБЛАСТИ В ШЕМЫШЕЙСКОМ РАЙОНЕ, "
         "дата выдачи 27.05.2015, код подразделения 580-033, зарегистрированный по адресу: Пензенская обл., Лопатинский р-н, ул. Южная, д. 6, кв. 5, "
         "настоящей распиской подтверждаю, что получил от", False)
    ], size_pt=11.0, space_after=2.5)

    add_p_parts(doc, [
        ("гр. РФ Коняева Петра Николаевича", True),
        (", 10.04.1960 года рождения, место рождения: село Мичкасс, Никольского р-на, Пензенской обл., "
         "паспорт гражданина РФ 5605 № 497454, выдан ОТДЕЛОМ ВНУТРЕННИХ ДЕЛ ЖЕЛЕЗНОДОРОЖНОГО РАЙОНА Г. ПЕНЗЫ, "
         "дата выдачи 29.06.2005, код подразделения 582-004, зарегистрированного по адресу: г. Пенза, ул. Кл. Цеткин, д. 23, кв. 125, ", False),
        ("возврат денежных средств, ранее внесенных в качестве первоначального взноса, в размере 281 400 (Двести восемьдесят одна тысяча четыреста) рублей 00 копеек", True),
        (" по Договору купли-продажи комнаты от 12 ноября 2026 года, имеющей следующие технические характеристики: "
         "комната, этаж № 4, назначение: жилое, общая площадь 12,7 кв.м., местоположение: Пензенская область, г. Пенза, ул. Медицинская, д. 6, "
         "кадастровый номер 58:29:2014005:5301.", False)
    ], size_pt=11.0, space_after=3.0)

    add_p(doc,
          "Денежные средства в размере 281 400 (Двести восемьдесят одна тысяча четыреста) рублей возвращены Продавцом Коняевым П.Н. Покупателю Ситникову А.С. лично в полном объеме. "
          "Обязательство по возврату указанных денежных средств исполнено полностью. Финансовых и материальных претензий по возврату первоначального взноса к Коняеву Петру Николаевичу не имею.",
          size_pt=11.0, space_after=3.0)

    add_p(doc,
          "Настоящая возвратная расписка составлена в двух, юридически равных экземплярах, по одному для каждой из сторон.",
          size_pt=11.0, space_after=6.0)

    add_sig_block(doc, "Денежные средства передал (Продавец):", keep_all=True)
    add_sig_block(doc, "Денежные средства получил (Покупатель):", keep_all=False)

    f_out = os.path.join(desktop_dir, "Расписка_281400_Первоначальный_взнос_Коняев_Ситников.docx")
    doc.save(f_out)
    print(f"ОБНОВЛЕН ФАЙЛ НА РАБОЧЕМ СТОЛЕ: {f_out}")

if __name__ == "__main__":
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    generate_receipts_file(desktop)
