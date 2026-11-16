# -*- coding: utf-8 -*-
"""
Пересобранный точный пакет документов по сделке Коняев П.Н. -> Ситников А.С.
(г. Пенза, ул. Медицинская, д. 6, комната, к/н 58:29:2014005:5301) от 12.11.2026 г.
с учетом ВСЕХ точных цифр и правок пользователя:
- Цена в основном ДКП: 1 400 000 руб.
- Первоначальный взнос: 281 400 руб.
- Кредитные средства ПАО Сбербанк: 1 118 600 руб.
- Фактическая действительная стоимость объекта по Допнику: 940 000 руб.
- Исключение из ДКП всех пунктов о ПВ 281 400 руб. и цене 1 400 000 руб.
- Возврат Продавцом 178 600 руб. (разница 1 118 600 - 940 000) в течение 2 календарных дней.
- Штраф за нарушение конфиденциальности: 100 000 (Сто тысяч) рублей (синхронизировано в обоих экземплярах!).
- Удален второй кадастровый номер из описания по правке пользователя, добавлен оборот «имеющую следующие технические характеристики».
- Отдельная страница в Допнике: Возвратная расписка с открытой датой на 178 600 руб.
- Отдельный файл: Расписка по ПВ на 281 400 руб.

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

def add_h(doc, title_text, space_before=2.0, space_after=1.5, size_pt=11.0):
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

def build_ds_body(doc, keeper_role):
    add_h(doc, "ДОПОЛНИТЕЛЬНОЕ СОГЛАШЕНИЕ К ДОГОВОРУ КУПЛИ-ПРОДАЖИ", space_before=0, space_after=0.5)
    add_h(doc, "от 12 ноября 2026 г.", space_before=0, space_after=3.0)
    add_city_date(doc, "г. Пенза", "12 ноября 2026 г.")

    add_p(doc, "Мы, нижеподписавшиеся:", size_pt=10.5, space_after=1.5)

    add_p_parts(doc, [
        ("гр. РФ Коняев Петр Николаевич", True),
        (", 10.04.1960 года рождения, место рождения: село Мичкасс, Никольского р-на, Пензенской обл., "
         "паспорт гражданина Российской Федерации 5605 № 497454, выдан ОТДЕЛОМ ВНУТРЕННИХ ДЕЛ ЖЕЛЕЗНОДОРОЖНОГО РАЙОНА Г. ПЕНЗЫ, "
         "дата выдачи 29.06.2005, код подразделения 582-004, зарегистрированный по адресу: г. Пенза, ул. Кл. Цеткин, д. 23, кв. 125, "
         "далее по тексту «Продавец», с одной стороны, и", False)
    ], size_pt=10.5, space_after=1.5)

    add_p_parts(doc, [
        ("гр. РФ Ситников Алексей Сергеевич", True),
        (", 26.04.1995 года рождения, место рождения: с. Дубровское, Лопатинского р-на, Пензенской обл., "
         "паспорт гражданина Российской Федерации 5614 № 264654, выдан МЕЖРАЙОННЫМ ОТДЕЛЕНИЕМ УФМС РОССИИ ПО ПЕНЗЕНСКОЙ ОБЛАСТИ "
         "В ШЕМЫШЕЙСКОМ РАЙОНЕ, дата выдачи 27.05.2015, код подразделения 580-033, зарегистрированный по адресу: Пензенская обл., "
         "Лопатинский р-н, ул. Южная, д. 6, кв. 5, далее по тексту «Покупатель», с другой стороны, при совместном упоминании именуемые «Стороны»,", False)
    ], size_pt=10.5, space_after=1.5)

    add_p(doc, "заключили настоящее Дополнительное соглашение о нижеследующем:", size_pt=10.5, space_after=1.5)

    add_p_parts(doc, [
        ("1. Стороны пришли к соглашению внести изменения в Договор купли-продажи комнаты от 12.11.2026 года "
         "(далее по тексту – Договор) в отношении следующего объекта недвижимого имущества: ", False),
        ("комната, этаж № 4, назначение: жилое, общая площадь 12,7 кв.м., местоположение: Пензенская область, г. Пенза, ул. Медицинская, д. 6, "
         "кадастровый номер 58:29:2014005:5301", True),
        (", далее именуемого «Объект», в соответствии с условиями настоящего Соглашения.", False)
    ], size_pt=10.5, space_after=1.5)

    add_p(doc, "2. Содержание изменений, вносимых сторонами в Договор:", bold_prefix="", size_pt=10.5, bold=True, space_after=1.0)
    add_p(doc,
          "Стороны пришли к соглашению о внесении изменений в Договор купли-продажи от 12.11.2026 г. "
          "Изменения, затрагивающие суть договора купли-продажи от 12.11.2026 г.:",
          size_pt=10.5, space_after=1.0)

    add_p_parts(doc, [
        ("2.1. Действительная стоимость продаваемого Объекта недвижимости составляет ", False),
        ("940 000 (Девятьсот сорок тысяч) рублей", True),
        (".", False)
    ], size_pt=10.5, space_after=1.0)

    add_p(doc, "2.2. Порядок оплаты за приобретаемый Объект недвижимости:", size_pt=10.5, space_after=1.0)
    add_p_parts(doc, [
        ("Указанная стоимость оплачивается в следующем порядке: за счет кредитных денежных средств, предоставляемых ПАО «Сбербанк России», "
         "перечисляемых на счет Продавца в размере ", False),
        ("1 118 600 (Один миллион сто восемнадцать тысяч шестьсот) рублей", True),
        (". Все иные пункты Договора купли-продажи в части порядка расчетов и оплаты (включая оплату наличных собственных средств в размере 281 400 рублей "
         "и первоначальную цену 1 400 000 рублей) Сторонами исключаются (признаются утратившими силу).", False)
    ], size_pt=10.5, space_after=1.5)

    add_p_parts(doc, [
        ("2.3. Стороны дополняют Договор следующим пунктом о возврате излишне перечисленных денежных средств: ", False),
        ("«Ввиду того, что размер перечисляемых по настоящему Договору на счет Продавца кредитных денежных средств составляет 1 118 600 рублей, "
         "а действительная согласованная цена Объекта составляет 940 000 рублей, Продавец обязуется вернуть Покупателю излишне перечисленную сумму "
         "в размере 178 600 (Сто семьдесят восемь тысяч шестьсот) рублей в срок, не превышающий 2 (двух) календарных дней с момента зачисления "
         "кредитных денежных средств на счет Продавца».", True)
    ], size_pt=10.5, space_after=1.5)

    add_p(doc, "3. Условия о конфиденциальности настоящего Соглашения и Договора купли-продажи:", bold=True, size_pt=10.5, space_after=1.0)
    add_p(doc, "3.1. Стороны обязуются сохранять конфиденциальность всей информации и документов, предоставленных в связи с заключением и исполнением настоящего Соглашения.", size_pt=10.5, space_after=1.0)
    add_p(doc, "3.2. Конфиденциальная информация не может быть раскрыта третьей стороне без предварительного письменного согласия другой стороны.", size_pt=10.5, space_after=1.0)
    add_p_parts(doc, [
        ("3.3. В случае разглашения конфиденциальной информации сторона, нарушившая условие о конфиденциальности, обязана возместить другой стороне все убытки, "
         "связанные с таким разглашением, а также выплатить компенсацию пострадавшей стороне в размере ", False),
        ("100 000 (Сто тысяч) рублей.", True)
    ], size_pt=10.5, space_after=1.0)
    add_p(doc, "3.4. Условия о конфиденциальности сохраняют свою силу и после окончания срока действия как Договора, так и настоящего Соглашения, если иное не предусмотрено договором.", size_pt=10.5, space_after=1.0)
    add_p(doc, "3.5. Условия о конфиденциальности не распространяются на случаи, когда раскрытие такой информации необходимо в соответствии с запросом правоохранительных органов.", size_pt=10.5, space_after=1.0)
    add_p(doc, "3.6. Стороны, подписывая настоящее соглашение, подтверждают, что соглашение заключено добровольно, без принуждения, с взаимного согласия, сознают последствия его заключения и принимают риски разглашения его содержания.", size_pt=10.5, space_after=1.5)

    add_p(doc, "4. Порядок хранения конфиденциальной информации:", bold=True, size_pt=10.5, space_after=1.0)
    add_p(doc, "5. Стороны обязуются осуществлять хранение настоящего Соглашения и платежных документов к нему таким образом, чтобы исключить получение любым лицом (как физическим, так и юридическим) копии, скан-образа настоящего соглашения, а также расписки к нему.", size_pt=10.5, space_after=1.0)
    add_p(doc, "6. Установление факта наличия настоящего соглашения, платежных документов по исполнению условий настоящего Соглашения, у третьих лиц, включая государственных (за исключением предоставление настоящего документа по запросу суда и (или) правоохранительных органов) и иных органов, означает нарушение порядка хранения настоящего Соглашения и конфиденциальной информации со стороны того участника сделки, чей экземпляр будет раскрыт третьим лицам в нарушении условий о конфиденциальности. На экземплярах каждой стороны будет сделана пометка «Экземпляр Продавца» и «Экземпляр Покупателя».", size_pt=10.5, space_after=1.0)
    add_p(doc, "7. В остальном, не урегулированном настоящим соглашением, Стороны руководствуются положениями Догора.", size_pt=10.5, space_after=1.0)
    add_p_parts(doc, [
        ("8. Настоящее соглашение составлено в 2-х экземплярах, которые хранятся у сторон по сделке. Настоящий экземпляр передается ", True),
        (keeper_role, True),
        (".", True)
    ], size_pt=10.5, space_after=2.5)

    add_p(doc, "Подписи сторон:", size_pt=10.5, bold=True, indent=0, space_after=1.5, keep_next=True)
    add_sig_block(doc, "Продавец:", keep_all=True)
    add_sig_block(doc, "Покупатель:", keep_all=False)

def generate_all(desktop_dir):
    # =========================================================================
    # ФАЙЛ 1: ДОПСОГЛАШЕНИЕ (2 ЭКЗЕМПЛЯРА) + ВОЗВРАТНАЯ РАСПИСКА НА 178 600 РУБ.
    # =========================================================================
    doc_ds = docx.Document()
    for s in doc_ds.sections:
        s.top_margin = Inches(0.55)
        s.bottom_margin = Inches(0.55)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.55)
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)

    # Экземпляр 1: Продавцу
    build_ds_body(doc_ds, "«Продавцу»")
    doc_ds.add_page_break()

    # Экземпляр 2: Покупателю
    build_ds_body(doc_ds, "«Покупателю»")
    doc_ds.add_page_break()

    # ОТДЕЛЬНАЯ СТРАНИЦА: ВОЗВРАТНАЯ РАСПИСКА С ОТКРЫТОЙ ДАТОЙ НА 178 600 РУБ.
    add_h(doc_ds, "ВОЗВРАТНАЯ РАСПИСКА", space_before=2.0, space_after=6.0, size_pt=11.5)
    add_city_date(doc_ds, "г. Пенза", "«___» _______________ 2026 г.", size_pt=10.5)

    add_p_parts(doc_ds, [
        ("Я, ", False),
        ("гр. РФ Ситников Алексей Сергеевич", True),
        (", 26.04.1995 года рождения, место рождения: с. Дубровское, Лопатинского р-на, Пензенской обл., "
         "паспорт гражданина РФ 5614 № 264654, выдан МЕЖРАЙОННЫМ ОТДЕЛЕНИЕМ УФМС РОССИИ ПО ПЕНЗЕНСКОЙ ОБЛАСТИ В ШЕМЫШЕЙСКОМ РАЙОНЕ, "
         "дата выдачи 27.05.2015, код подразделения 580-033, зарегистрированный по адресу: Пензенская обл., Лопатинский р-н, ул. Южная, д. 6, кв. 5, "
         "настоящей распиской подтверждаю, что получил от", False)
    ], size_pt=10.5, space_after=2.0)

    add_p_parts(doc_ds, [
        ("гр. РФ Коняева Петра Николаевича", True),
        (", 10.04.1960 года рождения, место рождения: село Мичкасс, Никольского р-на, Пензенской обл., "
         "паспорт гражданина РФ 5605 № 497454, выдан ОТДЕЛОМ ВНУТРЕННИХ ДЕЛ ЖЕЛЕЗНОДОРОЖНОГО РАЙОНА Г. ПЕНЗЫ, "
         "дата выдачи 29.06.2005, код подразделения 582-004, зарегистрированного по адресу: г. Пенза, ул. Кл. Цеткин, д. 23, кв. 125, ", False),
        ("возврат излишне перечисленных денежных средств в размере 178 600 (Сто семьдесят восемь тысяч шестьсот) рублей 00 копеек", True),
        (" во исполнение п. 2.3 Дополнительного соглашения от 12.11.2026 г. к Договору купли-продажи комнаты от 12.11.2026 г., "
         "имеющей следующие технические характеристики: комната, этаж № 4, назначение: жилое, общая площадь 12,7 кв.м., "
         "местоположение: Пензенская область, г. Пенза, ул. Медицинская, д. 6, кадастровый номер 58:29:2014005:5301.", False)
    ], size_pt=10.5, space_after=3.0)

    add_p(doc_ds,
          "Денежные средства в размере 178 600 рублей возвращены Продавцом Покупателю в полном объеме. "
          "Обязательство Продавца по возврату излишне полученных денежных средств исполнено полностью. Финансовых претензий по возврату указанной суммы к Продавцу Коняеву П.Н. не имею.",
          size_pt=10.5, space_after=2.5)

    add_p(doc_ds,
          "Настоящая возвратная расписка составлена в двух, юридически равных экземплярах, по одному для каждой из сторон.",
          size_pt=10.5, space_after=4.0)

    add_sig_block(doc_ds, "Денежные средства передал (Продавец):", keep_all=True)
    add_sig_block(doc_ds, "Денежные средства получил (Покупатель):", keep_all=False)

    f1_desktop = os.path.join(desktop_dir, "Дополнительное_соглашение_12_11_2026_Коняев_Ситников_с_возвратной_распиской.docx")
    doc_ds.save(f1_desktop)
    print(f"СОХРАНЕНО НА РАБОЧИЙ СТОЛ: {f1_desktop}")

    # =========================================================================
    # ФАЙЛ 2: ОТДЕЛЬНАЯ РАСПИСКА ПО ДОГОВОРУ НА ПЕРВОНАЧАЛЬНЫЙ ВЗНОС 281 400 РУБ.
    # =========================================================================
    doc_r = docx.Document()
    for s in doc_r.sections:
        s.top_margin = Inches(0.65)
        s.bottom_margin = Inches(0.65)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.55)
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)

    add_h(doc_r, "РАСПИСКА", space_before=2.0, space_after=2.0, size_pt=12.0)
    add_h(doc_r, "в получении денежных средств (первоначальный взнос)", space_before=0, space_after=6.0, size_pt=10.5)
    add_city_date(doc_r, "г. Пенза", "12 ноября 2026 г.", size_pt=10.5)

    add_p_parts(doc_r, [
        ("Я, ", False),
        ("гр. РФ Коняев Петр Николаевич", True),
        (", 10.04.1960 года рождения, место рождения: село Мичкасс, Никольского р-на, Пензенской обл., "
         "паспорт гражданина Российской Федерации 5605 № 497454, выдан ОТДЕЛОМ ВНУТРЕННИХ ДЕЛ ЖЕЛЕЗНОДОРОЖНОГО РАЙОНА Г. ПЕНЗЫ, "
         "дата выдачи 29.06.2005, код подразделения 582-004, зарегистрированный по адресу: г. Пенза, ул. Кл. Цеткин, д. 23, кв. 125, "
         "настоящей распиской подтверждаю, что лично получил от", False)
    ], size_pt=11.0, space_after=2.5)

    add_p_parts(doc_r, [
        ("гр. РФ Ситникова Алексея Сергеевича", True),
        (", 26.04.1995 года рождения, место рождения: с. Дубровское, Лопатинского р-на, Пензенской обл., "
         "паспорт гражданина Российской Федерации 5614 № 264654, выдан МЕЖРАЙОННЫМ ОТДЕЛЕНИЕМ УФМС РОССИИ ПО ПЕНЗЕНСКОЙ ОБЛАСТИ "
         "В ШЕМЫШЕЙСКОМ РАЙОНЕ, дата выдачи 27.05.2015, код подразделения 580-033, зарегистрированного по адресу: Пензенская обл., "
         "Лопатинский р-н, ул. Южная, д. 6, кв. 5,", False)
    ], size_pt=11.0, space_after=2.5)

    add_p_parts(doc_r, [
        ("денежные средства в размере ", False),
        ("281 400 (Двести восемьдесят одна тысяча четыреста) рублей 00 копеек", True),
        (" в качестве первоначального взноса (оплаты части стоимости за счет собственных денежных средств Покупателя) "
         "за продаваемую комнату по Договору купли-продажи комнаты от 12 ноября 2026 года, имеющую следующие технические характеристики:", False)
    ], size_pt=11.0, space_after=2.0)

    add_p(doc_r,
          "Комната, этаж № 4, назначение: жилое, общая площадь 12,7 кв.м., местоположение: Пензенская область, г. Пенза, ул. Медицинская, д. 6, "
          "кадастровый номер 58:29:2014005:5301.",
          size_pt=11.0, bold=True, space_after=3.5)

    add_p(doc_r,
          "Денежные средства в размере 281 400 (Двести восемьдесят одна тысяча четыреста) рублей переданы Покупателем и получены Продавцом лично в полном объеме в момент подписания Договора купли-продажи. "
          "Претензий по сумме и порядку передачи указанного первоначального взноса к Ситникову Алексею Сергеевичу не имею.",
          size_pt=11.0, space_after=3.0)

    add_p(doc_r,
          "Настоящая расписка составлена в двух, юридически равных экземплярах, по одному для каждой из сторон.",
          size_pt=11.0, space_after=6.0)

    add_sig_block(doc_r, "Деньги получил (Продавец):", keep_all=True)
    add_sig_block(doc_r, "Деньги передал (Покупатель):", keep_all=False)

    f2_desktop = os.path.join(desktop_dir, "Расписка_281400_Первоначальный_взнос_Коняев_Ситников.docx")
    doc_r.save(f2_desktop)
    print(f"СОХРАНЕНО НА РАБОЧИЙ СТОЛ: {f2_desktop}")

    # Удаляем старый файл с неправильной суммой на рабочем столе
    old_f2 = os.path.join(desktop_dir, "Расписка_287430_Первоначальный_взнос_Коняев_Ситников.docx")
    if os.path.exists(old_f2):
        os.remove(old_f2)
        print(f"УДАЛЕН СТАРЫЙ ФАЙЛ: {old_f2}")

if __name__ == "__main__":
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    generate_all(desktop)
