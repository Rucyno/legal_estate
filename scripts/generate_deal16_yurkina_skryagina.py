# -*- coding: utf-8 -*-
"""
Генератор полного пакета документов по сделке:
Продавец: Юркина Надежда Иосифовна
Покупатель: Скрягина Елена Ивановна
Квартира: г. Пенза, ул. Московская, д. 40, кв. 164
Дата: 12 ноября 2026 г.
Сумма: 7 000 000 руб. (50 000 руб. задаток наличными + 6 950 000 руб. наличными в день ОДКП)
Срок ОДКП: до 20 декабря 2026 г.
Передача ключей: в день подписания ОДКП.
Опись мебели и техники включена в общую цену.
Комплект:
1. ПДКП + Расписка на 50 000 руб. в одном файле.
2. СОПД Продавца (Юркина Н.И.) с логотипом «Этажи» на 1 стр. А4.
3. СОПД Покупателя (Скрягина Е.И.) с логотипом «Этажи» на 1 стр. А4.
4. Единый комплект СОПД для пакетной печати.

Все документы сохраняются на Рабочий стол и в deals/16_Юркина_Скрягина_Московская_40/.
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding="utf-8")

def set_run_ru(run, font_name="Times New Roman", size_pt=10.0, bold=False, italic=False, color_rgb=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    rPr = run._r.get_or_add_rPr()
    if not rPr.xpath("./w:lang"):
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="ru-RU" w:bidi="ru-RU"/>'))

def add_p(doc, text="", bold_prefix="", size_pt=10.0, bold=False, italic=False,
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

def add_h(doc, title_text, space_before=3.5, space_after=1.5, size_pt=10.5):
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

def add_city_date(doc, city_text="г. Пенза", date_text="«12» ноября 2026 г."):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    c1, c2 = tbl.cell(0, 0), tbl.cell(0, 1)
    c1.width = Pt(245)
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
    set_run_ru(r_l, font_name="Times New Roman", size_pt=10.0, bold=True)

    p_r = c2.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(2)
    p_r.paragraph_format.space_after = Pt(3)
    r_r = p_r.add_run(date_text)
    set_run_ru(r_r, font_name="Times New Roman", size_pt=10.0, bold=True)

def add_sig_block(doc, title_text, keep_all=True):
    add_p(doc, title_text, size_pt=10.0, bold=True, indent=0, space_after=1.5, keep_next=True)

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
    set_run_ru(r1, font_name="Times New Roman", size_pt=8.0, italic=False)
    r2 = p_sub.add_run("\tподпись")
    set_run_ru(r2, font_name="Times New Roman", size_pt=8.0, italic=False)

def build_pdkp(doc):
    add_h(doc, "ПРЕДВАРИТЕЛЬНЫЙ ДОГОВОР КУПЛИ-ПРОДАЖИ КВАРТИРЫ", space_before=0, space_after=0.5, size_pt=11.0)
    add_h(doc, "(с условием обеспечения обязательства задатком)", space_before=0, space_after=2.0, size_pt=10.0)
    add_city_date(doc, "г. Пенза", "«12» ноября 2026 г.")

    add_p(doc, "Мы, нижеподписавшиеся:", size_pt=10.0, space_after=1.5)

    add_p(doc,
          "гражданка РФ Юркина Надежда Иосифовна, пол: женский, 02.07.1959 года рождения, "
          "место рождения: с/з Серп и Молот Пензенского р-на Пензенской обл., паспорт гражданина РФ 5603 № 319157, "
          "выдан отделом внутренних дел Ленинского района г. Пензы 15.10.2004 г., код подразделения: 582-001, "
          "зарегистрированная по адресу: Пензенская область, г. Пенза, ул. Московская, д. 40, кв. 164, "
          "именуемая в дальнейшем «Продавец», с одной стороны, и",
          size_pt=10.0, space_after=1.5)

    add_p(doc,
          "гражданка РФ Скрягина Елена Ивановна, пол: женский, 04.03.1961 года рождения, "
          "место рождения: пос. Гусятник Белинского р-на Пензенской обл., паспорт гражданина РФ 5605 № 563953, "
          "выдан отделом внутренних дел г. Заречного Пензенской области 27.03.2006 г., код подразделения: 582-034, "
          "зарегистрированная по адресу: Пензенская область, г. Заречный, пр-кт Мира, д. 66, кв. 49, "
          "именуемая в дальнейшем «Покупатель», с другой стороны, "
          "при совместном упоминании именуемые «Стороны», находясь в здравом уме и твердой памяти, действуя добровольно, "
          "заключили настоящий Предварительный договор (далее – «Договор») о нижеследующем:",
          size_pt=10.0, space_after=2.0)

    # 1. ПРЕДМЕТ ДОГОВОРА
    add_h(doc, "1. Предмет предварительного договора")
    add_p(doc,
          "Стороны обязуются заключить в будущем договор купли-продажи жилого помещения (далее – «Основной договор») "
          "на согласованных в настоящем Договоре условиях.",
          bold_prefix="1.1. ")

    add_p(doc,
          "По Основному договору Продавец обязуется передать в собственность, а Покупатель обязуется принять "
          "в индивидуальную собственность и оплатить следующее недвижимое имущество: "
          "жилое помещение (квартира), назначение: жилое, расположенное по адресу: "
          "Российская Федерация, Пензенская область, г. Пенза, ул. Московская, д. 40, кв. 164 "
          "(кадастровый номер: ___________________________________, общая площадь: __________ кв.м) "
          "(далее по тексту – «Объект»).",
          bold_prefix="1.2. ")

    add_p(doc,
          "Отчуждаемый Объект принадлежит Продавцу на праве индивидуальной собственности на основании правоустанавливающих "
          "документов, о чем в Едином государственном реестре недвижимости (ЕГРН) сделана запись о государственной регистрации права "
          "№ ___________________________________ от «___» ____________ 20___ года.",
          bold_prefix="1.3. ")

    # 2. СРОКИ И ПОРЯДОК ЗАКЛЮЧЕНИЯ ОСНОВНОГО ДОГОВОРА И ПЕРЕДАЧА КЛЮЧЕЙ
    add_h(doc, "2. Сроки и порядок заключения Основного договора. Передача ключей")
    add_p(doc,
          "Стороны обязуются заключить Основной договор купли-продажи Объекта в срок до «20» декабря 2026 г. включительно. "
          "Основной договор может быть заключен в любой день до наступления указанного срока по взаимному согласованию Сторон.",
          bold_prefix="2.1. ")

    add_p(doc,
          "Основной договор совершается в простой письменной форме в соответствии с действующим законодательством РФ.",
          bold_prefix="2.2. ")

    add_p(doc,
          "Передача Объекта и ключей от него: фактическая передача Объекта, освобождение его от личных вещей Продавца "
          "и передача полного комплекта ключей Покупателю производятся непосредственно в день подписания Основного договора "
          "купли-продажи с составлением и подписанием Передаточного акта (акта приема-передачи квартиры).",
          bold_prefix="2.3. ")

    add_p(doc,
          "Расходы, связанные с государственной регистрацией перехода права собственности на Объект в органе регистрации прав "
          "(Росреестре), несет Покупатель. Расходы по сбору и подготовке предпродажного пакета документов несет Продавец.",
          bold_prefix="2.4. ")

    # 3. ЦЕНА ОБЪЕКТА И ПОРЯДОК РАСЧЕТОВ. ОБЕСПЕЧЕНИЕ ОБЯЗАТЕЛЬСТВ ЗАДАТКОМ
    add_h(doc, "3. Цена Объекта и порядок расчетов. Обеспечение обязательств задатком")
    add_p(doc,
          "Цена Объекта по соглашению Сторон составляет: 7 000 000 (Семь миллионов) рублей 00 копеек. "
          "Цена Объекта является окончательной, фиксированной и изменению в одностороннем порядке не подлежит.",
          bold_prefix="3.1. ")

    add_p(doc,
          "Расчет между Сторонами производится в наличной форме за счет собственных денежных средств Покупателя в следующем порядке:",
          bold_prefix="3.2. ", keep_next=True)

    add_p(doc,
          "В день подписания настоящего Предварительного договора Покупатель передает Продавцу в качестве ЗАДАТКА "
          "денежные средства в размере 50 000 (Пятьдесят тысяч) рублей 00 копеек за счет собственных наличных денежных средств. "
          "В подтверждение получения суммы задатка Продавец выдает Покупателю письменную Расписку, являющуюся Приложением № 1 "
          "к настоящему Договору. Сумма задатка засчитывается в счет оплаты общей стоимости Объекта.",
          bold_prefix="3.2.1. ")

    add_p(doc,
          "Оставшаяся часть стоимости Объекта в размере 6 950 000 (Шесть миллионов девятьсот пятьдесят тысяч) рублей 00 копеек "
          "оплачивается Покупателем за счет собственных наличных денежных средств непосредственно в день подписания Основного "
          "договора купли-продажи квартиры в срок до «20» декабря 2026 года.",
          bold_prefix="3.2.2. ")

    # 4. УСЛОВИЕ О ПЕРЕДАВАЕМОМ ИМУЩЕСТВЕ (МЕБЕЛЬ И ТЕХНИКА)
    add_h(doc, "4. Условие о передаваемом имуществе (мебель и бытовая техника)")
    add_p(doc,
          "Стороны договорились, что в продаваемом Объекте остается и передается в собственность Покупателя следующее имущество "
          "(мебель, предметы интерьера и бытовая техника):",
          bold_prefix="4.1. ", keep_next=True)

    furniture_items = [
        "- в прихожей: шкаф, обувница и банкетка;",
        "- в коридоре: большой зеркальный шкаф с туалетным столиком;",
        "- в зале: шкафы, полки, тумба под телевизор, диван (раскладывающийся в двуспальный, с местом для хранения), "
        "кровать (с местом для хранения), рабочий стол (с компьютерным стулом);",
        "- на балконе: стеллаж для хранения;",
        "- на кухне: кухонный гарнитур, встроенная бытовая техника, обеденный стол, стулья, табуретки, стеллаж."
    ]
    for item in furniture_items:
        add_p(doc, item, size_pt=9.5, indent=28.0)

    add_p(doc,
          "Стоимость указанного в п. 4.1 имущества полностью включена в общую цену Объекта (п. 3.1 Договора), "
          "дополнительной оплате не подлежит. Продавец гарантирует сохранение указанного имущества в исправном рабочем состоянии "
          "до момента фактической передачи Объекта Покупателю.",
          bold_prefix="4.2. ")

    # 5. ОБЕСПЕЧЕНИЕ ИСПОЛНЕНИЯ ОБЯЗАТЕЛЬСТВ И ОТВЕТСТВЕННОСТЬ СТОРОН
    add_h(doc, "5. Обеспечение исполнения обязательств и ответственность Сторон")
    add_p(doc,
          "Исполнение обязательств по настоящему Договору обеспечивается задатком в соответствии со ст. 380, 381 Гражданского кодекса РФ: "
          "в случае неисполнения обязательств по заключению Основного договора по вине Покупателя, сумма задатка в размере 50 000 рублей "
          "остается у Продавца; в случае неисполнения обязательств по заключению Основного договора по вине Продавца, Продавец обязан "
          "возвратить Покупателю сумму задатка в двойном размере — 100 000 (Сто тысяч) рублей в течение 3 (трех) рабочих дней со дня "
          "истечения срока, указанного в п. 2.1 Договора.",
          bold_prefix="5.1. ")

    add_p(doc,
          "В соответствии со ст. 429, 445 ГК РФ в случае уклонения одной из Сторон от заключения Основного договора добросовестная "
          "Сторона вправе обратиться в суд с требованием о понуждении заключить Основной договор и о возмещении причиненных этим убытков.",
          bold_prefix="5.2. ")

    # 6. ГАРАНТИИ И ЗАВЕРЕНИЯ ОБ ОБСТОЯТЕЛЬСТВАХ (СТ. 431.2 ГК РФ)
    add_h(doc, "6. Гарантии и заверения об обстоятельствах (ст. 431.2 ГК РФ)")
    add_p(doc,
          "Продавец гарантирует, что на момент подписания настоящего Договора Объект никому не продан, не подарен, не заложен, "
          "в споре и под арестом (запрещением) не состоит, правами третьих лиц не обременен.",
          bold_prefix="6.1. ")

    add_p(doc,
          "Продавец обязуется обеспечить снятие с регистрационного учета всех зарегистрированных в Объекте лиц (при их наличии) "
          "в срок не позднее 14 (четырнадцати) календарных дней с момента полной оплаты стоимости Объекта по Основному договору.",
          bold_prefix="6.2. ")

    add_p(doc,
          "Продавец гарантирует, что на приобретение Объекта не использовались средства материнского (семейного) капитала, "
          "обязательств по наделению несовершеннолетних долями не имеется. Продавец не имеет долгов и обязательств, которые могут повлечь "
          "его банкротство (127-ФЗ), не является иностранным агентом (255-ФЗ).",
          bold_prefix="6.3. ")

    add_p(doc,
          "Стороны подтверждают полную дееспособность, отсутствие опеки, попечительства и патронажа, добровольность заключения Договора.",
          bold_prefix="6.4. ")

    # 7. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ
    add_h(doc, "7. Заключительные положения")
    add_p(doc,
          "Настоящий Договор вступает в силу с момента его подписания Сторонами и действует до полного исполнения Сторонами своих обязательств.",
          bold_prefix="7.1. ")

    add_p(doc,
          "Настоящий Договор составлен и подписан в двух подлинных экземплярах, имеющих одинаковую юридическую силу: "
          "по одному экземпляру для каждой из Сторон.",
          bold_prefix="7.2. ")

    # 8. АДРЕСА, РЕКВИЗИТЫ И ПОДПИСИ СТОРОН
    add_h(doc, "8. Адреса, реквизиты и подписи Сторон", space_before=4.0, space_after=3.0)

    tbl_req = doc.add_table(rows=2, cols=2)
    tbl_req.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_req.autofit = False

    w_col = Pt(250)
    for row in tbl_req.rows:
        for cell in row.cells:
            cell.width = w_col
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for m, val in [('top', 60), ('bottom', 60), ('left', 70), ('right', 70)]:
                node = OxmlElement(f'w:{m}')
                node.set(qn('w:w'), str(val))
                node.set(qn('w:type'), 'dxa')
                tcMar.append(node)
            tcPr.append(tcMar)

    hdr_titles = ["ПРОДАВЕЦ:", "ПОКУПАТЕЛЬ:"]
    for i, title in enumerate(hdr_titles):
        cell = tbl_req.rows[0].cells[i]
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="C8102E"/>')
        tcPr.append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(title)
        set_run_ru(r, font_name="Times New Roman", size_pt=9.5, bold=True, color_rgb=RGBColor(255, 255, 255))

    b_cells = tbl_req.rows[1].cells
    for cell in b_cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(f"""
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
                <w:left w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
                <w:right w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/>
            </w:tcBorders>
        """)
        tcPr.append(tcBorders)

    seller_lines = [
        "Юркина Надежда Иосифовна",
        "Дата рождения: 02.07.1959 г.",
        "Паспорт гражданина РФ: 5603 № 319157",
        "Выдан: ОВД Ленинского района г. Пензы",
        "Дата выдачи: 15.10.2004 г., код подр.: 582-001",
        "Адрес регистрации: г. Пенза, ул. Московская, д. 40, кв. 164",
        "Телефон: ________________________________________"
    ]
    buyer_lines = [
        "Скрягина Елена Ивановна",
        "Дата рождения: 04.03.1961 г.",
        "Паспорт гражданина РФ: 5605 № 563953",
        "Выдан: ОВД г. Заречного Пензенской области",
        "Дата выдачи: 27.03.2006 г., код подр.: 582-034",
        "Адрес регистрации: Пензенская обл., г. Заречный, пр-кт Мира, д. 66, кв. 49",
        "Телефон: ________________________________________"
    ]
    for cell_idx, lines in enumerate([seller_lines, buyer_lines]):
        p_c0 = b_cells[cell_idx].paragraphs[0]
        for idx, line in enumerate(lines):
            p_item = b_cells[cell_idx].add_paragraph() if idx > 0 else p_c0
            p_item.paragraph_format.space_before = Pt(0)
            p_item.paragraph_format.space_after = Pt(1.5)
            p_item.paragraph_format.line_spacing = 1.05
            r_item = p_item.add_run(line)
            set_run_ru(r_item, font_name="Times New Roman", size_pt=9.0)

    add_p(doc, "", space_before=3.0, space_after=2.0)
    add_sig_block(doc, "Продавец: Юркина Надежда Иосифовна", keep_all=True)
    add_p(doc, "", space_before=2.0, space_after=2.0)
    add_sig_block(doc, "Покупатель: Скрягина Елена Ивановна", keep_all=False)

    # =========================================================================
    # ПРИЛОЖЕНИЕ № 1: РАСПИСКА В ПОЛУЧЕНИИ ЗАДАТКА 50 000 РУБ. (на отдельном листе)
    # =========================================================================
    doc.add_page_break()

    add_h(doc, "ПРИЛОЖЕНИЕ № 1", space_before=2.0, space_after=1.0, size_pt=10.0)
    add_h(doc, "РАСПИСКА В ПОЛУЧЕНИИ ЗАДАТКА", space_before=0, space_after=1.0, size_pt=11.5)
    add_h(doc, "к Предварительному договору купли-продажи квартиры от 12 ноября 2026 г.", space_before=0, space_after=4.0, size_pt=10.0)
    add_city_date(doc, "г. Пенза", "«12» ноября 2026 г.")

    add_p(doc,
          "Я, гражданка РФ Юркина Надежда Иосифовна, пол: женский, 02.07.1959 года рождения, "
          "место рождения: с/з Серп и Молот Пензенского р-на Пензенской обл., паспорт гражданина РФ 5603 № 319157, "
          "выдан отделом внутренних дел Ленинского района г. Пензы 15.10.2004 г., код подразделения: 582-001, "
          "зарегистрированная по адресу: Пензенская область, г. Пенза, ул. Московская, д. 40, кв. 164, "
          "настоящей распиской подтверждаю, что лично получила от",
          size_pt=10.5, space_after=2.0)

    add_p(doc,
          "гражданки РФ Скрягиной Елены Ивановны, пол: женский, 04.03.1961 года рождения, "
          "место рождения: пос. Гусятник Белинского р-на Пензенской обл., паспорт гражданина РФ 5605 № 563953, "
          "выдан отделом внутренних дел г. Заречного Пензенской области 27.03.2006 г., код подразделения: 582-034, "
          "зарегистрированной по адресу: Пензенская область, г. Заречный, пр-кт Мира, д. 66, кв. 49,",
          size_pt=10.5, space_after=2.0)

    add_p(doc,
          "наличные денежные средства в размере 50 000 (Пятьдесят тысяч) рублей 00 копеек в качестве обеспечения исполнения обязательств (ЗАДАТКА) "
          "по заключению Основного договора купли-продажи квартиры, расположенной по адресу: "
          "Российская Федерация, Пензенская область, г. Пенза, ул. Московская, д. 40, кв. 164, "
          "в соответствии с условиями Предварительного договора купли-продажи квартиры от «12» ноября 2026 года.",
          size_pt=10.5, bold=True, space_after=2.5)

    add_p(doc,
          "Денежные средства в сумме 50 000 (Пятьдесят тысяч) рублей переданы Покупателем Скрягиной Е.И. и приняты Продавцом Юркиной Н.И. "
          "лично наличными денежными средствами в полном объеме в момент подписания Предварительного договора. "
          "Материальных и финансовых претензий по передаче суммы задатка Продавец к Скрягиной Елене Ивановне не имеет.",
          size_pt=10.5, space_after=2.5)

    add_p(doc,
          "Оставшаяся часть стоимости квартиры в размере 6 950 000 (Шесть миллионов девятьсот пятьдесят тысяч) рублей 00 копеек "
          "подлежит уплате Покупателем наличными денежными средствами непосредственно в день подписания Основного договора купли-продажи квартиры "
          "в срок до «20» декабря 2026 года.",
          size_pt=10.5, space_after=3.5)

    add_p(doc,
          "Настоящая расписка составлена в 2 (двух) подлинных экземплярах, имеющих одинаковую юридическую силу: "
          "один экземпляр выдан Покупателю Скрягиной Е.И., второй хранится у Продавца Юркиной Н.И.",
          size_pt=10.5, space_after=5.0)

    add_sig_block(doc, "Деньги в качестве задатка приняла (Продавец): Юркина Надежда Иосифовна", keep_all=True)
    add_p(doc, "", space_before=2.0, space_after=2.0)
    add_sig_block(doc, "Деньги в качестве задатка передала (Покупатель): Скрягина Елена Ивановна", keep_all=False)

def build_sopd_page(doc, person_name, person_details, logo_path="templates/assets/etagi_logo.png"):
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(2)
        r_logo = p_logo.add_run()
        r_logo.add_picture(logo_path, width=Inches(1.8))

    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(2)
    p_t.paragraph_format.space_after = Pt(0.5)
    r1 = p_t.add_run("СОГЛАСИЕ")
    set_run_ru(r1, font_name="Times New Roman", size_pt=11.0, bold=True)

    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2.paragraph_format.space_before = Pt(0)
    p_t2.paragraph_format.space_after = Pt(2)
    r2 = p_t2.add_run("на обработку персональных данных")
    set_run_ru(r2, font_name="Times New Roman", size_pt=10.0, bold=True)

    preamble_text = (
        f"Я, гражданка РФ {person_name}, {person_details}, "
        "в соответствии со ст. 9 Федерального закона от 27.07.2006 № 152-ФЗ «О персональных данных», действуя свободно, "
        "своей волей и в своем интересе, а также подтверждая свою дееспособность, даю настоящее конкретное, предметное, "
        "информированное, сознательное и однозначное Согласие на обработку персональных данных (далее – Согласие) "
        "оператору персональных данных ООО «Межрегиональный Центр Недвижимости» (далее – Оператор), "
        "местом нахождения которого является: 440066, г. Пенза, пр-кт Строителей, д. 152, кв. 19, на следующих условиях:"
    )
    p_pre = doc.add_paragraph()
    p_pre.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_pre.paragraph_format.first_line_indent = Pt(14.0)
    p_pre.paragraph_format.space_before = Pt(2)
    p_pre.paragraph_format.space_after = Pt(2)
    p_pre.paragraph_format.line_spacing = 1.02
    r_pre = p_pre.add_run(preamble_text)
    set_run_ru(r_pre, font_name="Times New Roman", size_pt=8.5)

    clauses = [
        ("1. ", "Данное Согласие дается на обработку персональных данных, как без использования средств автоматизации, так и с их использованием."),
        ("2. ", "Согласие дается на обработку перечня следующих персональных данных, не являющихся специальными или биометрическими: фамилия, имя, отчество; дата и место рождения; гражданство; реквизиты документа, удостоверяющего личность; адрес места жительства (регистрации) и пребывания; контактный номер телефона; адрес электронной почты; ИНН; СНИЛС; сведения о зарегистрированных правах на объект недвижимого имущества и сведения о переходе прав."),
        ("3. ", "Цель обработки персональных данных: оказание услуг агентством недвижимости, содействие в подготовке, заключении и исполнении сделок с недвижимостью, заключение с Оператором любых договоров и соглашений на оказание таких услуг, а также проверка достоверности указанных сведений."),
        ("4. ", "Основанием для обработки персональных данных является: ст. 24 Конституции РФ; Федеральный закон № 152-ФЗ «О персональных данных»; Политика обработки персональных данных Оператора; настоящее Согласие."),
        ("5. ", "В ходе обработки с персональными данными будут совершены следующие действия: сбор, запись, систематизация, накопление, хранение, уточнение (обновление, изменение), извлечение, использование, передача (предоставление, доступ), обезличивание, блокирование, удаление, уничтожение."),
        ("6. ", "Субъект персональных данных согласен с тем, что в целях заключения, исполнения договоров купли-продажи/аренды и надлежащего оказания услуг Оператором его персональные данные могут быть переданы третьим лицам: банкам, органам регистрации прав (Росреестр), нотариусам, МФЦ."),
        ("7. ", "Субъект персональных данных подтверждает, что получено надлежащее согласие третьих лиц, чьи данные могут содержаться в документах, на их передачу и обработку Оператором."),
        ("8. ", "Персональные данные обрабатываются до момента достижения целей обработки либо отзыва согласия субъектом. Хранение осуществляется согласно Федеральному закону № 125-ФЗ «Об архивном деле в РФ»."),
        ("9. ", "Настоящее Согласие может быть отозвано субъектом персональных данных путем направления письменного заявления Оператору на электронную почту: mcn-realty@yandex.ru либо по адресу нахождения Оператора.")
    ]
    for num, txt in clauses:
        p_c = doc.add_paragraph()
        p_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_c.paragraph_format.first_line_indent = Pt(14.0)
        p_c.paragraph_format.space_before = Pt(0)
        p_c.paragraph_format.space_after = Pt(1.0)
        p_c.paragraph_format.line_spacing = 1.01
        r_n = p_c.add_run(num)
        set_run_ru(r_n, font_name="Times New Roman", size_pt=8.5, bold=True)
        r_t = p_c.add_run(txt)
        set_run_ru(r_t, font_name="Times New Roman", size_pt=8.5)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(4)
    p_sig.paragraph_format.space_after = Pt(1)
    p_sig.paragraph_format.keep_with_next = True
    r_s = p_sig.add_run("Субъект персональных данных:")
    set_run_ru(r_s, font_name="Times New Roman", size_pt=9.0, bold=True)

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
    set_run_ru(r_empty, size_pt=8.0)

    p_sub = doc.add_paragraph()
    pPr_sub = p_sub._p.get_or_add_pPr()
    tabs = parse_xml(f"""
    <w:tabs {nsdecls("w")}>
        <w:tab w:val="right" w:pos="9921"/>
    </w:tabs>
    """)
    pPr_sub.append(tabs)
    p_sub.paragraph_format.space_after = Pt(2.0)
    p_sub.paragraph_format.keep_with_next = True
    r_sub1 = p_sub.add_run(f"«12» ноября 2026 г. / {person_name}")
    set_run_ru(r_sub1, font_name="Times New Roman", size_pt=8.0)
    r_sub2 = p_sub.add_run("\tподпись")
    set_run_ru(r_sub2, font_name="Times New Roman", size_pt=8.0, italic=False)

def generate_deal_package():
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    deal_dir = r"d:\Антигравити\legal_estate\deals\16_Юркина_Скрягина_Московская_40"
    os.makedirs(deal_dir, exist_ok=True)

    # 1. ПДКП С РАСПИСКОЙ
    doc_pdkp = docx.Document()
    for s in doc_pdkp.sections:
        s.top_margin = Inches(0.55)
        s.bottom_margin = Inches(0.55)
        s.left_margin = Inches(0.65)
        s.right_margin = Inches(0.55)
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)

    build_pdkp(doc_pdkp)

    f_pdkp_desktop = os.path.join(desktop_dir, "ПДКП_с_распиской_Юркина_Скрягина_Московская_40.docx")
    f_pdkp_deal = os.path.join(deal_dir, "ПДКП_с_распиской_Юркина_Скрягина_Московская_40.docx")
    doc_pdkp.save(f_pdkp_desktop)
    doc_pdkp.save(f_pdkp_deal)
    print(f"ПДКП СОХРАНЕН: {f_pdkp_desktop}")

    # 2. СОПД ПРОДАВЦА (ЮРКИНА Н.И.)
    doc_sopd_seller = docx.Document()
    for s in doc_sopd_seller.sections:
        s.top_margin = Inches(0.45)
        s.bottom_margin = Inches(0.45)
        s.left_margin = Inches(0.60)
        s.right_margin = Inches(0.55)
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)

    details_seller = (
        "02.07.1959 года рождения, пол: женский, место рождения: с/з Серп и Молот Пензенского р-на Пензенской обл., "
        "паспорт гражданина РФ: 5603 № 319157, выдан отделом внутренних дел Ленинского района г. Пензы 15.10.2004 г., "
        "код подразделения: 582-001, адрес регистрации: г. Пенза, ул. Московская, д. 40, кв. 164"
    )
    build_sopd_page(doc_sopd_seller, "Юркина Надежда Иосифовна", details_seller)

    f_sopd_seller_desk = os.path.join(desktop_dir, "СОПД_Продавец_Юркина_Н_И.docx")
    f_sopd_seller_deal = os.path.join(deal_dir, "СОПД_Продавец_Юркина_Н_И.docx")
    doc_sopd_seller.save(f_sopd_seller_desk)
    doc_sopd_seller.save(f_sopd_seller_deal)
    print(f"СОПД ПРОДАВЦА СОХРАНЕНО: {f_sopd_seller_desk}")

    # 3. СОПД ПОКУПАТЕЛЯ (СКРЯГИНА Е.И.)
    doc_sopd_buyer = docx.Document()
    for s in doc_sopd_buyer.sections:
        s.top_margin = Inches(0.45)
        s.bottom_margin = Inches(0.45)
        s.left_margin = Inches(0.60)
        s.right_margin = Inches(0.55)
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)

    details_buyer = (
        "04.03.1961 года рождения, пол: женский, место рождения: пос. Гусятник Белинского р-на Пензенской обл., "
        "паспорт гражданина РФ: 5605 № 563953, выдан отделом внутренних дел г. Заречного Пензенской области 27.03.2006 г., "
        "код подразделения: 582-034, адрес регистрации: Пензенская обл., г. Заречный, пр-кт Мира, д. 66, кв. 49"
    )
    build_sopd_page(doc_sopd_buyer, "Скрягина Елена Ивановна", details_buyer)

    f_sopd_buyer_desk = os.path.join(desktop_dir, "СОПД_Покупатель_Скрягина_Е_И.docx")
    f_sopd_buyer_deal = os.path.join(deal_dir, "СОПД_Покупатель_Скрягина_Е_И.docx")
    doc_sopd_buyer.save(f_sopd_buyer_desk)
    doc_sopd_buyer.save(f_sopd_buyer_deal)
    print(f"СОПД ПОКУПАТЕЛЯ СОХРАНЕНО: {f_sopd_buyer_desk}")

    # 4. ЕДИНЫЙ КОМПЛЕКТ СОПД (ОБЕ СТОРОНЫ ДЛЯ ПАКЕТНОЙ ПЕЧАТИ)
    doc_sopd_bundle = docx.Document()
    for s in doc_sopd_bundle.sections:
        s.top_margin = Inches(0.45)
        s.bottom_margin = Inches(0.45)
        s.left_margin = Inches(0.60)
        s.right_margin = Inches(0.55)
        s.page_width = Inches(8.27)
        s.page_height = Inches(11.69)

    build_sopd_page(doc_sopd_bundle, "Юркина Надежда Иосифовна", details_seller)
    doc_sopd_bundle.add_page_break()
    build_sopd_page(doc_sopd_bundle, "Скрягина Елена Ивановна", details_buyer)

    f_sopd_bundle_desk = os.path.join(desktop_dir, "СОПД_Комплект_Юркина_Скрягина.docx")
    f_sopd_bundle_deal = os.path.join(deal_dir, "СОПД_Комплект_Юркина_Скрягина.docx")
    doc_sopd_bundle.save(f_sopd_bundle_desk)
    doc_sopd_bundle.save(f_sopd_bundle_deal)
    print(f"КОМПЛЕКТ СОПД СОХРАНЕН: {f_sopd_bundle_desk}")

if __name__ == "__main__":
    generate_deal_package()
