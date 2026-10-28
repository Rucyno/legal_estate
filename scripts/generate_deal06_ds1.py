# -*- coding: utf-8 -*-
"""
Генератор Дополнительного соглашения № 1 к Договору купли-продажи квартиры от «23» октября 2026 года
(Лунев А.А., Вострова В.В. -> Анохин К.В., г. Пенза, ул. Кордон Студеный, д. 48, кв. 266).
Строго 1 страница А4, Times New Roman 10 pt, Правило 10 по подписям, 0 мягких переносов строк.
"""
import os
import sys
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.stdout.reconfigure(encoding="utf-8")

def set_run_ru(run, font_name="Times New Roman", size_pt=10.0, bold=False, italic=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    if not rPr.xpath("./w:lang"):
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="ru-RU" w:bidi="ru-RU"/>'))

def add_p(doc, text="", size_pt=10.0, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
          space_after=1.5, line_spacing=1.04, indent=18.0, keep_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent and align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        p.paragraph_format.first_line_indent = Pt(indent)
    if keep_next:
        p.paragraph_format.keep_with_next = True
    if text:
        run = p.add_run(text)
        set_run_ru(run, size_pt=size_pt, bold=bold, italic=italic)
    return p

def add_header_table(doc, city_text="город Пенза", date_text="«28» октября 2026 года"):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for cell in tbl.rows[0].cells:
        cell.width = Pt(260)
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(r"""
            <w:tcBorders %s >
                <w:top w:val="none"/>
                <w:left w:val="none"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        """ % nsdecls("w"))
        tcPr.append(tcBorders)

    p_l = tbl.cell(0, 0).paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_l.paragraph_format.space_after = Pt(2.0)
    r_l = p_l.add_run(city_text)
    set_run_ru(r_l, size_pt=10.0, bold=True)

    p_r = tbl.cell(0, 1).paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_after = Pt(2.0)
    r_r = p_r.add_run(date_text)
    set_run_ru(r_r, size_pt=10.0, bold=True)

def add_sig_block(doc, title_text, keep_all=True):
    # 1. Заголовок
    add_p(doc, title_text, size_pt=9.5, bold=True, indent=0, space_after=0.5, keep_next=True)
    
    # 2. Полноширинная черта через нижнюю границу параграфа
    p_line = doc.add_paragraph()
    pPr = p_line._p.get_or_add_pPr()
    pBdr = parse_xml(f"""
    <w:pBdr {nsdecls("w")}>
        <w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>
    </w:pBdr>
    """)
    pPr.append(pBdr)
    p_line.paragraph_format.space_after = Pt(0.5)
    p_line.paragraph_format.keep_with_next = True
    r_empty = p_line.add_run(" ")
    set_run_ru(r_empty, size_pt=8.0)
    
    # 3. Подстрочник с табом
    p_sub = doc.add_paragraph()
    pPr_sub = p_sub._p.get_or_add_pPr()
    tabs = parse_xml(f"""
    <w:tabs {nsdecls("w")}>
        <w:tab w:val="right" w:pos="9640"/>
    </w:tabs>
    """)
    pPr_sub.append(tabs)
    p_sub.paragraph_format.space_after = Pt(2.5 if keep_all else 1.0)
    p_sub.paragraph_format.keep_with_next = True
    
    r1 = p_sub.add_run("ФИО")
    set_run_ru(r1, size_pt=7.5, italic=True)
    r2 = p_sub.add_run("\tподпись")
    set_run_ru(r2, size_pt=7.5, italic=True)

def generate_supplementary_agreement(out_path):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Pt(24.0)
        s.bottom_margin = Pt(24.0)
        s.left_margin = Pt(32.0)
        s.right_margin = Pt(32.0)

    # Заголовок
    add_p(doc, "ДОПОЛНИТЕЛЬНОЕ СОГЛАШЕНИЕ № 1", size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=1.0)
    add_p(doc, "к Договору купли-продажи квартиры от «23» октября 2026 года", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2.5)

    add_header_table(doc, "город Пенза", "«28» октября 2026 года")

    # Преамбула
    seller_text = (
        "Мы, нижеподписавшиеся: гр. Лунев Антон Александрович, 06.09.1991 года рождения, пол: мужской, "
        "место рождения: гор. Пенза, паспорт гражданина Российской Федерации: серия 5611 № 040130, выдан: ОУФМС России по Пензенской области "
        "в Октябрьском районе гор. Пензы 14.10.2011 г., код подразделения: 580-004, зарегистрированный по адресу: Российская Федерация, "
        "Пензенская область, г. Пенза, пр-кт Строителей, д. 160, кв. 318, и гр. Вострова Валерия Викторовна, 28.10.1989 года рождения, "
        "пол: женский, место рождения: гор. Никольск Пензенской области, паспорт гражданина Российской Федерации: серия 5611 № 038518, "
        "выдан: ОУФМС России по Пензенской области в Никольском районе 25.11.2011 г., код подразделения: 580-019, зарегистрированная "
        "по адресу: Российская Федерация, Пензенская область, Никольский район, г. Никольск, ул. Пугачевская, д. 51А, кв. 1, "
        "именуемые в дальнейшем совместно «Продавцы», с одной стороны, и"
    )
    add_p(doc, seller_text, size_pt=10.0, indent=18.0, space_after=1.5)

    buyer_text = (
        "гр. Анохин Константин Валерьевич, 14.07.1976 года рождения, пол: мужской, место рождения: гор. Пенза, "
        "паспорт гражданина Российской Федерации: серия 5621 № 629805, выдан: УМВД России по Пензенской области 16.07.2021 г., "
        "код подразделения: 580-001, зарегистрированный по адресу: Российская Федерация, Пензенская область, г. Пенза, "
        "ул. Ладожская, д. 105, кв. 220, именуемый в дальнейшем «Покупатель», с другой стороны, совместно именуемые «Стороны», "
        "являющиеся сторонами Договора купли-продажи квартиры от «23» октября 2026 года (далее — «Договор»), заключенного в отношении "
        "квартиры по адресу: Российская Федерация, Пензенская область, г. Пенза, ул. Кордон Студеный, д. 48, кв. 266 (кадастровый "
        "номер: 58:29:2005005:1345), заключили настоящее Дополнительное соглашение о нижеследующем:"
    )
    add_p(doc, buyer_text, size_pt=10.0, indent=18.0, space_after=2.0)

    # Пункт 1: Исправление адреса регистрации Лунева А.А.
    p1_1 = (
        "1. В преамбуле Договора исправить адрес регистрации Продавца Лунева Антона Александровича, указав верный номер дома: "
        "вместо номера дома «д. 150» читать «д. 160». В связи с этим абзац второй преамбулы Договора изложить в следующей редакции:"
    )
    p1_2 = (
        "«гр. Лунев Антон Александрович, 06.09.1991 года рождения, пол: мужской, место рождения: гор. Пенза, паспорт гражданина "
        "Российской Федерации: серия 5611 № 040130, выдан: ОУФМС России по Пензенской области в Октябрьском районе гор. Пензы 14.10.2011 г., "
        "код подразделения: 580-004, зарегистрированный по адресу: Российская Федерация, Пензенская область, г. Пенза, пр-кт Строителей, "
        "д. 160, кв. 318, и»."
    )
    add_p(doc, p1_1, size_pt=10.0, indent=18.0, space_after=1.0)
    add_p(doc, p1_2, size_pt=10.0, italic=True, indent=18.0, space_after=1.5)

    # Пункт 2: Исправление кредитного договора в п. 2.1
    p2_1 = (
        "2. В пункте 2.1 Договора исправить реквизиты Кредитного договора Продавцов в АО «Банк ДОМ.РФ», изложив пункт 2.1 Договора "
        "в следующей редакции:"
    )
    p2_2 = (
        "«2.1. Продавец ставит в известность Покупателя о том, что на момент заключения настоящего договора в отношении квартиры "
        "в силу ст. 77 Федерального закона «Об ипотеке (залоге недвижимости)» № 102-ФЗ от 16.07.1998 г. зарегистрировано обременение "
        "(ипотека в силу закона) в пользу Акционерного общества «Банк ДОМ.РФ», ИНН 7725038124, ОГРН 1037739527077, на основании "
        "Кредитного договора № 0133-00646/ИКР-21РБ от 23.12.2021 г. (номер записи регистрации ипотеки в силу закона: "
        "58:29:2005005:1345-58/067/2025-2 от 28.01.2025 г.).»."
    )
    add_p(doc, p2_1, size_pt=10.0, indent=18.0, space_after=1.0)
    add_p(doc, p2_2, size_pt=10.0, italic=True, indent=18.0, space_after=1.5)

    # Пункт 3: Исправление кредитного договора в п. 2.3
    p3_1 = (
        "3. В пункте 2.3 Договора исправить реквизиты Кредитного договора, изложив абзацы второй и третий пункта 2.3 Договора "
        "в следующей редакции:"
    )
    p3_2 = (
        "«- Первая часть стоимости Квартиры в размере 1 330 285 (Один миллион триста тридцать тысяч двести восемьдесят пять) рублей 87 копеек "
        "оплачивается Покупателем Продавцу за счёт собственных денежных средств для погашения ссудной задолженности по Кредитному договору "
        "Продавца № 0133-00646/ИКР-21РБ от 23.12.2021 г. и снятия обременения на вышеуказанную Квартиру в Росреестре."
    )
    p3_3 = (
        "Продавец обязуется в день подписания настоящего договора и получения вышеуказанной суммы подать заявление в банк на полное "
        "досрочное погашение задолженности по Кредитному договору № 0133-00646/ИКР-21РБ от 23.12.2021 г.»."
    )
    add_p(doc, p3_1, size_pt=10.0, indent=18.0, space_after=1.0)
    add_p(doc, p3_2, size_pt=10.0, italic=True, indent=18.0, space_after=1.0)
    add_p(doc, p3_3, size_pt=10.0, italic=True, indent=18.0, space_after=1.5)

    # Пункт 4 и 5
    p4 = (
        "4. Все остальные условия Договора от «23» октября 2026 года, не затронутые настоящим Дополнительным соглашением, "
        "остаются неизменными, и Стороны подтверждают свои обязательства по ним."
    )
    add_p(doc, p4, size_pt=10.0, indent=18.0, space_after=1.0)

    p5 = (
        "5. Настоящее Дополнительное соглашение вступает в силу с момента его подписания Сторонами, является неотъемлемой частью "
        "Договора и составлено в трех подлинных экземплярах, имеющих равную юридическую силу: по одному для каждой из Сторон и один "
        "экземпляр для органа регистрации прав (Росреестра)."
    )
    add_p(doc, p5, size_pt=10.0, indent=18.0, space_after=2.0, keep_next=True)

    # Подписи Сторон
    add_p(doc, "Подписи Сторон:", size_pt=10.0, bold=True, indent=0, space_after=1.0, keep_next=True)
    add_p(doc, "Продавцы:", size_pt=9.5, bold=True, indent=0, space_after=0.5, keep_next=True)
    add_sig_block(doc, "Лунев Антон Александрович:", keep_all=True)
    add_sig_block(doc, "Вострова Валерия Викторовна:", keep_all=True)
    add_p(doc, "Покупатель:", size_pt=9.5, bold=True, indent=0, space_after=0.5, keep_next=True)
    add_sig_block(doc, "Анохин Константин Валерьевич:", keep_all=False)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print("DOCX successfully generated:", out_path)

if __name__ == "__main__":
    out_docx = r"D:\Антигравити\legal_estate\deals\06_Салтыкова_Лунев_Вострова_Анохин\готовые\2026_10_28_ДС1_к_ДКП_Кордон_Студеный_48_Лунев_Вострова_Анохин.docx"
    generate_supplementary_agreement(out_docx)
