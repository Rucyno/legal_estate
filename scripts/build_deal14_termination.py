# -*- coding: utf-8 -*-
"""
Генератор Соглашения о расторжении предварительного договора купли-продажи квартиры
(Тюлюкина Л.Д. и Лазутина Е.В., с. Засечное, ул. Центральная, д. 3, кв. 30).
Формат: строго чистый DOCX, ровно 1 страница А4, Times New Roman 10 pt.
Полное 100% совпадение всех паспортных данных и реквизитов объекта с Предварительным договором.
Юридически железобетонное обоснование удержания задатка по ст. 380, абз. 1 п. 2 ст. 381 ГК РФ и п. 5 ПДКП.
Подписи: Правило 10 (построчные во всю ширину листа с нижней границей параграфа).
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
          space_after=2.5, line_spacing=1.08, indent=21.25, keep_next=False):
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

def add_header_table(doc, city_text="г. Пенза", date_text="«28» октября 2026 года"):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for cell in tbl.rows[0].cells:
        cell.width = Pt(255)
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
    p_l.paragraph_format.space_after = Pt(4.0)
    r_l = p_l.add_run(city_text)
    set_run_ru(r_l, size_pt=10.0, bold=True)

    p_r = tbl.cell(0, 1).paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_after = Pt(4.0)
    r_r = p_r.add_run(date_text)
    set_run_ru(r_r, size_pt=10.0, bold=True)

def add_sig_block(doc, title_text, keep_all=True):
    # 1. Заголовок стороны
    add_p(doc, title_text, size_pt=10.0, bold=True, indent=0, space_after=1.5, keep_next=True)
    
    # 2. Полноширинная сплошная черта через нижнюю границу параграфа
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
    
    # 3. Подстрочник с табом
    p_sub = doc.add_paragraph()
    pPr_sub = p_sub._p.get_or_add_pPr()
    tabs = parse_xml(f"""
    <w:tabs {nsdecls("w")}>
        <w:tab w:val="right" w:pos="9640"/>
    </w:tabs>
    """)
    pPr_sub.append(tabs)
    p_sub.paragraph_format.space_after = Pt(4.0 if keep_all else 1.0)
    p_sub.paragraph_format.keep_with_next = True
    
    r1 = p_sub.add_run("ФИО")
    set_run_ru(r1, size_pt=8.0, italic=True)
    r2 = p_sub.add_run("\tподпись")
    set_run_ru(r2, size_pt=8.0, italic=True)

def generate_clean_docx(out_path):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Pt(36.0)
        s.bottom_margin = Pt(36.0)
        s.left_margin = Pt(40.0)
        s.right_margin = Pt(40.0)

    # Заголовок
    add_p(doc, "СОГЛАШЕНИЕ О РАСТОРЖЕНИИ", size_pt=11.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=2.0)
    add_p(doc, "Предварительного договора купли-продажи квартиры от 22 октября 2026 года", size_pt=10.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=4.0)

    add_header_table(doc, "г. Пенза", "«28» октября 2026 года")

    # Преамбула - СТРОГО 1-в-1 по тексту Предварительного договора
    seller_text = (
        "Мы, Тюлюкина Любовь Даниловна, дата рождения: 19.02.1954г., паспорт гражданина РФ 56 02 987708, "
        "выдан 06.02.2003г. Отделом внутренних дел Первомайского района г. Пензы, адрес регистрации: "
        "Пензенская область, г. Пенза, ул. Бригадная, д. 22, именуемая в дальнейшем «Продавец», с одной стороны,"
    )
    add_p(doc, seller_text, size_pt=10.0, indent=21.25, space_after=2.0)

    buyer_text = (
        "и Лазутина Елена Викторовна, дата рождения: 10.09.1984г., паспорт гражданина РФ 56 05 562191, "
        "выдан 05.05.2006г. Отделом внутренних дел Первомайского района г. Пензы, адрес регистрации: "
        "Пензенская область, Пензенский район, с. Засечное, ул. Олимпийская, д.3, кв. 129, "
        "именуемая в дальнейшем «Покупатель», с другой стороны,"
    )
    add_p(doc, buyer_text, size_pt=10.0, indent=21.25, space_after=2.0)

    add_p(doc, "при совместном упоминании именуемые «Стороны», руководствуясь статьями 380, 381, 421, 429, 450, 452, 453 Гражданского кодекса Российской Федерации, заключили настоящее Соглашение о нижеследующем:", size_pt=10.0, indent=21.25, space_after=3.0)

    # Пункт 1
    p1 = (
        "1. По взаимному согласию Сторон и на основании статей 450, 452, 453 Гражданского кодекса РФ Стороны "
        "расторгают Предварительный договор купли-продажи квартиры от 22 октября 2026 года (далее — «Договор»), "
        "в соответствии с которым Стороны обязывались заключить в будущем договор купли-продажи недвижимого имущества: "
        "1-комнатной квартиры общей площадью 44,8 кв. м., расположенной на 1 этаже жилого дома по адресу: "
        "Российская Федерация, Пензенская область, м.р-н Пензенский, с.п. Засечный сельсовет, с. Засечное, "
        "ул. Центральная, д. 3, кв. 30, кадастровый номер: 58:24:0381402:206 (далее — «Квартира»). "
        "С момента подписания настоящего Соглашения действие Договора прекращается, все обязательства Сторон "
        "по заключению основного договора купли-продажи Квартиры, передаче имущества и расчетам прекращаются "
        "в полном объеме (ст. 407, 453 ГК РФ)."
    )
    add_p(doc, p1, size_pt=10.0, indent=21.25, space_after=2.5)

    # Пункт 2 - Юридически безупречное обоснование оставления задатка
    p2_part1 = (
        "2. Стороны подтверждают, что Продавец надлежащим образом исполнил все свои обязательства по Договору, "
        "своевременно подготовил Квартиру и комплект необходимых документов к сделке и выражал полную готовность "
        "заключить основной договор купли-продажи в установленный пунктом 2 Договора срок (до 01.11.2026г.) "
        "на согласованных Сторонами условиях. Покупатель заявляет о своем добровольном отказе от заключения "
        "основного договора купли-продажи Квартиры по собственным причинам, не связанным с нарушением Продавцом "
        "условий Договора либо качеством Квартиры. Незаключение основного договора купли-продажи признается Сторонами "
        "неисполнением обязательств по вине Покупателя."
    )
    add_p(doc, p2_part1, size_pt=10.0, indent=21.25, space_after=2.0)

    p2_part2 = (
        "В соответствии с абзацем первым пункта 2 статьи 381 Гражданского кодекса РФ и пунктом 5 Договора "
        "сумма задатка в размере 50 000 (пятьдесят тысяч) рублей 00 копеек, оплаченная Покупателем в день подписания Договора, "
        "в связи с неисполнением обязательств по вине Покупателя возврату не подлежит и в полном объеме остается у Продавца. "
        "Покупатель подтверждает, что положения статей 380, 381 ГК РФ о правовой природе задатка и последствиях неисполнения договора "
        "ему понятны, указанная сумма не является авансом либо неосновательным обогащением (ст. 1102 ГК РФ), "
        "Покупатель согласен с оставлением всей суммы задатка у Продавца и претензий по ее возврату не имеет."
    )
    add_p(doc, p2_part2, size_pt=10.0, indent=21.25, space_after=2.5)

    # Пункт 3 - Отсутствие претензий
    p3 = (
        "3. Стороны подтверждают, что с момента подписания настоящего Соглашения не имеют и впредь не будут иметь "
        "друг к другу каких-либо претензий материального, финансового или иного характера, связанных с заключением, "
        "исполнением или расторжением Договора, в том числе требований о возврате денежных средств, уплате неустоек, "
        "процентов за пользование денежными средствами (ст. 395 ГК РФ) либо возмещении любых убытков и расходов (ст. 15, 393 ГК РФ)."
    )
    add_p(doc, p3, size_pt=10.0, indent=21.25, space_after=2.5)

    # Пункт 4 - Дееспособность и добровольность
    p4 = (
        "4. Стороны заявляют, что действуют добровольно, находятся в здравом уме и ясной памяти, осознают значение "
        "своих действий и их последствия. Настоящее Соглашение не совершено под влиянием обмана, насилия, угрозы "
        "или стечения тяжелых обстоятельств (ст. 178, 179 ГК РФ)."
    )
    add_p(doc, p4, size_pt=10.0, indent=21.25, space_after=2.5)

    # Пункт 5 - Экземпляры
    p5 = (
        "5. Настоящее Соглашение составлено в двух имеющих одинаковую юридическую силу подлинных экземплярах, "
        "по одному для каждой из Сторон, и вступает в силу с момента его подписания."
    )
    add_p(doc, p5, size_pt=10.0, indent=21.25, space_after=4.0, keep_next=True)

    # Подписи Сторон
    add_p(doc, "Подписи Сторон:", size_pt=10.0, bold=True, indent=0, space_after=2.0, keep_next=True)
    add_sig_block(doc, "Продавец:", keep_all=True)
    add_sig_block(doc, "Покупатель:", keep_all=False)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print("DOCX successfully generated:", out_path)

if __name__ == "__main__":
    target = r"D:\Антигравити\legal_estate\deals\14_Тюлюкина_Лазутина_Центральная_3\готовые\2026_10_28_Соглашение_о_расторжении_ПДКП_Тюлюкина_Лазутина_Центральная_3.docx"
    generate_clean_docx(target)
