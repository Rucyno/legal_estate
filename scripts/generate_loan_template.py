# -*- coding: utf-8 -*-
"""
Генератор шаблона: Договор займа (15% годовых, без пеней) + Соглашение о расторжении
Единый Word-документ со строгим соблюдением правил юридического оформления и фирменных стандартов.
"""
import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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

def add_p(doc, text="", size_pt=10.0, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
          space_after=2.0, line_spacing=1.08, indent=21.25, keep_next=False, page_break_before=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent and align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        p.paragraph_format.first_line_indent = Pt(indent)
    if keep_next:
        p.paragraph_format.keep_with_next = True
    if page_break_before:
        p.paragraph_format.page_break_before = True
    if text:
        run = p.add_run(text)
        set_run_ru(run, size_pt=size_pt, bold=bold, italic=italic)
    return p

def add_header_table(doc, city_text="г. ___________________", date_text="«____» _________________ 20___ г."):
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
    p_l.paragraph_format.space_after = Pt(4.0)
    r_l = p_l.add_run(city_text)
    set_run_ru(r_l, size_pt=10.0, bold=True)

    p_r = tbl.cell(0, 1).paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_after = Pt(4.0)
    r_r = p_r.add_run(date_text)
    set_run_ru(r_r, size_pt=10.0, bold=True)

def add_section_header(doc, title_text, page_break_before=False):
    return add_p(doc, title_text, size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                 indent=0, space_after=3.0, keep_next=True, page_break_before=page_break_before)

def add_clean_requisites_table(doc):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    col_width = Pt(265)
    for cell in tbl.rows[0].cells:
        cell.width = col_width
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
    
    # Left cell: Займодавец
    c_left = tbl.cell(0, 0)
    p_lh = c_left.paragraphs[0]
    p_lh.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_lh.paragraph_format.space_after = Pt(2.0)
    r_lh = p_lh.add_run("Займодавец:")
    set_run_ru(r_lh, size_pt=10.0, bold=True)
    
    left_lines = [
        "Наименование: ____________________________________",
        "ОГРН / ОГРНИП: ___________________________________",
        "ИНН: ____________________ КПП: ____________________",
        "Юридический адрес: _______________________________",
        "_________________________________________________",
        "Банк: ____________________________________________",
        "Р/с: _____________________________________________",
        "К/с: _____________________________________________",
        "БИК: ____________________________________________",
        "Тел. / E-mail: ____________________________________"
    ]
    for line in left_lines:
        p = c_left.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(line)
        set_run_ru(r, size_pt=9.0)
        
    # Right cell: Заемщик
    c_right = tbl.cell(0, 1)
    p_rh = c_right.paragraphs[0]
    p_rh.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_rh.paragraph_format.space_after = Pt(2.0)
    r_rh = p_rh.add_run("Заемщик:")
    set_run_ru(r_rh, size_pt=10.0, bold=True)
    
    right_lines = [
        "Наименование: ____________________________________",
        "ОГРН / ОГРНИП: ___________________________________",
        "ИНН: ____________________ КПП: ____________________",
        "Юридический адрес: _______________________________",
        "_________________________________________________",
        "Банк: ____________________________________________",
        "Р/с: _____________________________________________",
        "К/с: _____________________________________________",
        "БИК: ____________________________________________",
        "Тел. / E-mail: ____________________________________"
    ]
    for line in right_lines:
        p = c_right.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(line)
        set_run_ru(r, size_pt=9.0)

def add_signature_block_signers(doc, is_agreement=False):
    # Полноширинный строгий блок подписей по Правилу 10
    add_p(doc, "Подписи Сторон:" if is_agreement else "8. ПОДПИСИ СТОРОН", size_pt=10.5, bold=True, indent=0, space_after=3.0, keep_next=True)
    
    # Займодавец
    add_p(doc, "От имени Займодавца:", size_pt=10.0, bold=True, indent=0, space_after=2.0, keep_next=True)
    
    p_pos1 = doc.add_paragraph()
    p_pos1.paragraph_format.space_after = Pt(2.0)
    p_pos1.paragraph_format.keep_with_next = True
    r_pos1 = p_pos1.add_run("Должность: ____________________________________________________________________")
    set_run_ru(r_pos1, size_pt=9.5)

    # Строка подписи
    p_line1 = doc.add_paragraph()
    pPr1 = p_line1._p.get_or_add_pPr()
    pBdr1 = parse_xml(f"""<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>""")
    pPr1.append(pBdr1)
    p_line1.paragraph_format.space_after = Pt(1.5)
    p_line1.paragraph_format.keep_with_next = True
    r_empty1 = p_line1.add_run(" ")
    set_run_ru(r_empty1, size_pt=9.5)
    
    p_sub1 = doc.add_paragraph()
    pPr_sub1 = p_sub1._p.get_or_add_pPr()
    tabs1 = parse_xml(f"""<w:tabs {nsdecls("w")}><w:tab w:val="right" w:pos="9640"/></w:tabs>""")
    pPr_sub1.append(tabs1)
    p_sub1.paragraph_format.space_after = Pt(6.0)
    p_sub1.paragraph_format.keep_with_next = True
    r1_1 = p_sub1.add_run("М.П.                     (подпись)                                                 (расшифровка подписи / ФИО)")
    set_run_ru(r1_1, size_pt=8.0, italic=True)
    
    # Заемщик
    add_p(doc, "От имени Заемщика:", size_pt=10.0, bold=True, indent=0, space_after=2.0, keep_next=True)
    
    p_pos2 = doc.add_paragraph()
    p_pos2.paragraph_format.space_after = Pt(2.0)
    p_pos2.paragraph_format.keep_with_next = True
    r_pos2 = p_pos2.add_run("Должность: ____________________________________________________________________")
    set_run_ru(r_pos2, size_pt=9.5)

    p_line2 = doc.add_paragraph()
    pPr2 = p_line2._p.get_or_add_pPr()
    pBdr2 = parse_xml(f"""<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>""")
    pPr2.append(pBdr2)
    p_line2.paragraph_format.space_after = Pt(1.5)
    p_line2.paragraph_format.keep_with_next = True
    r_empty2 = p_line2.add_run(" ")
    set_run_ru(r_empty2, size_pt=9.5)
    
    p_sub2 = doc.add_paragraph()
    pPr_sub2 = p_sub2._p.get_or_add_pPr()
    tabs2 = parse_xml(f"""<w:tabs {nsdecls("w")}><w:tab w:val="right" w:pos="9640"/></w:tabs>""")
    pPr_sub2.append(tabs2)
    p_sub2.paragraph_format.space_after = Pt(2.0)
    p_sub2.paragraph_format.keep_with_next = False
    r2_1 = p_sub2.add_run("М.П.                     (подпись)                                                 (расшифровка подписи / ФИО)")
    set_run_ru(r2_1, size_pt=8.0, italic=True)

def build_full_loan_package(out_path):
    doc = docx.Document()
    
    for s in doc.sections:
        s.top_margin = Pt(36.0)
        s.bottom_margin = Pt(36.0)
        s.left_margin = Pt(45.0)
        s.right_margin = Pt(45.0)

    # ==========================================
    # ЧАСТЬ 1: ДОГОВОР ЗАЙМА (ШАБЛОН) - СТР. 1
    # ==========================================
    add_p(doc, "ДОГОВОР ЗАЙМА № _____", size_pt=12.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=3.0)
    add_header_table(doc, "г. ___________________", "«____» _________________ 20___ г.")

    preamble = (
        "____________________________________________________________________________________, "
        "ОГРН/ОГРНИП: __________________________, ИНН: ________________________, именуемое(-ый) в дальнейшем "
        "«Займодавец», в лице __________________________________________________________________, "
        "действующего(-ей) на основании ____________________________________________, с одной стороны, и\n"
        "____________________________________________________________________________________, "
        "ОГРН/ОГРНИП: __________________________, ИНН: ________________________, именуемое(-ый) в дальнейшем "
        "«Заемщик», в лице ____________________________________________________________________, "
        "действующего(-ей) на основании ____________________________________________, с другой стороны, "
        "при совместном упоминании именуемые «Стороны», а по отдельности — «Сторона», заключили настоящий Договор о нижеследующем:"
    )
    add_p(doc, preamble, size_pt=10.0, indent=21.25, space_after=3.0)

    # 1. ПРЕДМЕТ ДОГОВОРА
    add_section_header(doc, "1. ПРЕДМЕТ ДОГОВОРА")
    add_p(doc, "1.1. По настоящему Договору Займодавец передает Заемщику в собственность денежные средства в размере ________________________ (________________________________________________) рублей (далее — «Сумма займа»), а Заемщик обязуется возвратить Займодавцу полученную Сумму займа и уплатить проценты за пользование ею в порядке и на условиях, установленных настоящим Договором.", size_pt=10.0, indent=21.25, space_after=2.0)
    add_p(doc, "1.2. Заем считается предоставленным с момента поступления (зачисления) денежных средств на расчетный счет Заемщика либо выдачи из кассы Займодавца в установленном законом порядке.", size_pt=10.0, indent=21.25, space_after=2.0)
    add_p(doc, "1.3. За пользование Суммой займа Заемщик уплачивает Займодавцу проценты из расчета 15 (Пятнадцать) процентов годовых. Проценты начисляются со дня, следующего за днем фактического предоставления Суммы займа Заемщику, до дня фактического возврата Суммы займа Займодавцу включительно. Начисленные проценты уплачиваются Заемщиком: [ежемесячно не позднее ____ числа каждого месяца / единовременно в дату полного возврата Суммы займа].", size_pt=10.0, indent=21.25, space_after=3.0)

    # 2. ПРАВА И ОБЯЗАННОСТИ СТОРОН
    add_section_header(doc, "2. ПРАВА И ОБЯЗАННОСТИ СТОРОН")
    add_p(doc, "2.1. Займодавец обязан передать (перечислить) Заемщику указанную в п. 1.1 настоящего Договора Сумму займа в срок не позднее «____» _________________ 20___ г.", size_pt=10.0, indent=21.25, space_after=2.0)
    add_p(doc, "2.2. Возврат полученной Суммы займа и уплата начисленных процентов осуществляются Заемщиком в срок до «____» _________________ 20___ г. Заемщик имеет право возвратить Сумму займа и причитающиеся проценты досрочно без применения штрафных санкций.", size_pt=10.0, indent=21.25, space_after=3.0)

    # 3. ОТВЕТСТВЕННОСТЬ СТОРОН
    add_section_header(doc, "3. ОТВЕТСТВЕННОСТЬ СТОРОН")
    add_p(doc, "3.1. В случае неисполнения или ненадлежащего исполнения обязательств по настоящему Договору Стороны несут ответственность в соответствии с действующим законодательством Российской Федерации и обязаны возместить другой Стороне причиненные документально подтвержденные убытки.", size_pt=10.0, indent=21.25, space_after=2.0)
    add_p(doc, "3.2. Стороны пришли к взаимному соглашению не устанавливать договором договорную неустойку (пени и штрафы) за нарушение сроков исполнения обязательств. Ответственность Сторон регулируется общими нормами Гражданского кодекса Российской Федерации.", size_pt=10.0, indent=21.25, space_after=3.0)

    # 4. РАЗРЕШЕНИЕ СПОРОВ
    add_section_header(doc, "4. РАЗРЕШЕНИЕ СПОРОВ")
    add_p(doc, "4.1. Все споры, разногласия или требования, возникающие из настоящего Договора или в связи с ним, разрешаются Сторонами путем переговоров с обязательным соблюдением досудебного претензионного порядка. Срок рассмотрения письменной претензии и направления мотивированного ответа составляет 10 (Десять) рабочих дней со дня ее получения Стороной-адресатом.", size_pt=10.0, indent=21.25, space_after=2.0)
    add_p(doc, "4.2. При недостижении согласия в досудебном порядке спор передается на рассмотрение в суд по месту нахождения Займодавца в соответствии с процессуальным законодательством Российской Федерации.", size_pt=10.0, indent=21.25, space_after=4.0)

    # ==========================================
    # СТР. 2: СРОК ДЕЙСТВИЯ, ЗАКЛЮЧИТЕЛЬНЫЕ, РЕКВИЗИТЫ И ПОДПИСИ
    # ==========================================
    add_section_header(doc, "5. СРОК ДЕЙСТВИЯ И ПРЕКРАЩЕНИЕ ДОГОВОРА", page_break_before=True)
    add_p(doc, "5.1. Настоящий Договор вступает в силу с момента передачи Займодавцем Заемщику Суммы займа в полном объеме и прекращает свое действие после полного и надлежащего исполнения Сторонами всех принятых на себя обязательств.", size_pt=10.0, indent=21.25, space_after=3.0)

    # 6. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ
    add_section_header(doc, "6. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ")
    add_p(doc, "6.1. Любые изменения, дополнения и соглашения к настоящему Договору действительны при условии, если они совершены в письменной форме и подписаны уполномоченными представителями обеих Сторон.", size_pt=10.0, indent=21.25, space_after=2.0)
    add_p(doc, "6.2. Все уведомления, требования и юридически значимые сообщения должны направляться Сторонами в письменной форме заказным письмом с уведомлением о вручении либо с курьерской доставкой, а также могут дублироваться по адресам электронной почты, указанным в разделе 7 настоящего Договора.", size_pt=10.0, indent=21.25, space_after=2.0)
    add_p(doc, "6.3. Настоящий Договор составлен в 2 (двух) подлинных экземплярах, имеющих одинаковую юридическую силу, по одному экземпляру для каждой из Сторон.", size_pt=10.0, indent=21.25, space_after=4.0, keep_next=True)

    # 7. АДРЕСА И РЕКВИЗИТЫ СТОРОН
    add_section_header(doc, "7. АДРЕСА И РЕКВИЗИТЫ СТОРОН")
    add_clean_requisites_table(doc)
    add_p(doc, "", space_after=4.0, keep_next=True)

    # 8. ПОДПИСИ СТОРОН
    add_signature_block_signers(doc, is_agreement=False)

    # ==========================================
    # РАЗРЫВ СТРАНИЦЫ
    # ==========================================
    doc.add_page_break()

    # ==========================================
    # ЧАСТЬ 2: СОГЛАШЕНИЕ О РАСТОРЖЕНИИ (ШАБЛОН) - СТР. 3
    # ==========================================
    add_p(doc, "СОГЛАШЕНИЕ О РАСТОРЖЕНИИ", size_pt=12.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=1.0)
    add_p(doc, "Договора займа от «____» _________________ 20___ г. № _____", size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=3.0)
    add_header_table(doc, "г. ___________________", "«____» _________________ 20___ г.")

    preamble_term = (
        "____________________________________________________________________________________, "
        "ОГРН/ОГРНИП: __________________________, ИНН: ________________________, именуемое(-ый) в дальнейшем "
        "«Займодавец», в лице __________________________________________________________________, "
        "действующего(-ей) на основании ____________________________________________, с одной стороны, и\n"
        "____________________________________________________________________________________, "
        "ОГРН/ОГРНИП: __________________________, ИНН: ________________________, именуемое(-ый) в дальнейшем "
        "«Заемщик», в лице ____________________________________________________________________, "
        "действующего(-ей) на основании ____________________________________________, с другой стороны, "
        "руководствуясь пунктом 1 статьи 450, статьями 451, 453 Гражданского кодекса Российской Федерации, "
        "заключили настоящее Соглашение о расторжении Договора займа (далее — «Соглашение») о нижеследующем:"
    )
    add_p(doc, preamble_term, size_pt=10.0, indent=21.25, space_after=3.0)

    add_p(doc, "1. По взаимному соглашению Сторон расторгнуть Договор займа от «____» _________________ 20___ г. № _____ (далее — «Договор») с момента подписания настоящего Соглашения.", size_pt=10.0, indent=21.25, space_after=2.5)
    add_p(doc, "2. Взаимные обязательства Сторон, вытекающие из Договора, считаются полностью прекращенными с момента заключения (подписания) настоящего Соглашения, за исключением обязательств, прямо предусмотренных настоящим Соглашением.", size_pt=10.0, indent=21.25, space_after=2.5)
    add_p(doc, "3. Стороны подтверждают, что обязательства по Договору исполнены в полном объеме и надлежащим образом: основная сумма займа в размере ________________________ (________________________________________________) рублей и причитающиеся проценты за пользование займом в размере ________________________ (________________________________________________) рублей возвращены Заемщиком и получены Займодавцем в полном объеме. Каких-либо финансовых, материальных или иных претензий и требований по Договору Стороны друг к другу не имеют.", size_pt=10.0, indent=21.25, space_after=2.5)
    add_p(doc, "4. Настоящее Соглашение составлено в 2 (двух) подлинных экземплярах, имеющих одинаковую юридическую силу, по одному экземпляру для каждой из Сторон, вступает в силу с момента его подписания уполномоченными представителями Сторон и скрепления печатями (при наличии печатей).", size_pt=10.0, indent=21.25, space_after=4.0, keep_next=True)

    add_section_header(doc, "5. АДРЕСА, РЕКВИЗИТЫ И ПОДПИСИ СТОРОН")
    add_clean_requisites_table(doc)
    add_p(doc, "", space_after=4.0, keep_next=True)
    add_signature_block_signers(doc, is_agreement=True)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print("Full Loan Template Package saved successfully:", out_path)

if __name__ == "__main__":
    out_file = r"D:\Антигравити\legal_estate\templates\ШАБЛОН_Договор_займа_с_расторжением.docx"
    build_full_loan_package(out_file)
