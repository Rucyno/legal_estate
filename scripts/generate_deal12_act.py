"""
Генератор Передаточного акта (акта приема-передачи квартиры) к ДКП.
Стандарт:
- Исключен п. 6 о взаиморасчетах (т.к. расчеты регулируются ДКП и аккредитивом/распиской, при физической передаче объекта этот пункт излишен).
- Пункт 5 (Показания приборов учета): показания каждого счетчика идут отдельной строкой с визуальным межстрочным интервалом/отступом для удобного рукописного заполнения цифр ручкой.
- Пункт 6: Экземпляры документа.
- Блок подписантов по Правилу 10 с комфортным отступом перед подписями.
- Верстка строго на 1 страницу А4.
- Шрифт: Times New Roman 9.5 pt (заголовки 10.5 pt, подстрочники 7.5 pt).
- 0 мягких переносов строк.
"""
import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import parse_xml

def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_run_ru(run, font_name="Times New Roman", size_pt=10.0, bold=False, italic=False, color_rgb=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.append(rFonts)

def add_p(doc, text="", size_pt=9.5, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, 
          space_before=0.0, space_after=1.5, line_spacing=1.03, indent=16.0, keep_with_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    if space_before > 0:
        p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent > 0:
        p.paragraph_format.first_line_indent = Pt(indent)
    if keep_with_next:
        p.paragraph_format.keep_with_next = True
    if text:
        r = p.add_run(text)
        set_run_ru(r, size_pt=size_pt, bold=bold, italic=italic)
    return p

def add_header_table(doc, left_text, right_text):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    col_widths = [Inches(3.35), Inches(3.35)]
    row = tbl.rows[0]
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
    
    cell_l = row.cells[0]
    cell_r = row.cells[1]
    
    cell_l.width = col_widths[0]
    cell_r.width = col_widths[1]
    
    set_cell_margins(cell_l, top=0, bottom=40, left=0, right=0)
    set_cell_margins(cell_r, top=0, bottom=40, left=0, right=0)
    
    p_l = cell_l.paragraphs[0]
    p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_l.paragraph_format.space_before = Pt(0)
    p_l.paragraph_format.space_after = Pt(2.0)
    p_l.paragraph_format.line_spacing = 1.0
    r_l = p_l.add_run(left_text)
    set_run_ru(r_l, size_pt=9.5, bold=False)
    
    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(0)
    p_r.paragraph_format.space_after = Pt(2.0)
    p_r.paragraph_format.line_spacing = 1.0
    r_r = p_r.add_run(right_text)
    set_run_ru(r_r, size_pt=9.5, bold=False)

def add_signer_block_rule10(doc, title, full_name, space_before=2.0, space_after=1.0):
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_before = Pt(space_before)
    p_title.paragraph_format.space_after = Pt(0.5)
    p_title.paragraph_format.line_spacing = 1.0
    p_title.paragraph_format.keep_with_next = True
    
    r_role = p_title.add_run(title + " ")
    set_run_ru(r_role, size_pt=9.5, bold=True)
    r_name = p_title.add_run(full_name)
    set_run_ru(r_name, size_pt=9.5, bold=False)
    
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_line.paragraph_format.space_before = Pt(1.5)
    p_line.paragraph_format.space_after = Pt(0.5)
    p_line.paragraph_format.line_spacing = 1.0
    p_line.paragraph_format.keep_with_next = True
    
    pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                     r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/>'
                     r'</w:pBdr>')
    p_line._p.get_or_add_pPr().append(pBdr)
    r_empty = p_line.add_run(" ")
    set_run_ru(r_empty, size_pt=7.0)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sub.paragraph_format.space_before = Pt(0.0)
    p_sub.paragraph_format.space_after = Pt(space_after)
    p_sub.paragraph_format.line_spacing = 1.0
    p_sub.paragraph_format.keep_with_next = True
    
    tabs = parse_xml(r'<w:tabs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                     r'<w:tab w:val="right" w:pos="9640"/>'
                     r'</w:tabs>')
    p_sub._p.get_or_add_pPr().append(tabs)
    
    r1 = p_sub.add_run("ФИО")
    set_run_ru(r1, size_pt=7.5, italic=True)
    r2 = p_sub.add_run("\tподпись")
    set_run_ru(r2, size_pt=7.5, italic=True)

def generate_deal12_act(out_path):
    doc = docx.Document()
    
    # Поля страницы А4 (выверенные для идеального 1 листа)
    for s in doc.sections:
        s.top_margin = Pt(22.0)
        s.bottom_margin = Pt(22.0)
        s.left_margin = Pt(32.0)
        s.right_margin = Pt(32.0)

    # Заголовок документа
    add_p(doc, "ПЕРЕДАТОЧНЫЙ АКТ", size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=0, space_after=1.0)
    add_p(doc, "к Договору купли-продажи квартиры от «26» октября 2026 года", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=0, space_after=2.0)

    # Дата и город
    add_header_table(doc, "город Пенза", "«29» октября 2026 года")

    # Преамбула Сторон
    preamble = (
        "Мы, нижеподписавшиеся: гражданин Орнатский Олег Александрович, 22.01.1978 года рождения, пол: мужской, "
        "паспорт гражданина Российской Федерации: серия 56 22 № 718302, выдан УМВД России по Пензенской области 27.01.2023 г., "
        "код подразделения 580-001, зарегистрированный по адресу: город Пенза, проезд 40 лет Октября 2-й, дом 3, и "
        "гражданка Орнатская Юлия Юрьевна, 27.09.1977 года рождения, пол: женский, паспорт гражданина Российской Федерации: "
        "серия 56 22 № 701399, выдан УМВД России по Пензенской области 07.10.2022 г., код подразделения 580-001, "
        "зарегистрированная по адресу: город Пенза, проезд 40 лет Октября 2-й, дом 3, именуемые в дальнейшем «Продавцы», с одной стороны, и "
        "гражданин Рыжаков Дмитрий Александрович, 09.03.1986 года рождения, пол: мужской, паспорт гражданина Российской Федерации: "
        "серия 56 05 № 588224, выдан Отделом Внутренних Дел Железнодорожного района г. Пензы 26.07.2006 г., код подразделения 582-004, "
        "зарегистрированный по адресу: город Пенза, улица Мира, дом 34, квартира 48, именуемый в дальнейшем «Покупатель», с другой стороны, "
        "совместно именуемые «Стороны», составили настоящий Передаточный акт о нижеследующем:"
    )
    add_p(doc, preamble, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 1: Передача квартиры
    p1 = (
        "1. В соответствии со статьей 556 Гражданского кодекса Российской Федерации и условиями Договора купли-продажи квартиры "
        "от «26» октября 2026 года Продавцы передали, а Покупатель принял в собственность следующее недвижимое имущество: "
        "жилое помещение — квартиру, расположенную по адресу: Пензенская область, город Пенза, улица Зеленодольская, дом 74, квартира 45, "
        "назначение: жилое, общей площадью 32,1 кв.м, расположенную на 8 этаже многоквартирного жилого дома, кадастровый номер: 58:29:3001006:1846 (далее — «Квартира»)."
    )
    add_p(doc, p1, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 2: Осмотр и качество
    p2 = (
        "2. Квартира осмотрена Покупателем до подписания настоящего акта. Покупатель подтверждает, что состояние Квартиры, "
        "конструктивных элементов, дверей, окон, остекления, сантехнического, инженерного и электрооборудования соответствует условиям "
        "Договора и пригодно для проживания. Скрытых и явных недостатков не обнаружено. Претензий по техническому состоянию у Покупателя не имеется."
    )
    add_p(doc, p2, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 3: Передаваемое имущество
    p3 = (
        "3. В соответствии с пунктом 1.3 Договора купли-продажи в Квартире передано и принято следующее имущество: "
        "встроенный кухонный гарнитур, варочная поверхность, вытяжка, встроенный холодильник, посудомоечная машина, микроволновая печь, "
        "бойлер, кондиционер, шкаф-купе в прихожей, диван. Претензий по наличию, исправности и комплектности указанного имущества не имеется."
    )
    add_p(doc, p3, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 4: Ключи и документы ЖКУ
    p4 = (
        "4. Одновременно с Квартирой Продавцы передали Покупателю полный комплект ключей от входной двери, подъезда и домофона, "
        "а также документацию и квитанции, подтверждающие отсутствие задолженности по коммунальным услугам, капремонту, электроэнергии и ТКО."
    )
    add_p(doc, p4, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 5: Показания счетчиков (построчно с удобным межстрочным пространством для записи рукой)
    add_p(doc, "5. Показания приборов учета на момент передачи: ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "электроэнергия № ___________, показ. _______ кВт*ч; ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "ХВС № __________, показ. _______ куб.м; ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "ГВС № __________, показ. _______ куб.м; ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "отопление № __________, показ. _______ Гкал.", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)

    # Пункт 6: Экземпляры (без п. 6 о взаиморасчетах)
    p6 = (
        "6. Настоящий акт составлен и подписан в двух подлинных экземплярах, имеющих равную юридическую силу: "
        "один экземпляр для Продавцов, один для Покупателя."
    )
    add_p(doc, p6, size_pt=9.5, indent=16.0, space_after=2.0, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=2.0, line_spacing=1.03)

    # Подписанты по Правилу 10 (построчные во всю ширину 17 см)
    add_signer_block_rule10(doc, "Сдал (Продавец):", "Орнатский Олег Александрович", space_before=2.0, space_after=1.0)
    add_signer_block_rule10(doc, "Сдала (Продавец):", "Орнатская Юлия Юрьевна", space_before=2.0, space_after=1.0)
    add_signer_block_rule10(doc, "Принял (Покупатель):", "Рыжаков Дмитрий Александрович", space_before=2.0, space_after=0.5)

    doc.save(out_path)
    print(f"File saved successfully to: {out_path}")

def generate_template_act(out_path):
    doc = docx.Document()
    
    for s in doc.sections:
        s.top_margin = Pt(22.0)
        s.bottom_margin = Pt(22.0)
        s.left_margin = Pt(32.0)
        s.right_margin = Pt(32.0)

    # Заголовок документа
    add_p(doc, "ПЕРЕДАТОЧНЫЙ АКТ", size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=0, space_after=1.0)
    add_p(doc, "к Договору купли-продажи квартиры от «___» ____________ 202_ года", size_pt=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_before=0, space_after=2.0)

    # Дата и город
    add_header_table(doc, "город Пенза", "«___» ____________ 202_ года")

    # Преамбула Сторон
    preamble = (
        "Мы, нижеподписавшиеся: гражданин [ФИО Продавца], [дата рождения] года рождения, пол: [пол], "
        "паспорт гражданина Российской Федерации: серия [серия] № [номер], выдан [кем выдан] [дата выдачи] г., "
        "код подразделения [код], зарегистрированный по адресу: [адрес регистрации], именуемый в дальнейшем «Продавец», с одной стороны, и "
        "гражданин [ФИО Покупателя], [дата рождения] года рождения, пол: [пол], паспорт гражданина Российской Федерации: "
        "серия [серия] № [номер], выдан [кем выдан] [дата выдачи] г., код подразделения [код], "
        "зарегистрированный по адресу: [адрес регистрации], именуемый в дальнейшем «Покупатель», с другой стороны, "
        "совместно именуемые «Стороны», составили настоящий Передаточный акт о нижеследующем:"
    )
    add_p(doc, preamble, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 1: Передача квартиры
    p1 = (
        "1. В соответствии со статьей 556 Гражданского кодекса Российской Федерации и условиями Договора купли-продажи квартиры "
        "от «___» ____________ 202_ года Продавец передал, а Покупатель принял в собственность следующее недвижимое имущество: "
        "жилое помещение — квартиру, расположенную по адресу: [Адрес объекта], "
        "назначение: жилое, общей площадью [Площадь] кв.м, расположенную на [Этаж] этаже многоквартирного жилого дома, кадастровый номер: [Кадастровый номер] (далее — «Квартира»)."
    )
    add_p(doc, p1, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 2: Осмотр и качество
    p2 = (
        "2. Квартира осмотрена Покупателем до подписания настоящего акта. Покупатель подтверждает, что состояние Квартиры, "
        "конструктивных элементов, дверей, окон, остекления, сантехнического, инженерного и электрооборудования соответствует условиям "
        "Договора и пригодно для проживания. Скрытых и явных недостатков не обнаружено. Претензий по техническому состоянию у Покупателя не имеется."
    )
    add_p(doc, p2, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 3: Передаваемое имущество
    p3 = (
        "3. В соответствии с условиями Договора купли-продажи в Квартире передано и принято следующее имущество: "
        "[Перечень передаваемой мебели, техники и предметов интерьера]. "
        "Претензий по наличию, исправности и комплектности указанного имущества не имеется."
    )
    add_p(doc, p3, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 4: Ключи и документы ЖКУ
    p4 = (
        "4. Одновременно с Квартирой Продавец передал Покупателю полный комплект ключей от входной двери, подъезда и домофона, "
        "а также документацию и квитанции, подтверждающие отсутствие задолженности по коммунальным услугам, капремонту, электроэнергии и ТКО."
    )
    add_p(doc, p4, size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)

    # Пункт 5: Показания счетчиков (построчно с интервалом под рукописное заполнение)
    add_p(doc, "5. Показания приборов учета на момент передачи: ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "электроэнергия № ___________, показ. _______ кВт*ч; ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "ХВС № __________, показ. _______ куб.м; ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "ГВС № __________, показ. _______ куб.м; ", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "отопление № __________, показ. _______ Гкал.", size_pt=9.5, indent=16.0, space_after=1.5, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=1.5, line_spacing=1.03)

    # Пункт 6: Экземпляры (без п. 6 о расчетах)
    p6 = (
        "6. Настоящий акт составлен и подписан в двух подлинных экземплярах, имеющих равную юридическую силу: "
        "один экземпляр для Продавца, один для Покупателя."
    )
    add_p(doc, p6, size_pt=9.5, indent=16.0, space_after=2.0, line_spacing=1.03)
    add_p(doc, "", size_pt=9.5, indent=0, space_after=2.0, line_spacing=1.03)

    # Подписанты по Правилу 10 (построчные во всю ширину 17 см)
    add_signer_block_rule10(doc, "Сдал (Продавец):", "[ФИО Продавца полностью]", space_before=2.0, space_after=1.0)
    add_signer_block_rule10(doc, "Принял (Покупатель):", "[ФИО Покупателя полностью]", space_before=2.0, space_after=0.5)

    doc.save(out_path)
    print(f"Template saved successfully to: {out_path}")

if __name__ == "__main__":
    # Обновляем файл сделки 12
    deal12_out = r"D:\Антигравити\legal_estate\deals\12_Орнатские_Рыжаков_Зеленодольская\готовые\2026_10_29_Передаточный_акт_Орнатские_Рыжаков.docx"
    generate_deal12_act(deal12_out)
    
    # Сохраняем в корпоративные эталонные шаблоны
    tpl_out_1 = r"D:\Антигравити\legal_estate\templates\ШАБЛОН_Передаточный_акт_Этажи_квартира.docx"
    generate_template_act(tpl_out_1)
    
    tpl_out_2 = r"D:\ШАБЛОНЫ_ДОГОВОРОВ_ЭТАЖИ\ШАБЛОН_Передаточный_акт_Этажи_квартира.docx"
    if os.path.exists(r"D:\ШАБЛОНЫ_ДОГОВОРОВ_ЭТАЖИ"):
        generate_template_act(tpl_out_2)
