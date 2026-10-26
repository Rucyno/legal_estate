import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

doc_path = 'deals/05_Финошина_Ветошкина_Шатрова/готовые/2026_10_19_Ветошкина_Шатрова_Основной_договор_с_распиской.docx'
doc = docx.Document(doc_path)

def set_run_ru(run, font_name="Times New Roman", size_pt=10.0, bold=False, italic=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
        rPr.append(rFonts)
    else:
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)
    lang = rPr.find(qn('w:lang'))
    if lang is None:
        lang = parse_xml(f'<w:lang {nsdecls("w")} w:val="ru-RU"/>')
        rPr.append(lang)
    else:
        lang.set(qn('w:val'), 'ru-RU')

# Find paragraph 24 (starts with 2.4.)
p24 = None
for p in doc.paragraphs:
    if p.text.strip().startswith('2.4. Одновременно с подачей'):
        p24 = p
        break

# Find paragraph 23
p23 = None
for p in doc.paragraphs:
    if p.text.strip().startswith('2.3. В соответствии с механизмом'):
        p23 = p
        break

p23.text = ""
p23.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p23.paragraph_format.line_spacing = 1.08
p23.paragraph_format.first_line_indent = Inches(0.35)
p23.paragraph_format.space_after = Pt(2.0)
r23 = p23.add_run("2.3. В соответствии с механизмом расчетов через аккредитив с трансферацией сумма, необходимая для полного погашения задолженности по Кредитному договору № 68669730 от 29.09.2023 г., списывается банком со счета покрытия по аккредитиву непосредственно в пользу ПАО Сбербанк для закрытия кредита Продавца и погашения записи об ипотеке, а оставшаяся после трансферации часть денежных средств со счета покрытия по аккредитиву перечисляется банком по следующим банковским реквизитам:")
set_run_ru(r23, size_pt=10.0)

req_lines = [
    "• Владелец счёта (получатель): Орнатский Олег Александрович;",
    "• Номер лицевого счёта: 42307 810 9 4800 0614234 (Универсальный на 5 лет);",
    "• Банк получателя: КУРСКОЕ ОТДЕЛЕНИЕ N8596 ПАО СБЕРБАНК;",
    "• БИК: 043807606;",
    "• Корреспондентский счёт: 30101 810 3 0000 0000606;",
    "• ИНН: 7707083893, КПП: 583402001;",
    "• Код подразделения банка (ТБ/ОСБ/ВСП): 54/8624/00014 (г. Пенза, ул. Кижеватова, 8)."
]

for line in req_lines:
    new_p = p24.insert_paragraph_before()
    new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    new_p.paragraph_format.line_spacing = 1.08
    new_p.paragraph_format.first_line_indent = Inches(0.5)
    new_p.paragraph_format.space_after = Pt(1.0)
    r = new_p.add_run(line)
    set_run_ru(r, size_pt=10.0)

# 2. Update signatures block before РАСПИСКА
receipt_p = None
for p in doc.paragraphs:
    if p.text.strip() == "РАСПИСКА":
        receipt_p = p
        break

sig_title_p = None
for p in doc.paragraphs:
    if p.text.strip() == "5. Подписи Сторон":
        sig_title_p = p
        break

# Collect elements to remove
remove_elems = []
curr = sig_title_p._p.getnext()
while curr is not None and curr != receipt_p._p:
    remove_elems.append(curr)
    curr = curr.getnext()

for el in remove_elems:
    el.getparent().remove(el)

# Now insert new elements before receipt_p
content_width_dxa = 10180

def add_p_before(ref_p, text, bold=False, italic=False, size_pt=10.0, align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=20, indent=0, keep=True):
    p = ref_p.insert_paragraph_before()
    p.alignment = align
    p.paragraph_format.line_spacing = 1.08
    p.paragraph_format.space_before = Pt(before / 20.0)
    p.paragraph_format.space_after = Pt(after / 20.0)
    p.paragraph_format.first_line_indent = Inches(indent)
    p.paragraph_format.keep_with_next = keep
    r = p.add_run(text)
    set_run_ru(r, size_pt=size_pt, bold=bold, italic=italic)
    return p

def add_line_before(ref_p, keep=True):
    p = ref_p.insert_paragraph_before()
    p.paragraph_format.space_before = Pt(3.0)
    p.paragraph_format.space_after = Pt(1.0)
    p.paragraph_format.keep_with_next = keep
    # Add bottom border
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>')
    p._p.get_or_add_pPr().append(pBdr)
    r = p.add_run(" ")
    set_run_ru(r, size_pt=10.0)
    return p

def add_subtext_before(ref_p, left_text="ФИО", right_text="подпись", keep=True):
    p = ref_p.insert_paragraph_before()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4.0)
    p.paragraph_format.keep_with_next = keep
    tabs = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:pos="{content_width_dxa}"/></w:tabs>')
    p._p.get_or_add_pPr().append(tabs)
    r1 = p.add_run(left_text)
    set_run_ru(r1, size_pt=8.0, italic=True)
    r2 = p.add_run()
    r2._r.append(parse_xml(f'<w:tab {nsdecls("w")}/>'))
    r3 = p.add_run(right_text)
    set_run_ru(r3, size_pt=8.0, italic=True)
    return p

# Insert signature elements:
add_p_before(receipt_p, "ПРОДАВЕЦ:", bold=True, size_pt=10.0, before=80, after=20)
add_p_before(receipt_p, "Ветошкина Ирина Борисовна (паспорт гражданина РФ: серия 0821 № 734780, выдан: УМВД России по Хабаровскому краю 27.07.2021 г., код подразделения: 270-002, зарегистрированная по адресу: Хабаровский край, г. Хабаровск, Иртышский пр-зд, д. 14, кв. 16)", size_pt=9.5, after=20)
add_p_before(receipt_p, "В лице представителя по доверенности серия 27 А А № 2656612 от 22.10.2026 г.: Финошиной Елены Александровны (паспорт гражданина РФ: серия 5623 № 770189, выдан: УМВД России по Пензенской области 02.12.2023 г., код подразделения: 580-001, зарегистрированная по адресу: Пензенская область, г. Пенза, ул. Красная/Свердлова, д. 19/55, кв. 222)", size_pt=9.5, after=20)
add_p_before(receipt_p, "Банковские реквизиты для перечисления денежных средств по сделке:", bold=True, size_pt=9.5, after=20)
add_p_before(receipt_p, "Получатель: Орнатский Олег Александрович", size_pt=9.0, after=10, indent=0.2)
add_p_before(receipt_p, "Номер лицевого счёта: 42307 810 9 4800 0614234 (Универсальный на 5 лет)", size_pt=9.0, after=10, indent=0.2)
add_p_before(receipt_p, "Банк получателя: КУРСКОЕ ОТДЕЛЕНИЕ N8596 ПАО СБЕРБАНК", size_pt=9.0, after=10, indent=0.2)
add_p_before(receipt_p, "БИК: 043807606, Корр. счёт: 30101 810 3 0000 0000606, ИНН: 7707083893, КПП: 583402001, Код подр. (ТБ/ОСБ/ВСП): 54/8624/00014 (г. Пенза, ул. Кижеватова, 8)", size_pt=9.0, after=30, indent=0.2)
add_line_before(receipt_p)
add_subtext_before(receipt_p, "Финошина Елена Александровна (по доверенности за Ветошкину И.Б.)", "подпись")

add_p_before(receipt_p, "ПОКУПАТЕЛЬ:", bold=True, size_pt=10.0, before=80, after=20)
add_p_before(receipt_p, "Шатрова Наталья Александровна (паспорт гражданина РФ: серия 5619 № 516090, выдан: УМВД России по Пензенской области 19.07.2019 г., код подразделения: 580-001, зарегистрированная по адресу: Пензенская область, г. Пенза, ул. Кронштадтская, д. 4а, кв. 28)", size_pt=9.5, after=30)
add_line_before(receipt_p)
add_subtext_before(receipt_p, "Шатрова Наталья Александровна", "подпись", keep=False)

# Ensure page break on receipt
receipt_p.paragraph_format.page_break_before = True

doc.save(doc_path)
print("Updated deal 5 DKP with exact requisites successfully!")
