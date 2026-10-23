# -*- coding: utf-8 -*-
import docx, os, sys
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.stdout.reconfigure(encoding="utf-8")

base = r"D:\Антигравити\legal_estate\deals\10_Аренда_Греченкова_Артамонов\готовые"
src_file = os.path.join(base, "Договор_аренды_Греченкова_Артамонов_УСИЛЕННЫЙ_ПРОЕКТ.docx")
doc = docx.Document(src_file)

def set_run_ru(run, font_name="Times New Roman", size_pt=10, bold=False, italic=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    if not rPr.xpath("./w:lang"):
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="ru-RU" w:bidi="ru-RU"/>'))

# 1. Update Table 0 (date of contract)
tbl0 = doc.tables[0]
p_date = tbl0.cell(0, 1).paragraphs[0]
p_date.text = ""
r_d = p_date.add_run("«26» октября 2026 года")
set_run_ru(r_d, size_pt=10.0, bold=True)

# 2. Update P9 (clause 1.4: term 26.10.2026 to 25.09.2027)
for p in doc.paragraphs:
    if "1.4. Срок аренды по настоящему Договору" in p.text:
        p.text = ""
        p.paragraph_format.first_line_indent = Pt(21.25)
        p.paragraph_format.space_after = Pt(2.0)
        p.paragraph_format.line_spacing = 1.0791666666666666
        r14 = p.add_run("1.4. Срок аренды по настоящему Договору устанавливается на 11 (Одиннадцать) месяцев: с «26» октября 2026 года по «25» сентября 2027 года включительно. Фактический допуск Арендатора в Помещение для начала коммерческой деятельности осуществляется с «26» октября 2026 года по Акту приема-передачи (Приложение № 2). Настоящий Договор в соответствии с п. 2 ст. 651 ГК РФ не подлежит государственной регистрации и вступает в силу с момента его подписания Сторонами.")
        set_run_ru(r14)
        break

# 3. Update P17, P18, P19 (clauses 3.3, 3.3.1, 3.3.2 - payment order)
for p in doc.paragraphs:
    if "3.3.1. При подписании настоящего Договора" in p.text:
        p.text = ""
        p.paragraph_format.first_line_indent = Pt(21.25)
        p.paragraph_format.space_after = Pt(2.0)
        p.paragraph_format.line_spacing = 1.0791666666666666
        r331 = p.add_run("3.3.1. При подписании настоящего Договора Арендатор оплачивает арендную плату за первый месяц аренды в размере 19 000 (Девятнадцать тысяч) рублей за период аренды с «26» октября 2026 года по «25» ноября 2026 года, а также одновременно вносит Обеспечительный платеж в размере 19 000 (Девятнадцать тысяч) рублей в соответствии с разделом 4 настоящего Договора (всего к единовременной оплате при заключении Договора: 38 000 рублей).")
        set_run_ru(r331)
    elif "3.3.2. Начиная с декабря 2026 года" in p.text:
        p.text = ""
        p.paragraph_format.first_line_indent = Pt(21.25)
        p.paragraph_format.space_after = Pt(2.0)
        p.paragraph_format.line_spacing = 1.0791666666666666
        r332 = p.add_run("3.3.2. Начиная со второго месяца аренды, оплата каждого последующего месяца аренды производится Арендатором ежемесячно в размере 19 000 (Девятнадцать тысяч) рублей в порядке 100% предоплаты не позднее 25 (двадцать пятого) числа текущего месяца за следующий подлежащий оплате месяц аренды (в частности: арендная плата за период с 26.11.2026 г. по 25.12.2026 г. оплачивается не позднее 25 ноября 2026 года; за период с 26.12.2026 г. по 25.01.2027 г. — не позднее 25 декабря 2026 года и далее в аналогичном порядке).")
        set_run_ru(r332)
    elif "4.1. В целях обеспечения надлежащего исполнения" in p.text:
        p.text = ""
        p.paragraph_format.first_line_indent = Pt(21.25)
        p.paragraph_format.space_after = Pt(2.0)
        p.paragraph_format.line_spacing = 1.0791666666666666
        r41 = p.add_run("4.1. В целях обеспечения надлежащего исполнения обязательств по сохранению Помещения, находящегося в нем имущества, своевременной оплате аренды, коммунальных услуг и возмещению возможного ущерба Арендатор вносит Арендодателю Обеспечительный платеж (ст. 381.1 ГК РФ) в размере 19 000 (Девятнадцать тысяч) рублей одновременно с оплатой первого месяца аренды при подписании настоящего Договора.")
        set_run_ru(r41)

# 4. Update Appendix 2 and 3 dates in paragraph texts
for p in doc.paragraphs:
    if "от «25» октября 2026 г." in p.text:
        p.text = p.text.replace("«25» октября 2026 г.", "«26» октября 2026 г.")
        for r in p.runs:
            set_run_ru(r, size_pt=9.5, italic=True)
    if "Договором аренды от «25» октября 2026 г." in p.text:
        p.text = p.text.replace("«25» октября 2026 г.", "«26» октября 2026 г.")
        for r in p.runs:
            set_run_ru(r, size_pt=10.0)
    if "Договора аренды от «25» октября 2026 г." in p.text:
        p.text = p.text.replace("«25» октября 2026 г.", "«26» октября 2026 г.")
        for r in p.runs:
            set_run_ru(r, size_pt=10.0)

# Also check date in Table 2 (Appendix 2 header table)
tbl2 = doc.tables[2]
p_t2 = tbl2.cell(0, 1).paragraphs[0]
p_t2.text = ""
r_t2 = p_t2.add_run("«26» октября 2026 года")
set_run_ru(r_t2, size_pt=10.0, bold=True)

# Clean soft breaks \n from Table 1 cell texts
t1 = doc.tables[1]
for cell in t1.rows[0].cells:
    for p in cell.paragraphs:
        if "\n" in p.text:
            lines = [line.strip() for line in p.text.split("\n") if line.strip()]
            p.text = lines[0] if lines else ""
            for extra_line in lines[1:]:
                p_extra = cell.add_paragraph(extra_line)
                p_extra.paragraph_format.space_after = Pt(2.0)
                p_extra.paragraph_format.line_spacing = 1.05
                for r in p_extra.runs:
                    set_run_ru(r, size_pt=9.0)

out_file = os.path.join(base, "2026_10_26_Договор_аренды_Греченкова_Артамонов_ОБЕСПЕЧИТЕЛЬНЫЙ.docx")
doc.save(out_file)
# Also update the primary project docx
doc.save(os.path.join(base, "Договор_аренды_Греченкова_Артамонов_УСИЛЕННЫЙ_ПРОЕКТ.docx"))
print("Saved updated rent contracts successfully:", out_file)
