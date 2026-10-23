import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

target_dir = r"D:\Антигравити\legal_estate\deals\06_Салтыкова_Лунев_Вострова_Анохин\готовые"
src_dkp = os.path.join(target_dir, "2026_10_23_ДКП_БАНК_ДОМ_РФ_Лунев_Вострова_Анохин.docx")
out_highlighted = os.path.join(target_dir, "2026_10_23_ДКП_БАНК_ДОМ_РФ_с_пунктами.docx")

doc = docx.Document(src_dkp)

def highlight_paragraph(p, color='yellow'):
    for run in p.runs:
        rPr = run._r.get_or_add_rPr()
        highlight = parse_xml(f'<w:highlight {nsdecls("w")} w:val="{color}"/>')
        rPr.append(highlight)

highlighted_count = 0
for p in doc.paragraphs:
    t = p.text.strip()
    should_highlight = False
    
    if (t.startswith('1.4.') or t.startswith('2.2.') or t.startswith('2.3.') or 
        t.startswith('2.4.') or t.startswith('2.5.') or t.startswith('2.7.') or 
        t.startswith('2.8.') or t.startswith('3.2.')):
        should_highlight = True
    elif t.startswith('- Денежные средства в размере 50 000') or t.startswith('- Оставшаяся сумма в размере 2 350 000'):
        should_highlight = True
    elif (t.startswith('- Сумма аккредитива') or t.startswith('- Срок действия аккредитива') or 
          t.startswith('- Банк-эмитент:') or t.startswith('- Исполняющий банк:') or 
          t.startswith('- Плательщик по аккредитиву:') or t.startswith('- Порядок распределения и списания')):
        should_highlight = True
    elif 'а) Часть денежных средств аккредитива' in t or 'б) Оставшаяся после полного погашения' in t:
        should_highlight = True
    elif ('• Банк получателя:' in t or '• Получатель средств:' in t or '• Номер счета получателя:' in t or 
          '• БИК банка получателя:' in t or '• Корр. счет банка' in t or '• ИНН банка получателя:' in t or 
          '• Назначение платежа:' in t):
        should_highlight = True
        
    if should_highlight:
        highlight_paragraph(p, 'yellow')
        highlighted_count += 1
        print(f"Highlighted: {t[:60]}...")

doc.save(out_highlighted)
print(f"Total highlighted paragraphs: {highlighted_count}")
print(f"Successfully saved: {out_highlighted}")
