"""
Универсальный скоростной модуль мгновенной замены текста и вставки реквизитов в DOCX без оверинжиниринга.
Использование:
  python scripts/fast_doc_edit.py --file <path> --replace "старый текст" "новый текст"
  python scripts/fast_doc_edit.py --file <path> --insert-after "маркер" "текст для вставки"
"""

import sys
import argparse
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

def set_run_ru(run, font_name="Times New Roman", size_pt=10.0, bold=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
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

def fast_replace(file_path, old_text, new_text):
    doc = docx.Document(file_path)
    count = 0
    for p in doc.paragraphs:
        if old_text in p.text:
            p.text = p.text.replace(old_text, new_text)
            for r in p.runs:
                set_run_ru(r)
            count += 1
    
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        p.text = p.text.replace(old_text, new_text)
                        for r in p.runs:
                            set_run_ru(r)
                        count += 1

    doc.save(file_path)
    print(f"Заменено вхождений: {count} в {file_path}")

def fast_insert_after(file_path, marker_text, insert_text, indent=0.35):
    doc = docx.Document(file_path)
    inserted = False
    for i, p in enumerate(doc.paragraphs):
        if marker_text in p.text:
            new_p = doc.paragraphs[i+1].insert_paragraph_before() if i + 1 < len(doc.paragraphs) else doc.add_paragraph()
            new_p.paragraph_format.line_spacing = 1.08
            new_p.paragraph_format.space_after = Pt(2.0)
            if indent > 0:
                new_p.paragraph_format.first_line_indent = Inches(indent)
            r = new_p.add_run(insert_text)
            set_run_ru(r)
            inserted = True
            break
    
    if inserted:
        doc.save(file_path)
        print(f"Текст успешно вставлен после '{marker_text}' в {file_path}")
    else:
        print(f"Маркер '{marker_text}' не найден в {file_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fast DOCX Text Tool")
    parser.add_argument("--file", required=True, help="Path to docx file")
    parser.add_argument("--replace", nargs=2, metavar=("OLD", "NEW"), help="Replace OLD with NEW")
    parser.add_argument("--insert-after", nargs=2, metavar=("MARKER", "TEXT"), help="Insert TEXT after MARKER paragraph")
    args = parser.parse_args()

    if args.replace:
        fast_replace(args.file, args.replace[0], args.replace[1])
    elif args.insert_after:
        fast_insert_after(args.file, args.insert_after[0], args.insert_after[1])
