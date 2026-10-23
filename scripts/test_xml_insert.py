# -*- coding: utf-8 -*-
import docx, os, sys
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.stdout.reconfigure(encoding="utf-8")

def set_run_ru(run, font_name="Times New Roman", size_pt=10, bold=False, italic=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    rPr = run._r.get_or_add_rPr()
    if not rPr.xpath("./w:lang"):
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="ru-RU" w:bidi="ru-RU"/>'))

base = r"D:\Антигравити\legal_estate\deals\07_Колесова\готовые"
f_dkp = os.path.join(base, "2026_10_23_ДКП_с_Распиской_Лазутина_Медведев.docx")
doc = docx.Document(f_dkp)

# We want to replace Table 1 (DKP signatures) with paragraph block:
t1 = doc.tables[1]

# Insert paragraphs before t1
p_sec = t1._tbl.getparent().insert(t1._tbl.getparent().index(t1._tbl), parse_xml(f"""
<w:p {nsdecls("w")}>
    <w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:sz w:val="20"/>
            <w:lang w:val="ru-RU"/>
        </w:rPr>
        <w:keepNext/>
    </w:pPr>
    <w:r>
        <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:b/>
            <w:sz w:val="20"/>
            <w:lang w:val="ru-RU"/>
        </w:rPr>
        <w:t>Продавец:</w:t>
    </w:r>
</w:p>
"""))

print("Inserted successfully")
