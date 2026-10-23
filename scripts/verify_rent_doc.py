# -*- coding: utf-8 -*-
import docx, os, sys
sys.stdout.reconfigure(encoding="utf-8")

base = r"D:\Антигравити\legal_estate\deals\10_Аренда_Греченкова_Артамонов\готовые"
f = os.path.join(base, "2026_10_26_Договор_аренды_Греченкова_Артамонов_ОБЕСПЕЧИТЕЛЬНЫЙ.docx")
doc = docx.Document(f)

for i, p in enumerate(doc.paragraphs):
    if any(k in p.text for k in ["1.4.", "3.3.", "4.1."]):
        print(f"P{i}: {p.text}\n")
